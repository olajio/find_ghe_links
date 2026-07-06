#!/usr/bin/env python3
"""
SharePoint -> GHE link scraper.

Crawls the configured SharePoint Online site (pages + documents) and reports
every occurrence of a target host (default: ghe.hedgeserv.net), whether it
appears as raw text or is hidden behind hyperlink display text.

See README.md for setup and usage.
"""

import argparse
import configparser
import io
import json
import logging
import os
import re
import sys
import time
import zipfile
from dataclasses import dataclass, asdict
from typing import Iterable, Iterator, Optional
from xml.etree import ElementTree as ET

import requests

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
DEFAULT_SCOPES = ["Sites.Read.All", "Files.Read.All"]
TOKEN_CACHE_FILE = ".token_cache.bin"

# OOXML namespaces used when parsing .docx internals.
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

SUPPORTED_DOC_EXTS = {".docx", ".xlsx", ".pdf", ".doc", ".xls", ".ppt"}

log = logging.getLogger("ghe_scraper")


# --------------------------------------------------------------------------- #
# Match model
# --------------------------------------------------------------------------- #
@dataclass
class Match:
    site: str             # site display name / URL the match was found in
    source_type: str      # "page" | "document"
    title: str            # page title or file name
    location_url: str     # link to the page / document in SharePoint
    match_type: str       # "hyperlink" | "raw-text" | "webpart-data" | "binary"
    display_text: str     # visible text the link hid behind (if any)
    target_url: str       # the actual ghe.hedgeserv.net URL


# --------------------------------------------------------------------------- #
# Matching helpers
# --------------------------------------------------------------------------- #
def build_patterns(needle: str):
    """Return (bare_re, contains) for the configured needle."""
    escaped = re.escape(needle)
    # Grabs the whole URL token that contains the needle, with or without a
    # scheme, stopping at whitespace / quotes / brackets.
    bare_re = re.compile(
        r"[^\s\"'<>()\[\]]*" + escaped + r"[^\s\"'<>()\[\]]*",
        re.IGNORECASE,
    )
    return bare_re


def scan_html(html: str, needle: str, bare_re: re.Pattern) -> list[tuple[str, str, str]]:
    """
    Parse an HTML fragment and return a list of
    (match_type, display_text, target_url) tuples.

    Catches both <a href> hyperlinks (the "hidden behind display text" case)
    and raw-text mentions, deduping raw text already covered by a hyperlink.
    """
    from bs4 import BeautifulSoup

    results: list[tuple[str, str, str]] = []
    needle_l = needle.lower()
    soup = BeautifulSoup(html or "", "html.parser")

    hyperlink_targets: set[str] = set()
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if needle_l in href.lower():
            display = a.get_text(strip=True)
            results.append(("hyperlink", display, href))
            hyperlink_targets.add(href)

    # Any other element attribute that carries the host (iframe src, etc.).
    for tag in soup.find_all(True):
        for attr, val in tag.attrs.items():
            if attr == "href":
                continue
            values = val if isinstance(val, list) else [val]
            for v in values:
                if isinstance(v, str) and needle_l in v.lower():
                    for m in bare_re.findall(v):
                        if m not in hyperlink_targets:
                            results.append(("hyperlink", "", m))
                            hyperlink_targets.add(m)

    # Raw-text mentions in the visible text.
    text = soup.get_text(" ")
    for m in bare_re.findall(text):
        if any(m in t for t in hyperlink_targets):
            continue
        results.append(("raw-text", "", m))

    return results


def scan_text(text: str, bare_re: re.Pattern) -> list[tuple[str, str, str]]:
    """Return raw-text matches from a plain-text blob."""
    return [("raw-text", "", m) for m in bare_re.findall(text or "")]


# --------------------------------------------------------------------------- #
# Auth
# --------------------------------------------------------------------------- #
def acquire_token(client_id: str, authority_tenant: str, scopes: list[str]) -> str:
    """Interactive device-code login; caches the token for reuse."""
    import msal

    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_FILE):
        try:
            cache.deserialize(open(TOKEN_CACHE_FILE, "r").read())
        except Exception:
            log.warning("Could not read token cache; starting fresh.")

    app = msal.PublicClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{authority_tenant}",
        token_cache=cache,
    )

    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            raise RuntimeError(
                "Failed to start device flow: "
                + json.dumps(flow, indent=2)
            )
        print("\n" + "=" * 70)
        print(flow["message"])  # "To sign in, use a web browser to open ..."
        print("=" * 70 + "\n")
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise RuntimeError(
            "Authentication failed: "
            + result.get("error_description", json.dumps(result))
        )

    if cache.has_state_changed:
        with open(TOKEN_CACHE_FILE, "w") as fh:
            fh.write(cache.serialize())

    return result["access_token"]


# --------------------------------------------------------------------------- #
# Graph HTTP helpers (with throttling handling)
# --------------------------------------------------------------------------- #
class Graph:
    def __init__(self, token: str):
        self.session = requests.Session()
        self.session.headers.update({"Authorization": f"Bearer {token}"})

    def get(self, url: str, *, stream: bool = False, raw: bool = False) -> requests.Response:
        """GET with retry/backoff on 429 and 5xx, honoring Retry-After."""
        if url.startswith("/"):
            url = GRAPH_ROOT + url
        attempt = 0
        while True:
            attempt += 1
            resp = self.session.get(url, stream=stream)
            if resp.status_code in (429, 503, 504) and attempt <= 6:
                wait = int(resp.headers.get("Retry-After", 0)) or min(2 ** attempt, 60)
                log.warning("Throttled (%s). Waiting %ss …", resp.status_code, wait)
                time.sleep(wait)
                continue
            resp.raise_for_status()
            return resp

    def get_json(self, url: str) -> dict:
        return self.get(url).json()

    def paged(self, url: str) -> Iterator[dict]:
        """Yield items from a paged Graph collection, following @odata.nextLink."""
        while url:
            data = self.get_json(url)
            for item in data.get("value", []):
                yield item
            url = data.get("@odata.nextLink")


# --------------------------------------------------------------------------- #
# Site + page crawling
# --------------------------------------------------------------------------- #
def resolve_site_id(graph: Graph, hostname: str, site_path: str) -> tuple[str, str]:
    """Return (site_id, web_url) for hostname:site_path."""
    path = site_path.strip("/")
    data = graph.get_json(f"/sites/{hostname}:/{path}")
    return data["id"], data.get("webUrl", "")


def iter_webparts(canvas: dict) -> Iterator[dict]:
    """Yield every web part in a sitePage canvasLayout."""
    if not canvas:
        return
    sections = list(canvas.get("horizontalSections", []))
    vertical = canvas.get("verticalSection")
    if vertical:
        sections.append(vertical)
    for section in sections:
        for column in section.get("columns", [section]):
            for wp in column.get("webparts", column.get("webParts", [])):
                yield wp
    # Some payloads expose a flat webParts array on the page itself.
    for wp in canvas.get("webParts", []):
        yield wp


def crawl_pages(
    graph: Graph, site_id: str, site_name: str, needle: str, bare_re: re.Pattern
) -> list[Match]:
    matches: list[Match] = []
    url = (
        f"/sites/{site_id}/pages/microsoft.graph.sitePage"
        "?$expand=canvasLayout&$top=50"
    )
    pages = list(graph.paged(url))
    log.info("Found %d page(s).", len(pages))

    for page in pages:
        title = page.get("title") or page.get("name") or "(untitled)"
        web_url = page.get("webUrl", "")
        canvas = page.get("canvasLayout")

        # Some tenants don't populate canvasLayout on the list response; refetch.
        if canvas is None and page.get("id"):
            try:
                detail = graph.get_json(
                    f"/sites/{site_id}/pages/{page['id']}"
                    "/microsoft.graph.sitePage?$expand=canvasLayout"
                )
                canvas = detail.get("canvasLayout")
            except requests.HTTPError as exc:
                log.warning("Could not expand page '%s': %s", title, exc)

        page_hits = 0
        for wp in iter_webparts(canvas or {}):
            inner_html = wp.get("innerHtml") or wp.get("innerHTML")
            if inner_html:
                for mt, disp, target in scan_html(inner_html, needle, bare_re):
                    matches.append(Match(site_name, "page", title, web_url, mt, disp, target))
                    page_hits += 1
            # Link web parts (Quick Links, Hero, …) keep URLs in JSON data, not
            # HTML — scan the serialized web part to catch those targets.
            data_blob = json.dumps(wp, ensure_ascii=False)
            if needle.lower() in data_blob.lower():
                for m in set(bare_re.findall(data_blob)):
                    matches.append(Match(site_name, "page", title, web_url, "webpart-data", "", m))
                    page_hits += 1

        if page_hits:
            log.info("  %-50s %d hit(s)", title[:50], page_hits)

    return matches


# --------------------------------------------------------------------------- #
# Document crawling
# --------------------------------------------------------------------------- #
def iter_drive_files(graph: Graph, drive_id: str) -> Iterator[dict]:
    """Recursively yield file items (driveItems with a 'file' facet) in a drive."""
    stack = [f"/drives/{drive_id}/root/children"]
    while stack:
        url = stack.pop()
        for item in graph.paged(url):
            if item.get("folder"):
                stack.append(f"/drives/{drive_id}/items/{item['id']}/children")
            elif item.get("file"):
                yield item


def download_item(graph: Graph, item: dict) -> Optional[bytes]:
    """Download a driveItem's content, preferring the pre-authenticated URL."""
    dl = item.get("@microsoft.graph.downloadUrl")
    if dl:
        resp = requests.get(dl)  # pre-signed; no auth header needed
        resp.raise_for_status()
        return resp.content
    if item.get("parentReference", {}).get("driveId"):
        drive_id = item["parentReference"]["driveId"]
        resp = graph.get(f"/drives/{drive_id}/items/{item['id']}/content", stream=True)
        return resp.content
    return None


def parse_docx(data: bytes, needle: str, bare_re: re.Pattern) -> list[tuple[str, str, str]]:
    results: list[tuple[str, str, str]] = []
    needle_l = needle.lower()
    zf = zipfile.ZipFile(io.BytesIO(data))

    # Map relationship id -> external target.
    rel_targets: dict[str, str] = {}
    try:
        rels = ET.fromstring(zf.read("word/_rels/document.xml.rels"))
        for rel in rels:
            rid, target = rel.get("Id"), rel.get("Target")
            if rid and target:
                rel_targets[rid] = target
    except KeyError:
        pass

    matched_targets: set[str] = set()
    try:
        doc = ET.fromstring(zf.read("word/document.xml"))
    except KeyError:
        doc = None

    if doc is not None:
        for hlink in doc.iter(f"{{{W_NS}}}hyperlink"):
            rid = hlink.get(f"{{{R_NS}}}id")
            target = rel_targets.get(rid, "")
            if target and needle_l in target.lower():
                display = "".join(t.text or "" for t in hlink.iter(f"{{{W_NS}}}t"))
                results.append(("hyperlink", display.strip(), target))
                matched_targets.add(target)
        full_text = "".join(t.text or "" for t in doc.iter(f"{{{W_NS}}}t"))
        for m in bare_re.findall(full_text):
            if not any(m in t for t in matched_targets):
                results.append(("raw-text", "", m))

    # Hyperlink rels not tied to a <w:hyperlink> element (rare, but be thorough).
    for target in rel_targets.values():
        if needle_l in target.lower() and target not in matched_targets:
            results.append(("hyperlink", "", target))
            matched_targets.add(target)

    return results


def parse_xlsx(data: bytes, needle: str, bare_re: re.Pattern) -> list[tuple[str, str, str]]:
    import openpyxl

    results: list[tuple[str, str, str]] = []
    needle_l = needle.lower()
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=False)
    try:
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    link = getattr(cell, "hyperlink", None)
                    if link and link.target and needle_l in link.target.lower():
                        results.append(("hyperlink", str(cell.value or ""), link.target))
                    if isinstance(cell.value, str) and needle_l in cell.value.lower():
                        for m in bare_re.findall(cell.value):
                            results.append(("raw-text", "", m))
    finally:
        wb.close()
    return results


def parse_pdf(data: bytes, needle: str, bare_re: re.Pattern) -> list[tuple[str, str, str]]:
    from pypdf import PdfReader

    results: list[tuple[str, str, str]] = []
    needle_l = needle.lower()
    reader = PdfReader(io.BytesIO(data))
    for page in reader.pages:
        try:
            text = page.extract_text() or ""
        except Exception:
            text = ""
        for m in bare_re.findall(text):
            results.append(("raw-text", "", m))

        annots = page.get("/Annots")
        if not annots:
            continue
        for ref in annots:
            try:
                obj = ref.get_object()
                action = obj.get("/A") or {}
                uri = action.get("/URI")
                if uri and needle_l in str(uri).lower():
                    results.append(("hyperlink", "", str(uri)))
            except Exception:
                continue
    return results


def parse_ole_raw(data: bytes, needle: str, bare_re: re.Pattern) -> list[tuple[str, str, str]]:
    """
    Fallback for legacy OLE2 binary formats (.doc, .ppt, and .xls when a
    dedicated reader is unavailable).

    These formats have no simple text/hyperlink API in pure Python, so we scan
    the raw bytes for the needle in both ASCII and UTF-16-LE (Office stores most
    text, including HYPERLINK field targets, in one of those). We can recover the
    URL but not reliably associate display text, so matches are typed "binary".
    """
    # Decoded binary windows have no clean delimiters, so use a strict
    # URL-legal character class (excludes CJK noise from adjacent bytes).
    strict = re.compile(
        r"[A-Za-z0-9._~:/?#@%&=+\-]*" + re.escape(needle) + r"[A-Za-z0-9._~:/?#@%&=+\-]*",
        re.IGNORECASE,
    )
    results: list[tuple[str, str, str]] = []
    seen: set[str] = set()
    for enc in ("ascii", "utf-16-le"):
        try:
            token = needle.encode(enc)
        except UnicodeEncodeError:
            continue
        start = 0
        while True:
            i = data.find(token, start)
            if i == -1:
                break
            lo = max(0, i - 512)
            # UTF-16-LE is 2 bytes/char; keep the window start on the same byte
            # parity as the match so decoding stays aligned to char boundaries.
            if enc == "utf-16-le" and (i - lo) % 2:
                lo += 1
            window = data[lo: i + 512].decode(enc, errors="ignore")
            for raw in strict.findall(window):
                # If a scheme is present, trim any leading junk before it.
                k = raw.lower().find("http")
                m = (raw[k:] if k > 0 else raw).strip(".")
                if m and m not in seen:
                    seen.add(m)
                    results.append(("binary", "", m))
            start = i + len(token)
    return results


def parse_xls(data: bytes, needle: str, bare_re: re.Pattern) -> list[tuple[str, str, str]]:
    """Legacy .xls via xlrd (cell text + hyperlinks); falls back to raw scan."""
    try:
        import xlrd
    except ImportError:
        return parse_ole_raw(data, needle, bare_re)

    needle_l = needle.lower()
    try:
        # formatting_info=True is what exposes hyperlink_map for .xls files.
        book = xlrd.open_workbook(file_contents=data, formatting_info=True)
    except Exception:
        return parse_ole_raw(data, needle, bare_re)

    results: list[tuple[str, str, str]] = []
    for sheet in book.sheets():
        for (row, col), link in (getattr(sheet, "hyperlink_map", {}) or {}).items():
            url = getattr(link, "url_or_path", "") or ""
            if url and needle_l in url.lower():
                try:
                    disp = str(sheet.cell_value(row, col))
                except Exception:
                    disp = ""
                results.append(("hyperlink", disp, url))
        for row in range(sheet.nrows):
            for col in range(sheet.ncols):
                val = sheet.cell_value(row, col)
                if isinstance(val, str) and needle_l in val.lower():
                    for m in bare_re.findall(val):
                        results.append(("raw-text", "", m))
    return results


DOC_PARSERS = {
    ".docx": parse_docx,
    ".xlsx": parse_xlsx,
    ".pdf": parse_pdf,
    ".xls": parse_xls,
    ".doc": parse_ole_raw,
    ".ppt": parse_ole_raw,
}


def crawl_documents(
    graph: Graph, site_id: str, site_name: str, needle: str, bare_re: re.Pattern,
    max_file_mb: int,
) -> list[Match]:
    matches: list[Match] = []
    drives = list(graph.paged(f"/sites/{site_id}/drives"))
    log.info("Found %d document librar%s.", len(drives), "y" if len(drives) == 1 else "ies")

    max_bytes = max_file_mb * 1024 * 1024
    for drive in drives:
        log.info("Scanning library: %s", drive.get("name", drive["id"]))
        for item in iter_drive_files(graph, drive["id"]):
            name = item.get("name", "")
            ext = os.path.splitext(name)[1].lower()
            parser = DOC_PARSERS.get(ext)
            if not parser:
                continue
            if item.get("size", 0) > max_bytes:
                log.info("  skip (too large): %s", name)
                continue
            web_url = item.get("webUrl", "")
            try:
                content = download_item(graph, item)
                if content is None:
                    continue
                hits = parser(content, needle, bare_re)
            except Exception as exc:
                log.warning("  could not parse %s: %s", name, exc)
                continue
            for mt, disp, target in hits:
                matches.append(Match(site_name, "document", name, web_url, mt, disp, target))
            if hits:
                log.info("  %-50s %d hit(s)", name[:50], len(hits))

    return matches


# --------------------------------------------------------------------------- #
# Reporting
# --------------------------------------------------------------------------- #
def dedupe(matches: list[Match]) -> list[Match]:
    seen = set()
    unique = []
    for m in matches:
        key = (m.site, m.source_type, m.location_url, m.match_type, m.target_url, m.display_text)
        if key not in seen:
            seen.add(key)
            unique.append(m)
    return unique


def write_report(matches: list[Match], csv_path: str, xlsx_path: str) -> None:
    import pandas as pd

    rows = [asdict(m) for m in matches]
    df = pd.DataFrame(
        rows,
        columns=[
            "site",
            "source_type",
            "title",
            "location_url",
            "match_type",
            "display_text",
            "target_url",
        ],
    )
    df.to_csv(csv_path, index=False)
    try:
        df.to_excel(xlsx_path, index=False)
    except Exception as exc:
        log.warning("Could not write Excel report (%s); CSV was written.", exc)

    print("\n" + "-" * 70)
    print(f"Total occurrences: {len(matches)}")
    if not df.empty:
        print("\nBy source type:")
        print(df["source_type"].value_counts().to_string())
        print("\nBy match type:")
        print(df["match_type"].value_counts().to_string())
    print(f"\nReport written to:\n  {csv_path}\n  {xlsx_path}")
    print("-" * 70)


def discover_sites(
    graph: Graph, root_id: str, root_web_url: str, recursive: bool
) -> list[dict]:
    """
    Return [{id, webUrl, name}] for the root site, plus all nested subsites
    when recursive is True (walks /sites/{id}/sites depth-first).
    """
    root_name = root_web_url.rstrip("/").split("/")[-1] or root_web_url or root_id
    sites = [{"id": root_id, "webUrl": root_web_url, "name": root_name}]
    if not recursive:
        return sites

    stack = [root_id]
    seen = {root_id}
    while stack:
        sid = stack.pop()
        try:
            for sub in graph.paged(f"/sites/{sid}/sites"):
                sub_id = sub.get("id")
                if not sub_id or sub_id in seen:
                    continue
                seen.add(sub_id)
                sites.append({
                    "id": sub_id,
                    "webUrl": sub.get("webUrl", ""),
                    "name": sub.get("displayName") or sub.get("name") or sub.get("webUrl", ""),
                })
                stack.append(sub_id)
        except requests.HTTPError as exc:
            log.warning("Could not list subsites of %s: %s", sid, exc)

    return sites


# --------------------------------------------------------------------------- #
# Config + main
# --------------------------------------------------------------------------- #
def load_config(path: str) -> configparser.ConfigParser:
    cfg = configparser.ConfigParser()
    if os.path.exists(path):
        cfg.read(path)
    elif os.path.exists("config.example.ini"):
        log.warning("%s not found; falling back to config.example.ini", path)
        cfg.read("config.example.ini")
    else:
        raise FileNotFoundError(f"No config file found at {path}")
    return cfg


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="SharePoint -> GHE link scraper")
    parser.add_argument("-c", "--config", default="config.ini", help="Path to config file")
    parser.add_argument("--no-documents", action="store_true", help="Scan pages only")
    parser.add_argument(
        "--recursive", action="store_true",
        help="Also crawl all subsites of the target site (e.g. point site_path at "
             "the parent GlobalTechnology site to sweep everything under it).",
    )
    parser.add_argument("--site-path", help="Override site_path from the config file")
    parser.add_argument("--verbose", action="store_true", help="Debug logging")
    args = parser.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)-7s %(message)s",
        datefmt="%H:%M:%S",
    )

    cfg = load_config(args.config)
    hostname = cfg.get("sharepoint", "hostname")
    site_path = args.site_path or cfg.get("sharepoint", "site_path")
    recursive = args.recursive or cfg.getboolean("sharepoint", "recurse_subsites", fallback=False)
    needle = cfg.get("search", "needle", fallback="ghe.hedgeserv.net")
    client_id = cfg.get("auth", "client_id")
    authority_tenant = cfg.get("auth", "authority_tenant", fallback="organizations")
    docs_enabled = cfg.getboolean("documents", "enabled", fallback=True) and not args.no_documents
    max_file_mb = cfg.getint("documents", "max_file_mb", fallback=25)
    csv_path = cfg.get("output", "csv_path", fallback="ghe_links_report.csv")
    xlsx_path = cfg.get("output", "xlsx_path", fallback="ghe_links_report.xlsx")

    bare_re = build_patterns(needle)

    log.info("Authenticating (device-code) …")
    token = acquire_token(client_id, authority_tenant, DEFAULT_SCOPES)
    graph = Graph(token)

    log.info("Resolving site %s:%s …", hostname, site_path)
    site_id, site_web_url = resolve_site_id(graph, hostname, site_path)
    log.info("Site resolved: %s", site_web_url or site_id)

    sites = discover_sites(graph, site_id, site_web_url, recursive)
    if recursive:
        log.info("Recursive mode: crawling %d site(s) (root + subsites).", len(sites))

    matches: list[Match] = []
    for site in sites:
        if len(sites) > 1:
            log.info("== Site: %s ==", site["webUrl"] or site["name"])
        matches += crawl_pages(graph, site["id"], site["name"], needle, bare_re)
        if docs_enabled:
            matches += crawl_documents(
                graph, site["id"], site["name"], needle, bare_re, max_file_mb
            )
    if not docs_enabled:
        log.info("Document scan disabled.")

    matches = dedupe(matches)
    write_report(matches, csv_path, xlsx_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())

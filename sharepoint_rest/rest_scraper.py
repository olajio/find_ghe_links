#!/usr/bin/env python3
"""
SharePoint -> GHE link scraper (SharePoint REST API edition).

Same goal as the Graph edition in the parent folder, but authenticates against
the SharePoint REST API with the delegated **AllSites.Read** scope instead of
Microsoft Graph's admin-restricted Sites.Read.All. AllSites.Read is scoped to
"read what the signed-in user can already read in SharePoint," and in many
tenants it is user-consentable without an admin approval prompt.

Finds every occurrence of a target host (default: ghe.hedgeserv.net) across the
site's pages and documents — whether it appears as raw text or is hidden behind
hyperlink display text — and always captures the FULL URL, including the URI
path after the domain that points to the actual resource.

See README.md for setup and usage.
"""

import argparse
import configparser
import logging
import os
import re
import sys
import time

import requests

from matchers import (
    DOC_PARSERS,
    Match,
    build_patterns,
    dedupe,
    scan_html,
    write_report,
)

TOKEN_CACHE_FILE = ".sp_token_cache.bin"

log = logging.getLogger("ghe_scraper")


# --------------------------------------------------------------------------- #
# Auth (device-code, SharePoint resource)
# --------------------------------------------------------------------------- #
def acquire_token(client_id: str, authority_tenant: str, scopes: list[str]) -> str:
    """Interactive device-code login for the SharePoint resource; caches token."""
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
            import json
            raise RuntimeError("Failed to start device flow: " + json.dumps(flow, indent=2))
        print("\n" + "=" * 70)
        print(flow["message"])
        print("=" * 70 + "\n")
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        import json
        raise RuntimeError(
            "Authentication failed: "
            + result.get("error_description", json.dumps(result))
        )

    if cache.has_state_changed:
        with open(TOKEN_CACHE_FILE, "w") as fh:
            fh.write(cache.serialize())

    return result["access_token"]


# --------------------------------------------------------------------------- #
# SharePoint REST client
# --------------------------------------------------------------------------- #
class SharePoint:
    def __init__(self, token: str, hostname: str):
        self.hostname_url = f"https://{hostname}"
        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": f"Bearer {token}",
            "Accept": "application/json;odata=verbose",
        })

    def get(self, url: str, *, stream: bool = False) -> requests.Response:
        """GET with retry/backoff on 429 and 5xx, honoring Retry-After."""
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

    def entity(self, url: str) -> dict:
        """Return the 'd' object of a single-entity REST response."""
        return self.get(url).json().get("d", {})

    def collection(self, url: str):
        """Yield items from a REST collection, following server-driven paging."""
        while url:
            data = self.get(url).json().get("d", {})
            for item in data.get("results", []):
                yield item
            url = data.get("__next")

    def absolute(self, server_relative_url: str) -> str:
        return self.hostname_url + server_relative_url

    def download(self, server_relative_url: str) -> bytes:
        # GetFileByServerRelativePath(decodedurl=...) handles spaces/specials
        # better than GetFileByServerRelativeUrl; single quotes are doubled.
        literal = server_relative_url.replace("'", "''")
        url = (
            f"{self.hostname_url}/_api/web"
            f"/GetFileByServerRelativePath(decodedurl='{literal}')/$value"
        )
        return self.get(url, stream=True).content


# --------------------------------------------------------------------------- #
# Web (site) discovery
# --------------------------------------------------------------------------- #
def resolve_web(sp: SharePoint, site_path: str) -> dict:
    site_abs = sp.hostname_url + "/" + site_path.strip("/")
    web = sp.entity(f"{site_abs}/_api/web?$select=Title,Url,ServerRelativeUrl")
    if not web:
        raise RuntimeError(f"Could not resolve site at {site_abs}")
    return web


def discover_webs(sp: SharePoint, root_web: dict, recursive: bool) -> list[dict]:
    """Return the root web plus all nested subwebs when recursive."""
    webs = [root_web]
    if not recursive:
        return webs

    stack = [root_web["Url"]]
    seen = {root_web["Url"].rstrip("/")}
    while stack:
        base = stack.pop()
        try:
            for sub in sp.collection(f"{base}/_api/web/webs?$select=Title,Url,ServerRelativeUrl"):
                key = sub["Url"].rstrip("/")
                if key in seen:
                    continue
                seen.add(key)
                webs.append(sub)
                stack.append(sub["Url"])
        except requests.HTTPError as exc:
            log.warning("Could not list subwebs of %s: %s", base, exc)
    return webs


# --------------------------------------------------------------------------- #
# Page crawling (Site Pages library: CanvasContent1 + WikiField)
# --------------------------------------------------------------------------- #
def pages_library_title(sp: SharePoint, web_abs: str) -> str:
    """Find the Site Pages library title (BaseTemplate 119), locale-safe."""
    try:
        libs = list(sp.collection(
            f"{web_abs}/_api/web/lists?$filter=BaseTemplate eq 119&$select=Title"
        ))
        if libs:
            return libs[0]["Title"]
    except requests.HTTPError:
        pass
    return "Site Pages"


def crawl_pages(sp: SharePoint, web: dict, needle: str, bare_re: re.Pattern) -> list[Match]:
    web_abs = web["Url"]
    site_name = web.get("Title") or web_abs.rstrip("/").split("/")[-1]
    title = pages_library_title(sp, web_abs)
    esc = title.replace("'", "''")
    url = (
        f"{web_abs}/_api/web/lists/getbytitle('{esc}')/items"
        "?$select=Title,FileRef,FileLeafRef,CanvasContent1,WikiField,FileSystemObjectType"
        "&$top=500"
    )
    try:
        items = list(sp.collection(url))
    except requests.HTTPError as exc:
        log.warning("Could not read pages for %s: %s", web_abs, exc)
        return []

    log.info("  %d page item(s) in '%s'.", len(items), title)
    matches: list[Match] = []
    for it in items:
        if it.get("FileSystemObjectType") != 0:  # 0 = file, 1 = folder
            continue
        page_title = it.get("Title") or it.get("FileLeafRef") or "(untitled)"
        page_url = sp.absolute(it.get("FileRef", ""))
        page_hits = 0
        for html in (it.get("CanvasContent1"), it.get("WikiField")):
            if not html:
                continue
            found: set[str] = set()
            for mt, disp, target in scan_html(html, needle, bare_re):
                matches.append(Match(site_name, "page", page_title, page_url, mt, disp, target))
                found.add(target)
                page_hits += 1
            # Web-part JSON (Quick Links / Hero) that HTML parsing didn't surface.
            if needle.lower() in html.lower():
                for m in set(bare_re.findall(html)):
                    if m not in found and not any(m in t for t in found):
                        matches.append(Match(site_name, "page", page_title, page_url,
                                             "webpart-data", "", m))
                        found.add(m)
                        page_hits += 1
        if page_hits:
            log.info("    %-48s %d hit(s)", page_title[:48], page_hits)
    return matches


# --------------------------------------------------------------------------- #
# Document crawling (document libraries: BaseTemplate 101)
# --------------------------------------------------------------------------- #
def crawl_documents(
    sp: SharePoint, web: dict, needle: str, bare_re: re.Pattern, max_file_mb: int
) -> list[Match]:
    web_abs = web["Url"]
    site_name = web.get("Title") or web_abs.rstrip("/").split("/")[-1]
    try:
        libs = list(sp.collection(
            f"{web_abs}/_api/web/lists?$filter=BaseTemplate eq 101 and Hidden eq false"
            "&$select=Title"
        ))
    except requests.HTTPError as exc:
        log.warning("Could not list document libraries for %s: %s", web_abs, exc)
        return []

    max_bytes = max_file_mb * 1024 * 1024
    matches: list[Match] = []
    for lib in libs:
        lib_title = lib["Title"]
        esc = lib_title.replace("'", "''")
        log.info("  Scanning library: %s", lib_title)
        url = (
            f"{web_abs}/_api/web/lists/getbytitle('{esc}')/items"
            "?$select=FileRef,FileLeafRef,FileSystemObjectType,File_x0020_Size&$top=500"
        )
        try:
            items = sp.collection(url)
            for it in items:
                if it.get("FileSystemObjectType") != 0:
                    continue
                name = it.get("FileLeafRef", "")
                ext = os.path.splitext(name)[1].lower()
                parser = DOC_PARSERS.get(ext)
                if not parser:
                    continue
                try:
                    size = int(it.get("File_x0020_Size") or 0)
                except (TypeError, ValueError):
                    size = 0
                if size > max_bytes:
                    log.info("    skip (too large): %s", name)
                    continue
                file_ref = it.get("FileRef", "")
                file_url = sp.absolute(file_ref)
                try:
                    content = sp.download(file_ref)
                    hits = parser(content, needle, bare_re)
                except Exception as exc:
                    log.warning("    could not parse %s: %s", name, exc)
                    continue
                for mt, disp, target in hits:
                    matches.append(Match(site_name, "document", name, file_url, mt, disp, target))
                if hits:
                    log.info("    %-48s %d hit(s)", name[:48], len(hits))
        except requests.HTTPError as exc:
            log.warning("  could not read library %s: %s", lib_title, exc)
    return matches


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
    parser = argparse.ArgumentParser(description="SharePoint -> GHE link scraper (REST edition)")
    parser.add_argument("-c", "--config", default="config.ini", help="Path to config file")
    parser.add_argument("--no-documents", action="store_true", help="Scan pages only")
    parser.add_argument(
        "--recursive", action="store_true",
        help="Also crawl all subsites of the target site.",
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
    scopes = [f"https://{hostname}/AllSites.Read"]

    log.info("Authenticating (device-code, SharePoint AllSites.Read) …")
    token = acquire_token(client_id, authority_tenant, scopes)
    sp = SharePoint(token, hostname)

    log.info("Resolving site %s%s …", hostname, site_path)
    root_web = resolve_web(sp, site_path)
    log.info("Site resolved: %s", root_web.get("Url"))

    webs = discover_webs(sp, root_web, recursive)
    if recursive:
        log.info("Recursive mode: crawling %d site(s) (root + subsites).", len(webs))

    matches: list[Match] = []
    for web in webs:
        if len(webs) > 1:
            log.info("== Site: %s ==", web.get("Url"))
        matches += crawl_pages(sp, web, needle, bare_re)
        if docs_enabled:
            matches += crawl_documents(sp, web, needle, bare_re, max_file_mb)
    if not docs_enabled:
        log.info("Document scan disabled.")

    matches = dedupe(matches)
    write_report(matches, csv_path, xlsx_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())

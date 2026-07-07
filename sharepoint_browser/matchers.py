"""
Format-parsing and link-matching helpers, independent of transport.

Shared by the SharePoint REST scraper. Everything here operates on in-memory
strings/bytes; nothing talks to the network. The functions detect the target
host both as raw text and hidden behind hyperlink display text, and always
capture the FULL URL (domain + the URI path after it), not just the domain.
"""

import io
import logging
import re
import zipfile
from dataclasses import dataclass, asdict
from xml.etree import ElementTree as ET

# OOXML namespaces used when parsing .docx internals.
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

log = logging.getLogger("ghe_scraper")


@dataclass
class Match:
    site: str             # site display name / URL the match was found in
    source_type: str      # "page" | "document"
    title: str            # page title or file name
    location_url: str     # link to the page / document in SharePoint
    match_type: str       # "hyperlink" | "raw-text" | "webpart-data" | "binary"
    display_text: str     # visible text the link hid behind (if any)
    target_url: str       # the full ghe.hedgeserv.net/... URL (domain + path)


# --------------------------------------------------------------------------- #
# Matching
# --------------------------------------------------------------------------- #
def build_patterns(needle: str) -> re.Pattern:
    """
    Compile the URL-token matcher for the configured needle.

    It grabs the whole token containing the needle — scheme (optional), host,
    and the URI path/query after it — stopping only at whitespace, quotes, or
    brackets. So `https://ghe.hedgeserv.net/org/repo/blob/main/x` is captured
    in full, not truncated to the domain.
    """
    escaped = re.escape(needle)
    return re.compile(
        r"[^\s\"'<>()\[\]]*" + escaped + r"[^\s\"'<>()\[\]]*",
        re.IGNORECASE,
    )


def scan_html(html: str, needle: str, bare_re: re.Pattern) -> list[tuple[str, str, str]]:
    """
    Parse an HTML fragment and return (match_type, display_text, target_url)
    tuples. Catches both <a href> hyperlinks (the "hidden behind display text"
    case) and raw-text mentions, deduping raw text already covered by a link.
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

    # Any other element attribute carrying the host (iframe src, data-* blobs
    # that link web parts like Quick Links stash their URLs in, etc.).
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
# Document parsers
# --------------------------------------------------------------------------- #
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
    full URL but not reliably associate display text, so matches are "binary".
    """
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

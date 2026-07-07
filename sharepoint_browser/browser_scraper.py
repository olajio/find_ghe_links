#!/usr/bin/env python3
"""
SharePoint -> GHE link scraper (browser-session edition).

For tenants that block user consent for all OAuth apps (so both the Graph and
REST editions hit "Need admin approval"). Instead of an OAuth app, this drives a
real browser with Playwright: you sign in to SharePoint yourself (password +
MFA) in a normal browser window, and the tool reuses your **session cookies** to
call the SharePoint REST API — exactly what your browser does when you click
around the site. No app registration, nothing to consent to.

Finds every occurrence of a target host (default: ghe.hedgeserv.net) across the
site's pages and documents — as raw text or hidden behind hyperlink display text
— and captures the FULL URL, including the URI path after the domain.

See README.md for setup and usage.
"""

import argparse
import configparser
import logging
import os
import re
import sys
import time

from playwright.sync_api import sync_playwright

from matchers import (
    DOC_PARSERS,
    Match,
    build_patterns,
    dedupe,
    scan_html,
    write_report,
)

STATE_FILE = ".sp_session.json"          # persisted cookies/login state
ACCEPT_VERBOSE = "application/json;odata=verbose"

log = logging.getLogger("ghe_scraper")


# --------------------------------------------------------------------------- #
# SharePoint client backed by the browser's authenticated session
# --------------------------------------------------------------------------- #
class BrowserSharePoint:
    """Same interface as the REST edition's client, but requests ride on the
    browser context's cookies via Playwright's APIRequestContext."""

    def __init__(self, request_context, hostname: str):
        self.rc = request_context
        self.hostname_url = f"https://{hostname}"

    def _get(self, url: str):
        """GET with retry/backoff on 429/5xx, honoring Retry-After."""
        attempt = 0
        while True:
            attempt += 1
            resp = self.rc.get(url, headers={"Accept": ACCEPT_VERBOSE})
            if resp.status in (429, 503, 504) and attempt <= 6:
                wait = int(resp.headers.get("retry-after", 0)) or min(2 ** attempt, 60)
                log.warning("Throttled (%s). Waiting %ss …", resp.status, wait)
                time.sleep(wait)
                continue
            if not resp.ok:
                raise RuntimeError(f"HTTP {resp.status} for {url}")
            return resp

    def entity(self, url: str) -> dict:
        return self._get(url).json().get("d", {})

    def collection(self, url: str):
        while url:
            data = self._get(url).json().get("d", {})
            for item in data.get("results", []):
                yield item
            url = data.get("__next")

    def absolute(self, server_relative_url: str) -> str:
        return self.hostname_url + server_relative_url

    def download(self, server_relative_url: str) -> bytes:
        literal = server_relative_url.replace("'", "''")
        url = (
            f"{self.hostname_url}/_api/web"
            f"/GetFileByServerRelativePath(decodedurl='{literal}')/$value"
        )
        return self._get(url).body()


# --------------------------------------------------------------------------- #
# Login / session
# --------------------------------------------------------------------------- #
def wait_for_login(context, site_abs: str, timeout: int = 300) -> None:
    """Poll the SharePoint REST API until the session is authenticated.

    Returns silently once a cached session is valid; otherwise waits (up to
    `timeout` seconds) for the user to finish signing in in the browser window.
    """
    test_url = f"{site_abs}/_api/web?$select=Title"
    deadline = time.time() + timeout
    announced = False
    while time.time() < deadline:
        try:
            resp = context.request.get(test_url, headers={"Accept": ACCEPT_VERBOSE})
            if resp.ok and "application/json" in resp.headers.get("content-type", ""):
                return
        except Exception:
            pass
        if not announced:
            log.info("Please complete the sign-in in the browser window …")
            announced = True
        time.sleep(3)
    raise RuntimeError(
        "Timed out waiting for sign-in. Re-run and finish logging in within "
        f"{timeout} seconds."
    )


# --------------------------------------------------------------------------- #
# Web (site) discovery  — identical logic to the REST edition
# --------------------------------------------------------------------------- #
def resolve_web(sp: BrowserSharePoint, site_path: str) -> dict:
    site_abs = sp.hostname_url + "/" + site_path.strip("/")
    web = sp.entity(f"{site_abs}/_api/web?$select=Title,Url,ServerRelativeUrl")
    if not web:
        raise RuntimeError(f"Could not resolve site at {site_abs}")
    return web


def discover_webs(sp: BrowserSharePoint, root_web: dict, recursive: bool) -> list[dict]:
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
        except Exception as exc:
            log.warning("Could not list subwebs of %s: %s", base, exc)
    return webs


# --------------------------------------------------------------------------- #
# Page crawling (Site Pages library: CanvasContent1 + WikiField)
# --------------------------------------------------------------------------- #
def pages_library_title(sp: BrowserSharePoint, web_abs: str) -> str:
    try:
        libs = list(sp.collection(
            f"{web_abs}/_api/web/lists?$filter=BaseTemplate eq 119&$select=Title"
        ))
        if libs:
            return libs[0]["Title"]
    except Exception:
        pass
    return "Site Pages"


def crawl_pages(sp: BrowserSharePoint, web: dict, needle: str, bare_re: re.Pattern) -> list[Match]:
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
    except Exception as exc:
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
    sp: BrowserSharePoint, web: dict, needle: str, bare_re: re.Pattern, max_file_mb: int
) -> list[Match]:
    web_abs = web["Url"]
    site_name = web.get("Title") or web_abs.rstrip("/").split("/")[-1]
    try:
        libs = list(sp.collection(
            f"{web_abs}/_api/web/lists?$filter=BaseTemplate eq 101 and Hidden eq false"
            "&$select=Title"
        ))
    except Exception as exc:
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
            for it in sp.collection(url):
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
        except Exception as exc:
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
    parser = argparse.ArgumentParser(
        description="SharePoint -> GHE link scraper (browser-session edition)"
    )
    parser.add_argument("-c", "--config", default="config.ini", help="Path to config file")
    parser.add_argument("--no-documents", action="store_true", help="Scan pages only")
    parser.add_argument("--recursive", action="store_true", help="Also crawl all subsites")
    parser.add_argument("--site-path", help="Override site_path from the config file")
    parser.add_argument(
        "--headless", action="store_true",
        help="Run without a visible window (only works when a cached session in "
             f"{STATE_FILE} is still valid; can't do a fresh interactive login).",
    )
    parser.add_argument(
        "--channel",
        help="Use an installed browser instead of Playwright's bundled Chromium "
             "(e.g. 'chrome' or 'msedge'). Avoids needing 'playwright install'.",
    )
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
    docs_enabled = cfg.getboolean("documents", "enabled", fallback=True) and not args.no_documents
    max_file_mb = cfg.getint("documents", "max_file_mb", fallback=25)
    login_timeout = cfg.getint("browser", "login_timeout", fallback=300)
    browser_channel = args.channel or cfg.get("browser", "channel", fallback="").strip()
    csv_path = cfg.get("output", "csv_path", fallback="ghe_links_report.csv")
    xlsx_path = cfg.get("output", "xlsx_path", fallback="ghe_links_report.xlsx")

    bare_re = build_patterns(needle)
    hostname_url = f"https://{hostname}"
    site_abs = hostname_url + "/" + site_path.strip("/")

    matches: list[Match] = []
    with sync_playwright() as p:
        launch_kwargs = {"headless": args.headless}
        if browser_channel:
            # Use an already-installed browser (Chrome/Edge) so no Chromium
            # download is needed — handy behind proxies that block the download.
            launch_kwargs["channel"] = browser_channel
            log.info("Using installed browser channel: %s", browser_channel)
        browser = p.chromium.launch(**launch_kwargs)
        ctx_kwargs = {"accept_downloads": True}
        if os.path.exists(STATE_FILE):
            ctx_kwargs["storage_state"] = STATE_FILE
            log.info("Reusing cached session from %s", STATE_FILE)
        context = browser.new_context(**ctx_kwargs)

        page = context.new_page()
        log.info("Opening %s …", site_abs)
        try:
            page.goto(site_abs, wait_until="domcontentloaded")
        except Exception:
            pass  # login redirects can interrupt navigation; the poll below decides

        wait_for_login(context, site_abs, timeout=login_timeout)
        context.storage_state(path=STATE_FILE)  # persist for next run
        log.info("Authenticated. Session saved to %s", STATE_FILE)

        sp = BrowserSharePoint(context.request, hostname)

        root_web = resolve_web(sp, site_path)
        log.info("Site resolved: %s", root_web.get("Url"))
        webs = discover_webs(sp, root_web, recursive)
        if recursive:
            log.info("Recursive mode: crawling %d site(s).", len(webs))

        for web in webs:
            if len(webs) > 1:
                log.info("== Site: %s ==", web.get("Url"))
            matches += crawl_pages(sp, web, needle, bare_re)
            if docs_enabled:
                matches += crawl_documents(sp, web, needle, bare_re, max_file_mb)
        if not docs_enabled:
            log.info("Document scan disabled.")

        browser.close()

    matches = dedupe(matches)
    write_report(matches, csv_path, xlsx_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())

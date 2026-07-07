# SharePoint → GHE link scraper (browser-session edition)

For tenants that block user consent for **all** OAuth apps — where both the
Graph edition and the REST/`AllSites.Read` edition hit **"Need admin approval."**

Instead of registering or consenting to an app, this edition drives a **real
browser** with Playwright. You sign in to SharePoint yourself (password + MFA)
in a normal browser window, and the tool reuses your **session cookies** to call
the SharePoint REST API — exactly what your browser does when you click around
the site. There is no OAuth app involved, so there is nothing for an admin to
approve.

It does the same job as the other editions: finds every occurrence of a target
host (default `ghe.hedgeserv.net`) across the site's pages and documents — as
raw text **or** hidden behind hyperlink display text — and captures the **full
URL, including the URI path after the domain** (e.g.
`https://ghe.hedgeserv.net/org/repo/blob/main/file`, not just the domain).

> **Must run where you can sign in.** This needs a visible browser window for
> the interactive login, so run it on your **own machine** (laptop/desktop),
> not on a headless server. Session cookies expire after a few hours; when they
> do, it simply opens the login window again.

## Files

| File | Purpose |
|------|---------|
| `browser_scraper.py` | Main script: browser login → crawl pages → crawl documents → report |
| `matchers.py` | Self-contained link matching + document parsers (no network) |
| `config.example.ini` | Copy to `config.ini` and edit |
| `requirements.txt` | Python dependencies |

This folder is independent — nothing here imports from the other editions.

## What it finds

| Column | Meaning |
|--------|---------|
| `site` | The site/subsite the match was found in |
| `source_type` | `page` or `document` |
| `title` | Page title or file name |
| `location_url` | Link to the page/document in SharePoint |
| `match_type` | `hyperlink` (hidden behind text), `raw-text`, `webpart-data`, or `binary` |
| `display_text` | The visible text the link hid behind (if any) |
| `target_url` | The **full** `ghe.hedgeserv.net/...` URL (domain + path) |

## Setup

```bash
cd sharepoint_browser
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
playwright install chromium         # one-time: download the browser
cp config.example.ini config.ini    # then edit if needed
```

Defaults target
`https://hedgeservcorp.sharepoint.com/sites/GlobalTechnology/MonitoringAndAnalytics`.

## Run

```bash
python browser_scraper.py                 # opens a browser window for you to sign in
python browser_scraper.py --no-documents  # pages only (faster)
python browser_scraper.py --recursive     # also crawl every subsite under site_path
python browser_scraper.py --site-path /sites/GlobalTechnology --recursive   # sweep the parent tree
python browser_scraper.py --headless      # reuse a cached session with no window (see below)
python browser_scraper.py --verbose       # debug logging
```

### How sign-in works

1. A Chromium window opens on your SharePoint site.
2. Sign in normally — username, password, MFA, "stay signed in", whatever your
   org uses. The tool waits until it detects an authenticated session.
3. It saves the session to `.sp_session.json` and starts crawling.
4. On later runs within the cookie lifetime, it reuses that file and starts
   immediately. Add `--headless` to run with no window while the session is
   still valid. Once it expires, drop `--headless` to sign in again.

Nothing about this asks for app consent — you are simply using the SharePoint
website as yourself, and the tool reads through that same session.

## What gets scanned

- **Pages:** the Site Pages library (`CanvasContent1` for modern pages,
  `WikiField` for classic wiki pages). Parsed for `<a href>` links and raw text;
  link web parts (Quick Links, Hero) are also scanned in their embedded JSON.
- **Documents:** every non-hidden document library (`BaseTemplate 101`):
  - `.docx` — text + hyperlink relationship targets
  - `.xlsx` — cell text + cell hyperlinks
  - `.pdf` — page text + URI link annotations
  - `.xls` — cell text + hyperlinks (via `xlrd`)
  - `.doc`, `.ppt` — raw OLE byte scan (ASCII + UTF-16-LE); recovers the URL but
    not the display text, so these are typed `binary`

  Files larger than `max_file_mb` (default 25 MB) are skipped.

## Scope

By default, only the site at `site_path`. Use `--recursive` (or
`recurse_subsites = true`) to include subsites; point `site_path` at
`/sites/GlobalTechnology` to sweep the whole parent tree. The `site` column
attributes each match to its subsite.

## Notes / limitations

- Only content **you** can access is returned — it's your own session.
- Session cookies expire after a few hours; you'll be prompted to sign in again.
- Throttling (HTTP 429/503) is handled automatically with backoff.
- If your org enforces Conditional Access that blocks scripted API access even
  from a browser session, some calls may be limited — but a genuine browser
  session usually passes.

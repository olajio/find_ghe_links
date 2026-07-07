# SharePoint → GHE link scraper (REST edition)

A self-contained variant of the scraper that authenticates against the
**SharePoint REST API** using the delegated **`AllSites.Read`** scope, instead
of Microsoft Graph's admin-restricted `Sites.Read.All`.

**Why this edition exists:** the Graph edition in the parent folder needs the
`Sites.Read.All` scope, which many tenants gate behind an admin-consent prompt
("Need admin approval"). `AllSites.Read` means *"read what the signed-in user
can already read in SharePoint"* and is frequently **user-consentable without an
admin**. If your org only restricts the Graph admin scopes, this edition lets
you run without waiting on IT.

> **Honest caveat:** if your tenant blocks **all** user consent for any app
> (a stricter policy), even `AllSites.Read` will still prompt for admin — no
> code can escape that. You'd then need a one-time admin consent (see the parent
> folder's README, "Option A").

It does the same job: finds every occurrence of a target host (default
`ghe.hedgeserv.net`) across the site's pages and documents — as raw text **or**
hidden behind hyperlink display text — and **captures the full URL, including
the URI path after the domain** that points to the actual resource (e.g.
`https://ghe.hedgeserv.net/org/repo/blob/main/file`, not just the domain).

## Files

| File | Purpose |
|------|---------|
| `rest_scraper.py` | Main script: auth → resolve site → crawl pages → crawl documents → report |
| `matchers.py` | Self-contained link matching + document parsers (no network) |
| `config.example.ini` | Copy to `config.ini` and edit |
| `requirements.txt` | Python dependencies |

This folder is independent of the parent Graph scraper — nothing here imports
from it.

## What it finds

For every match it records:

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
cd sharepoint_rest
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
cp config.example.ini config.ini   # then edit if needed
```

Defaults target
`https://hedgeservcorp.sharepoint.com/sites/GlobalTechnology/MonitoringAndAnalytics`.

## Run

```bash
python rest_scraper.py                 # uses config.ini
python rest_scraper.py --no-documents  # pages only (faster)
python rest_scraper.py --recursive     # also crawl every subsite under site_path
python rest_scraper.py --site-path /sites/GlobalTechnology --recursive   # sweep the parent tree
python rest_scraper.py --verbose       # debug logging
```

### Sign-in (device code)

On first run the script prints a URL and a code:

```
To sign in, use a web browser to open the page https://microsoft.com/devicelogin
and enter the code ABCD-EFGH to authenticate.
```

Open it, enter the code, sign in with your **hedgeserv** account. The token is
cached in `.sp_token_cache.bin` for subsequent runs.

If sign-in still shows "Need admin approval", your tenant blocks user consent
entirely — switch the `client_id` in `config.ini` to the Azure CLI public
client (`04b07795-8ddb-461a-bbee-02f9e1bf7b46`) and try once more; if it still
prompts, a one-time admin consent is required (see parent folder README).

## What gets scanned

- **Pages:** the Site Pages library (`CanvasContent1` for modern pages,
  `WikiField` for classic wiki pages) via SharePoint REST. Content is HTML,
  parsed for `<a href>` links and raw text; link web parts (Quick Links, Hero)
  are additionally scanned in their embedded JSON.
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
`recurse_subsites = true`) to include all subsites; point `site_path` at
`/sites/GlobalTechnology` to sweep the whole parent tree. The `site` column
attributes each match to its subsite.

## Notes / limitations

- Only content **you** can access is returned (delegated permissions).
- Throttling (HTTP 429/503) is handled automatically with backoff.
- The report is regenerated fresh on each run.

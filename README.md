# SharePoint → GHE link scraper

Crawls a SharePoint Online site (pages **and** documents) and reports every
occurrence of a target host — by default `ghe.hedgeserv.net` — whether it
appears as **raw text** or is **hidden behind hyperlink display text**.

Built to find leftover links to the decommissioned `https://ghe.hedgeserv.net`
after the migration to `https://github.com`.

See [`PLAN.md`](./PLAN.md) for the full design and rationale.

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
| `target_url` | The actual `ghe.hedgeserv.net/...` URL |

The tricky case — a link where the visible text is something like "our repo"
but the `href` points to `ghe.hedgeserv.net` — is handled by parsing each
page's raw HTML and each document's embedded hyperlink targets, rather than
relying on SharePoint search (which indexes visible text, not `href`s).

## Setup

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
cp config.example.ini config.ini   # then edit if needed
```

The defaults in `config.example.ini` are already set for
`https://hedgeservcorp.sharepoint.com/sites/GlobalTechnology/MonitoringAndAnalytics`.

## Run

```bash
python scraper.py                 # uses config.ini
python scraper.py --no-documents  # pages only (faster)
python scraper.py --recursive     # also crawl every subsite under site_path
python scraper.py --site-path /sites/GlobalTechnology --recursive   # sweep the whole parent tree
python scraper.py --verbose       # debug logging
python scraper.py -c other.ini    # alternate config
```

### Sign-in (device code)

On first run the script prints something like:

```
To sign in, use a web browser to open the page https://microsoft.com/devicelogin
and enter the code ABCD-EFGH to authenticate.
```

Open that URL, enter the code, and sign in with your **hedgeserv** account.
The token is cached in `.token_cache.bin` so subsequent runs are silent until
it expires.

No custom Azure AD app registration is required — the script uses Microsoft's
first-party **"Microsoft Graph Command Line Tools"** public client, which
supports device-code login with delegated permissions.

> **Note on permissions:** the login requests the delegated scopes
> `Sites.Read.All` and `Files.Read.All`. Some tenants require a one-time admin
> consent for these. If your sign-in shows a "needs admin approval" screen, ask
> a SharePoint/Azure AD admin to consent once; after that the script works for
> you. You only ever see content **you** already have access to.

## Output

Two files are written (paths configurable in `config.ini`):

- `ghe_links_report.csv`
- `ghe_links_report.xlsx`

Plus a console summary with counts by source type and match type.

## What gets scanned

- **Pages:** modern SharePoint pages and news posts (via Microsoft Graph
  `sitePage` + `canvasLayout`). Text web parts are HTML-parsed for links and
  raw text; link web parts (Quick Links, Hero, …) are scanned in their JSON
  data, where those URLs actually live.
- **Documents:** files in the site's document libraries:
  - `.docx` — text + hyperlink relationship targets
  - `.xlsx` — cell text + cell hyperlinks
  - `.pdf` — page text + URI link annotations
  - `.xls` — cell text + hyperlinks (via `xlrd`)
  - `.doc`, `.ppt` — raw OLE byte scan (ASCII + UTF-16-LE); recovers the URL but
    not the display text, so these are typed `binary`

  Files larger than `max_file_mb` (default 25 MB) are skipped.

## Scope

By default, scoped to the single subsite configured in `config.ini`
(`MonitoringAndAnalytics`) — it does not walk sibling sites.

To crawl a site **and all its subsites**, use `--recursive` (or set
`recurse_subsites = true` in `config.ini`). To sweep the whole parent tree,
point `site_path` at `/sites/GlobalTechnology` and run recursively:

```bash
python scraper.py --site-path /sites/GlobalTechnology --recursive
```

The `site` column in the report tells you which site each match came from.

## Notes / limitations

- Only content you can access is returned (delegated permissions).
- Throttling (HTTP 429/503) is handled automatically with backoff.
- The report is regenerated fresh on each run.

# Plan: SharePoint → GHE link scraper

## Goal
Crawl every page (and document) in the target SharePoint site, find every
occurrence of `https://ghe.hedgeserv.net` — whether it appears as raw text
**or** as a hyperlink hidden behind display text (an `<a href>` pointing there)
— and produce a report of:

- the **actual target URL** (the full `ghe.hedgeserv.net/...` link)
- the **display text** it was hiding behind (if any)
- a **link to the SharePoint page/document** where it was found

## The core challenge: two kinds of matches
A naive text search is not enough:

| Case | Example in page | How to catch it |
|------|-----------------|-----------------|
| Raw URL | `See https://ghe.hedgeserv.net/org/repo` | Text/regex match |
| Hyperlinked text | "our repo" linking to `https://ghe.hedgeserv.net/org/repo` — user only sees "our repo" | Must inspect the `href` attribute, not the visible text |

The second case is the one SharePoint's own search often **misses**, because
search indexes visible/rendered text, not `href` attributes. So we parse the
actual HTML of each page and the embedded hyperlink targets of each document.

## Confirmed decisions
- **Platform:** SharePoint Online, tenant `hedgeservcorp.sharepoint.com`
- **Target:** `https://hedgeservcorp.sharepoint.com/sites/GlobalTechnology/MonitoringAndAnalytics`
  (this subsite only — do not walk up to `GlobalTechnology` or sideways to siblings)
- **Auth:** device-code login as the user (no pre-existing Azure AD app registration)
- **Content:** pages **and** documents

## Approach: Microsoft Graph API (not screen-scraping)
The Graph API returns raw page HTML (including every `href`) directly, handles
pagination cleanly, is faster and more reliable than a headless-browser crawl,
and avoids fighting JavaScript rendering and auth timeouts.

`Graph API base: https://graph.microsoft.com/v1.0`

### 1. Authentication (device-code, no app registration needed)
MSAL device-code flow needs *a* client ID even for delegated login. Since there
is no dedicated app registration, use Microsoft's well-known public client
**"Microsoft Graph Command Line Tools"** (`14d82eec-204b-4c2f-b7e8-296a70dab67e`)
— a first-party multi-tenant app that supports device-code flow. The script
prints a code + URL; the user signs in with their hedgeserv account in a
browser.

- Scopes: `Sites.Read.All`, `Files.Read.All`.
- **Caveat:** some tenants require admin consent for `Sites.Read.All` even via
  delegated flow. If `hedgeservcorp` enforces that, the first login hits a
  consent screen the script cannot bypass, requiring a one-time admin approval.
  We find out on first run. Fallback: a quick app registration.

### 2. Resolve the site
`GET /sites/hedgeservcorp.sharepoint.com:/sites/GlobalTechnology/MonitoringAndAnalytics`
→ site `id`. Scoped to just this subsite.

### 3. Crawl pages
- List modern pages/news:
  `GET /sites/{id}/pages/microsoft.graph.sitePage?$expand=canvasLayout`.
- Also read the classic `SitePages` library for any wiki/web-part pages the
  modern API misses.
- Parse each page's HTML (BeautifulSoup): collect every `<a href>` containing
  `ghe.hedgeserv.net` **and** every raw-text occurrence; dedupe so a linked URL
  is not double-counted as loose text.

### 4. Crawl documents
- Enumerate document libraries: `GET /sites/{id}/drives` → walk each drive's
  folders/files.
- For supported types, download and extract **both** visible text and embedded
  hyperlink targets (this is what catches links hidden behind display text
  inside docs):
  - **.docx** — parse the zip; hyperlink targets live in `document.xml.rels`.
  - **.xlsx** — sheet rels + cell text/formulas.
  - **.pdf** — text plus URI link annotations.
- Match `ghe.hedgeserv.net` in either stream.

### 5. Output report
CSV + Excel with columns:

`source_type` (page | document), `title`, `location_url` (link to the page or
the file in SharePoint), `match_type` (hyperlink | raw-text), `display_text`,
`target_url`.

Plus a console summary (counts per source).

### 6. Robustness
- 429 handling with `Retry-After` backoff — Graph rate-limits on a bulk crawl.
- Resumable/idempotent logging so a long crawl can restart.
- `config` for the target URL + tenant.

## Files to create
```
scraper.py           # main: auth → crawl pages → crawl docs → report
config.example.ini   # tenant + target site URL
requirements.txt     # msal, requests, beautifulsoup4, python-docx, openpyxl, pypdf, pandas
README.md            # setup + run steps
```

## Testing note
The internal SharePoint is not reachable from the build environment. The code
will be built to be correct against the documented Graph API and structured
cleanly, but real validation is running it against the tenant. Logging will be
verbose enough to diagnose environment mismatches quickly.

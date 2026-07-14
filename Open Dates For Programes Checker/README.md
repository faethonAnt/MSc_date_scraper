# MSc Application Dates Scraper

> Automatically scrapes Greek (and other) university MSc programme websites to find current application open dates, deadlines, and apply links — then exports the results for a dashboard.

---

## What it does

The scraper reads a CSV of MSc programmes, visits each programme's website, and uses an LLM (via ScrapeGraphAI) to extract structured admissions data: when applications open, when they close, and where to apply. It handles Greek and English content, PDFs, DOCX files, announcement pages, and external portals. Results are saved as JSON and CSV.

---

## Installation

**Prerequisites:** Python 3.10+, an OpenAI-compatible API key

```bash
pip install scrapegraphai playwright pdfplumber python-docx requests beautifulsoup4 python-dotenv python-dateutil
playwright install chromium
```

Create a `.env` file with your API key:
```
OPENAI_API_KEY=your_key_here
```

---

## How to run

```bash
# Full run (Pass 1 + Pass 2)
python scraper.py

# Resume an interrupted run
python scraper.py --resume

# Target specific programme IDs only
python scraper.py --ids 996,1202,1387

# Only re-run programmes with no dates found yet
python scraper.py --missing-only

# Skip Pass 1, re-run Pass 2 only
python scraper.py --pass2-only

# Change the LLM model
python scraper.py --model openai/gpt-4o

# Limit to first N programmes (useful for testing)
python scraper.py --limit 10 --offset 0

# Include archived programmes
python scraper.py --include-archived

# Control parallel workers (default: 3)
python scraper.py --workers 5
```

---

## Overall flow

```
programmes_final.csv
        │
        ▼
  load_programmes()          ← reads all active MSc programmes
        │
        ▼
┌─────────────────────────────────────┐
│              PASS 1                 │
│  worker_pass1() for each programme  │
│                                     │
│  1. fetch_html(url)                 │
│  2. Domain hub? → hub logic         │
│  3. Extract + validate dates        │
│  4. scrape_url() with LLM           │
│  5. No dates? → try PDF/DOCX links  │
│  6. Still no dates? → try           │
│     announcement links              │
│  7. Still no dates? → try_subpages  │
└────────────────┬────────────────────┘
                 │
                 ▼
        pass1_results.json
                 │
                 ▼
┌─────────────────────────────────────┐
│              PASS 2                 │
│  worker_pass2() for programmes      │
│  where Pass 1 had no deadline or    │
│  found an external apply portal     │
│                                     │
│  Follows apply_button_url:          │
│  - If PDF/DOCX → extract text       │
│  - If HTML page → scrape_url()      │
└────────────────┬────────────────────┘
                 │
                 ▼
        pass2_results.json
                 │
                 ▼
           merge() + export()
                 │
                 ▼
    results.json + results.csv
```

---

## Configuration constants

| Name | Value | Purpose |
|---|---|---|
| `DEFAULT_CSV` | `programmes_final.csv` | Input file of MSc programmes |
| `DEFAULT_MODEL` | `openai/gpt-4.1-nano` | LLM used for extraction |
| `OUTPUT_DIR` | `results/` | Folder for all output files |
| `MAX_RETRIES` | `3` | LLM call retries per URL |
| `MAX_RETRIES_PER_DOMAIN` | `15` | Max subpage attempts per domain |
| `MAX_SITEMAP_ATTEMPTS` | `15` | Max sitemap URLs tried per domain |
| `MAX_SUFFIX_ATTEMPTS` | `50` | Max suffix paths tried per domain |
| `RETRY_DELAY` | `4s` | Wait between failed LLM retries |
| `REQUEST_DELAY` | `1.0s` | Politeness delay after each scrape |
| `SCRAPE_TIMEOUT` | `90s` | Hard cap per LLM scrape call |
| `DEFAULT_WORKERS` | `3` | Parallel threads |

---

## Dictionaries

### `DOMAIN_ANNOUNCEMENT_HUBS`
Maps a domain substring to a central announcement/news page URL. When a programme's domain matches a key, the scraper goes to the hub page first instead of the programme's own homepage — because some universities (like EAP) post all admissions dates centrally rather than on individual programme pages.

```python
DOMAIN_ANNOUNCEMENT_HUBS = {
    "eap.gr": "https://www.eap.gr/en/tag/invitation-en/",
    "mscie.hmu.gr": "https://mscie.hmu.gr/news/",
    ...
}
```

Results from a hub are cached in `_hub_cache` (protected by `_hub_cache_lock`) so multiple programmes on the same domain share one scrape.

### `ANNOUNCEMENT_KEYWORDS`
Strings matched against link text on a page to identify links that lead to admissions announcements or calls for applications. Used by `find_announcement_links()`. Includes English terms (`"call for applications"`, `"intake"`), Greek terms (`"προκήρυξη"`, `"πρόσκληση"`), and Greeklish (`"prosklisi"`).

### `PDF_KEYWORDS`
Strings matched against both the link text and the href of `<a>` tags to identify PDF or DOCX files that are likely to contain admissions information. Used by `find_pdf_links()`. Broader than `ANNOUNCEMENT_KEYWORDS` — includes generic words like `"application"`, `"brochure"`, `"εισαγωγή"`.

### `SUBPAGE_SUFFIXES`
A list of URL path suffixes (English, Greek, and Greeklish) the scraper appends to a programme's base domain when trying to find application info via `try_subpages()`. Examples: `/admissions`, `/call-for-applications`, `/προκήρυξη`, `/aitisi`.

### `CONTEXT_KEYWORDS` and `STRONG_DEADLINE_KEYWORDS`
Used during raw HTML date extraction (`extract_dates_with_context()`) to decide whether a found date is application-related. `CONTEXT_KEYWORDS` is a broad set; `STRONG_DEADLINE_KEYWORDS` is a stricter subset. A date surrounded only by irrelevant text (e.g. a bare post-date line) is discarded unless a strong keyword is nearby.

### `DATE_PATTERNS` / `DATE_RE`
A compiled regex combining all date formats the scraper recognises:
- `15/06/2026`, `2026-06-15`, `15-06-26`
- `15 June 2026`, `June 15, 2026`, `June 8–26, 2026`
- `15 Jun 2026`, `15th of June 2026`
- `15 Ιουνίου 2026`, `Ιούνιος 2026`

`DATE_RE` is used everywhere a date needs to be detected in text.

### `DEADLINE_CSS_CLASSES`
CSS class names that signal a deadline field in HTML: `"ew-deadline-text"`, `"deadline"`, `"application-deadline"`, etc. (Reserved for future use in targeted HTML parsing.)

### `_GREEK_LOOKALIKE_TO_LATIN`
A character translation table that converts Greek letters that visually look like Latin letters (e.g. Greek `Α` → Latin `A`) to their Latin equivalents. Used by `_delookalike()` so keyword matching isn't defeated by font/encoding tricks.

---

## Functions

### Data loading

**`load_programmes(csv_path, active_only)`**
Reads `programmes_final.csv` and returns a list of programme dicts. Each dict holds the programme's website URL, English/Greek name, university, department, city, topics, languages, ECTS, semesters, tuition, study modes, email, phone, and whether a scholarship is offered. Archived programmes are filtered out unless `--include-archived` is passed.

**`get_graph_config(model)`**
Reads `OPENAI_API_KEY` from the environment and builds the configuration dict for ScrapeGraphAI.

**`import_scraper()`**
Lazily imports `SmartScraperGraph` from ScrapeGraphAI, giving a clean error message if the package isn't installed.

---

### HTML / document fetching

**`fetch_html(url)`**
Downloads a page's raw HTML using `requests`, with a browser-like User-Agent. Returns `None` on failure.

**`fetch_sitemap_urls(base_url)`**
Fetches `/sitemap.xml` for a domain (recursing into sitemap indexes), filters URLs to those containing application-related keywords, sorts them newest-year-first, and returns up to 10. Results are cached in `_sitemap_cache`.

**`extract_pdf_text(url)`**
Downloads a PDF and extracts its text with `pdfplumber`. Returns up to 8 000 characters, or `None` if the download or extraction fails.

**`extract_docx_text(url)`**
Downloads a DOCX file and extracts paragraph text with `python-docx`. Returns up to 8 000 characters, or `None` on failure.

---

### Link finding

**`find_pdf_links(html, base_url)`**
Parses HTML with BeautifulSoup and returns up to 5 absolute URLs of `.pdf` or `.docx` links whose href or link text matches any `PDF_KEYWORDS` term. Uses `_delookalike()` to normalise Greek-lookalike characters before matching.

**`find_announcement_links(html, base_url)`**
Same approach as `find_pdf_links`, but matches `<a>` tag text against `ANNOUNCEMENT_KEYWORDS` and returns up to 5 absolute URLs (any file type). These are links to announcement articles, not documents.

---

### Date extraction (pre-LLM)

**`extract_dates_with_context(text, source_label)`**
Scans plain text line by line, finds all `DATE_RE` matches, and for each collects 2 lines of surrounding context. Discards dates that look like bare publication-date lines (short line, no nearby keyword). Returns a list of `{raw, context, source, has_keyword}` dicts.

**`validate_dates(dates)`**
Takes the output of `extract_dates_with_context` and discards any date whose parsed year is outside 2026–2027. Adds `parsed_iso` and `year` fields.

**`build_date_list(prog, dates)`**
Formats the validated dates into a plain-text "structured list" string (including programme name and context snippets) that is passed as a `data:text/plain,...` URL to the LLM instead of a raw webpage — this is more reliable than letting the LLM scrape the full HTML.

**`extract_date_contexts(html)`**
Alternative helper that strips HTML to plain text and returns lines that contain either a date or an application keyword. Used as a compact context summary for the LLM. Returns up to 3 000 characters.

---

### LLM scraping

**`scrape_url(url, prompt, graph_config, SmartScraperGraph)`**
Core function. Calls ScrapeGraphAI's `SmartScraperGraph` with a `SCRAPE_TIMEOUT` second hard cap (via `_run_with_timeout`). Handles retries (up to `MAX_RETRIES`), rate-limit back-off, JSON fence stripping, nested-content unwrapping, and LangChain "Invalid json output" error recovery. Returns a cleaned dict with `status: "ok"` or `status: "error"`.

**`_run_with_timeout(func, timeout, *args, **kwargs)`**
Runs any function in a daemon thread and raises `TimeoutError` if it doesn't finish within `timeout` seconds. Using a daemon thread (not a ThreadPoolExecutor) ensures a stuck scrape never blocks process shutdown.

**`try_subpages(base_url, prompt, graph_config, SmartScraperGraph)`**
When the main page yields nothing, tries a combined list of sitemap URLs + `SUBPAGE_SUFFIXES` paths, scraping each one until it finds a page with `has_dates_on_page: true`. Caps at `MAX_RETRIES_PER_DOMAIN` total attempts and stops after 5 consecutive failures.

---

### Result helpers

**`has_current_dates(result)`**
Returns `True` if any of `application_deadline`, `application_open_date`, or `notes` contains a year ≥ 2026. Used to gate whether a result is "good enough" to accept.

**`_has_specific_date(result)`**
Stricter than `has_current_dates`. Returns `True` only if `application_deadline` or `application_open_date` contains a string that fully matches `DATE_RE` *and* includes a year ≥ 2026. A bare "2026" in a notes field is not enough.

**`date_relevance_score(text)`**
Returns 0–3 based on the most recent year found in `text` (3 = 2026+, 2 = 2025, 1 = 2023-2024, 0 = none). Used by `best_record` and `merge` to prefer higher-year dates.

**`best_record(records)`**
Given multiple pass-1 records for the same programme ID (can happen on resume), picks the one with the highest combined score: PDF-sourced > has_dates_on_page > deadline relevance > open-date relevance > has any deadline > has any open date > has notes.

**`extract_date_from_notes(notes)`**
Falls back to `notes` field: scans it with `DATE_RE` and returns the first match. Used in `merge()` when both `application_deadline` and `application_open_date` are empty but the notes text contains a date.

**`_needs_pass2(r)`**
Returns `True` if a Pass 1 result should trigger a Pass 2 follow-up: no dates on page, missing deadline, or open date is only a bare year without a month/day.

---

### JSON / response cleaning

**`strip_json_fences(text)`**
Removes ` ```json ` and ` ``` ` markdown fences the LLM sometimes wraps its response in.

**`unwrap_content(raw)`**
Some models return `{"content": "{...json...}"}` instead of the JSON directly. This function detects and unwraps that pattern.

**`clean_result(raw)`**
Converts placeholder strings (`"NA"`, `"N/A"`, `"NONE"`, `"NULL"`, `""`) to `None`, and collapses list values into comma-joined strings.

---

### Workers

**`worker_pass1(args)`**
The main per-programme scrape function, run in a thread pool. Full logic:

1. `fetch_html()` the programme URL.
2. If the domain is in `DOMAIN_ANNOUNCEMENT_HUBS`, scrape the hub (with caching + lock).
3. Otherwise, extract dates from HTML → `validate_dates()` → `build_date_list()` → `scrape_url()` with the structured list.
4. If no specific date found: try PDF/DOCX links via `find_pdf_links()` + `extract_pdf_text()` / `extract_docx_text()`.
5. If still no specific date: try announcement links via `find_announcement_links()`; for each, also try any PDF/DOCX links *on that announcement page*.
6. If still no specific date and no apply button: `try_subpages()`.
7. Tag result with `prog_id` and `scraped_at`.

**`worker_pass2(args)`**
Follows the `apply_button_url` (or `external_portal`) found in Pass 1. If the URL is a PDF or DOCX, extracts text and scrapes that; if it's a page, scrapes it with `PASS2_PROMPT`. If the domain is external, checks the sitemap first to find a better entry page.

---

### Orchestrators

**`run_pass1(programmes, ...)`**
Submits all programmes as `worker_pass1` tasks to a `ThreadPoolExecutor`. Handles graceful shutdown (SIGINT/SIGTERM), cancels remaining futures, and collects results.

**`run_pass2(pass1_results, ...)`**
Filters Pass 1 results to those needing Pass 2 (`_needs_pass2()`), then submits `worker_pass2` tasks. Same shutdown handling.

---

### Merge & export

**`merge(programmes, pass1, pass2)`**
Combines Pass 1 and Pass 2 results into one row per programme:
- If either pass found a PDF-sourced date, it wins.
- Otherwise, picks whichever pass has the higher `date_relevance_score` for deadline and open date separately.
- Falls back to extracting a date from the `notes` field if both date fields are empty.
- Returns a flat list of dicts ready for export.

**`export(rows)`**
Writes `results.json` and `results.csv`, then logs a summary: total programmes, found count (%), missing count, error count.

**`save_errors(pass1_results, prog_map)`**
Writes `results/errors.json` — a list of programmes where Pass 1 returned `status: "error"`, with their URL and error message.

---

### Utilities

**`load_json_safe(path)`** — reads a JSON file, returns `[]` if missing or corrupt.

**`save_json(data, path)`** — writes a list to a JSON file with indentation.

**`now_utc()`** — returns the current UTC time as an ISO 8601 string.

**`parse_args()`** — defines and parses all CLI arguments.

**`main()`** — top-level entry point: loads programmes, applies CLI filters (`--ids`, `--missing-only`, `--limit`, `--offset`), runs Pass 1, runs Pass 2, merges, exports.

---

## LLM prompts

### `PASS1_PROMPT`
Used for all standard page scrapes in Pass 1. Asks the LLM to return a JSON object with `application_open_date`, `application_deadline`, `intake_semester`, `apply_button_url`, `has_dates_on_page`, `application_status`, and `notes`. Key rules:
- Only extract dates from 2026 onward.
- Do not invent or guess years.
- Ignore ΦΕΚ/FEK government gazette citations.
- Prefer deadline/opening dates over publication or event dates.

### `PASS2_PROMPT`
Used for external apply portals in Pass 2. Similar to `PASS1_PROMPT` but also extracts `portal_name` and `requires_login` (for detecting login walls).

---

## Output files

All outputs are written to the `results/` folder.

| File | Contents |
|---|---|
| `results.json` | Final merged result, one object per programme |
| `results.csv` | Same as above in CSV format |
| `pass1_results.json` | Raw Pass 1 outputs (before merge) |
| `pass2_results.json` | Raw Pass 2 outputs (before merge) |
| `errors.json` | Programmes where scraping failed entirely |
| `scraper.log` | Full run log with timestamps |

### Key fields in the final results

| Field | Description |
|---|---|
| `open_date` | Date applications open (any format the site uses) |
| `deadline` | Application deadline |
| `intake` | Intake semester/year (e.g. "Fall 2026") |
| `apply_url` | Direct link to the application form or portal |
| `application_status` | `open` / `closed` / `rolling` / `coming_soon` / `not_mentioned` |
| `found` | `"yes"` if any date was found, `"no"` otherwise |
| `found_in_announcement` | URL of the announcement page the date came from |
| `scrape_status` | `ok` / `error` / `cancelled` |
| `pass2_status` | `ok` / `error` / `skipped` |
| `notes` | Free-text admissions notes from the LLM |
| `requires_login` | Whether the apply portal is behind a login wall |

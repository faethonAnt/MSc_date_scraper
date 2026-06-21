"""
deadline_scraper.py — universal, deadline-focused scraper.

Goal (deliberately narrower than scraper.py): for every programme in the CSV,
get a CORRECT application deadline for the large majority of cases, using one
identical process for every programme (no per-domain special cases).

Pipeline, kept as three separate, independently-testable phases:

  1. DISCOVERY — build one combined text pool per programme from:
       - the homepage
       - sitemap.xml (if present)
       - a fixed list of common subpage suffixes (admissions, apply, etc.)
       - every PDF/DOCX linked from any of the above pages
     Everything is gathered into one pool instead of being tried as
     sequential fallbacks, so a deadline buried in a PDF two clicks deep is
     caught in the same pass as one sitting on the homepage.

  2. EXTRACTION — pure code, no LLM. A bilingual (EL/EN), range-aware regex
     pulls every date / date-range candidate out of the pool, each tagged
     with: source URL, surrounding text, whether it's near a deadline-type
     keyword (deadline, until, closes, λήξη, μέχρι, προθεσμία) or an
     open/period-type keyword (from, period, αιτήσεις από, περίοδος), and
     whether a stale academic-year mention (e.g. "2024-25") sits nearby.

  3. DECISION — rule-based by default (deterministic, no API key needed):
     pick the best surviving candidate. Optionally, an LLM decision step can
     be wired in (see decide_with_llm) which must choose only from the
     literal candidate list — it is never shown raw page text and can never
     invent a date.

Output: results/deadline_results.json (+ .csv), one row per programme id,
with: deadline, deadline_source_url, deadline_context, found_via,
needs_review, candidates_found.

This file does not import or modify scraper.py, and does not read or write
the project's .env file. The optional LLM decision step (decide_with_llm)
is left for the user to wire up in their own environment, exactly like the
existing scraper.py already does with its own OPENAI_API_KEY / dotenv setup.
"""

import argparse
import csv
import io
import json
import logging
import re
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin, urlparse
from xml.etree.ElementTree import fromstring

import requests
from bs4 import BeautifulSoup

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s",
                     datefmt="%H:%M:%S")
log = logging.getLogger("deadline_scraper")

THIS_DIR = Path(__file__).resolve().parent
CSV_PATH = THIS_DIR / "programmes-api_msc-programmes.csv"
OUT_DIR = THIS_DIR / "results"

CURRENT_YEAR = datetime.now().year
TARGET_YEARS = {CURRENT_YEAR, CURRENT_YEAR + 1}  # accept this cycle / next

HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; DeadlineScraper/1.0)"}
TIMEOUT = 15
REQUEST_DELAY = 0.5

SUBPAGE_SUFFIXES = [
    "/admissions/", "/admission/", "/apply/", "/applications/",
    "/how-to-apply/", "/prospective-students/", "/registration/",
    "/registrations/", "/announcements/", "/news/", "/call-for-applications/",
    "/el/admissions/", "/el/apply/", "/el/announcements/",
]

# ---- bilingual, range-aware date patterns ----

_MONTHS_EN = r"(?:January|February|March|April|May|June|July|August|September|October|November|December)"
_MONTHS_EN_SHORT = r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"
_MONTHS_GR = (
    r"(?:Ιανουαρίου|Φεβρουαρίου|Μαρτίου|Απριλίου|Μαΐου|Μαίου|Ιουνίου|"
    r"Ιουλίου|Αυγούστου|Σεπτεμβρίου|Οκτωβρίου|Νοεμβρίου|Δεκεμβρίου|"
    r"Ιανουάριος|Φεβρουάριος|Μάρτιος|Απρίλιος|Μάιος|Μαϊος|Ιούνιος|"
    r"Ιούλιος|Αύγουστος|Σεπτέμβριος|Οκτώβριος|Νοέμβριος|Δεκέμβριος|"
    r"ιανουαρίου|φεβρουαρίου|μαρτίου|απριλίου|μαΐου|μαίου|ιουνίου|"
    r"ιουλίου|αυγούστου|σεπτεμβρίου|οκτωβρίου|νοεμβρίου|δεκεμβρίου)"
)

# Range patterns are listed BEFORE single-date patterns of the same family
# so the regex engine greedily matches the full range, not just its tail.
RANGE_PATTERNS = [
    # June 8 – 26, 2026   (same month)
    rf'\b{_MONTHS_EN}\s+\d{{1,2}}\s*[-–—]\s*\d{{1,2}},?\s+\d{{4}}\b',
    # April 1 – June 30, 2026   (different months, EN)
    rf'\b{_MONTHS_EN}\s+\d{{1,2}}\s*[-–—]\s*{_MONTHS_EN}\s+\d{{1,2}},?\s+\d{{4}}\b',
    # 1 April – 30 June 2026   (day-first EN)
    rf'\b\d{{1,2}}\s+{_MONTHS_EN}\s*[-–—]\s*\d{{1,2}}\s+{_MONTHS_EN}\s+\d{{4}}\b',
    # 1 Απριλίου – 30 Ιουνίου 2026   (GR, different months — the EAP case)
    rf'\b\d{{1,2}}\s+{_MONTHS_GR}\s*[-–—]\s*\d{{1,2}}\s+{_MONTHS_GR}\s+\d{{4}}\b',
    # 1 – 30 Ιουνίου 2026   (GR, same month, day range only)
    rf'\b\d{{1,2}}\s*[-–—]\s*\d{{1,2}}\s+{_MONTHS_GR}\s+\d{{4}}\b',
]

SINGLE_PATTERNS = [
    rf'\b\d{{1,2}}[/\-\.]\d{{1,2}}[/\-\.]\d{{2,4}}\b',          # 15/06/2025
    rf'\b\d{{4}}[/\-\.]\d{{1,2}}[/\-\.]\d{{1,2}}\b',             # 2025-06-15
    rf'\b\d{{1,2}}\s+{_MONTHS_EN}\s+\d{{4}}\b',                  # 15 June 2025
    rf'\b{_MONTHS_EN}\s+\d{{1,2}},?\s+\d{{4}}\b',                # June 15, 2025
    rf'\b\d{{1,2}}\s+{_MONTHS_EN_SHORT}\.?\s+\d{{4}}\b',         # 15 Jun 2025
    rf'\b\d{{1,2}}(?:st|nd|rd|th)\s+(?:of\s+)?{_MONTHS_EN}\s+\d{{4}}\b',  # 15th of June 2025
    rf'\b\d{{1,2}}\s+{_MONTHS_GR}\s+\d{{4}}\b',                  # 15 Ιουνίου 2025
    rf'\b{_MONTHS_GR}\s+\d{{4}}\b',                               # Ιούνιος 2025 (vague, low priority)
]

RANGE_RE = re.compile("|".join(RANGE_PATTERNS), re.IGNORECASE)
SINGLE_RE = re.compile("|".join(SINGLE_PATTERNS), re.IGNORECASE)
ANY_DATE_RE = re.compile("|".join(RANGE_PATTERNS + SINGLE_PATTERNS), re.IGNORECASE)

DEADLINE_KEYWORDS = [
    "deadline", "until", "no later than", "closes", "closing", "due",
    "last day", "end of applications",
    "προθεσμία", "προθεσμια", "λήξη", "ληξη", "μέχρι", "μεχρι", "έως", "εως",
]
OPEN_PERIOD_KEYWORDS = [
    "period", "from", "starts", "opens", "application period",
    "περίοδος", "περιοδος", "αιτήσεις από", "αιτησεις απο",
]
STALE_YEAR_RE = re.compile(r'\b(20\d{2})\s*[-/]\s*(\d{2,4})\b')


# ---- phase 1: discovery ----

def fetch_html(url: str) -> str | None:
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
        if r.status_code == 200:
            return r.text
    except Exception as e:
        log.warning(f"fetch_html failed for {url}: {e}")
    return None


def fetch_sitemap_urls(base_url: str) -> list[str]:
    parsed = urlparse(base_url)
    base = f"{parsed.scheme}://{parsed.netloc}"
    for path in ("/sitemap.xml", "/sitemap_index.xml", "/wp-sitemap.xml"):
        xml = fetch_html(base + path)
        if not xml:
            continue
        try:
            root = fromstring(xml.encode("utf-8"))
            urls = [el.text.strip() for el in root.iter() if el.tag.endswith("loc") and el.text]
            if urls:
                return urls[:200]  # cap — large sitemaps aren't worth fully walking
        except Exception:
            continue
    return []


def find_doc_links(html: str, base_url: str) -> list[str]:
    try:
        soup = BeautifulSoup(html, "html.parser")
        links = []
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if re.search(r'\.(pdf|docx?)(\?|$)', href, re.IGNORECASE):
                links.append(urljoin(base_url, href))
        return list(dict.fromkeys(links))
    except Exception as e:
        log.warning(f"find_doc_links failed: {e}")
        return []


def extract_pdf_text(url: str) -> str | None:
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
        if r.status_code != 200:
            return None
        import pdfplumber
        text = []
        with pdfplumber.open(io.BytesIO(r.content)) as pdf:
            for page in pdf.pages[:15]:
                t = page.extract_text()
                if t:
                    text.append(t)
        return "\n".join(text) if text else None
    except Exception as e:
        log.warning(f"extract_pdf_text failed for {url}: {e}")
        return None


def extract_docx_text(url: str) -> str | None:
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
        if r.status_code != 200:
            return None
        from docx import Document
        doc = Document(io.BytesIO(r.content))
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    except Exception as e:
        log.warning(f"extract_docx_text failed for {url}: {e}")
        return None


def gather_text_pool(homepage_url: str) -> list[dict]:
    """Returns a list of {url, text} for the homepage, subpages, sitemap
    entries, and every PDF/DOCX discovered along the way."""
    pool = []
    visited_pages = set()
    to_visit = [homepage_url]

    parsed = urlparse(homepage_url)
    base = f"{parsed.scheme}://{parsed.netloc}"
    to_visit += [base.rstrip("/") + s for s in SUBPAGE_SUFFIXES]
    to_visit += fetch_sitemap_urls(homepage_url)
    to_visit = list(dict.fromkeys(to_visit))

    doc_urls = set()
    for url in to_visit:
        if url in visited_pages:
            continue
        visited_pages.add(url)
        html = fetch_html(url)
        time.sleep(REQUEST_DELAY)
        if not html:
            continue
        text = BeautifulSoup(html, "html.parser").get_text(separator="\n")
        pool.append({"url": url, "text": text})
        for doc_url in find_doc_links(html, url):
            doc_urls.add(doc_url)

    for doc_url in doc_urls:
        if doc_url.lower().endswith(".docx") or doc_url.lower().endswith(".doc"):
            text = extract_docx_text(doc_url)
        else:
            text = extract_pdf_text(doc_url)
        time.sleep(REQUEST_DELAY)
        if text:
            pool.append({"url": doc_url, "text": text})

    return pool


# ---- phase 2: extraction ----

def extract_candidates(pool: list[dict]) -> list[dict]:
    candidates = []
    for entry in pool:
        url, text = entry["url"], entry["text"]
        lines = text.split("\n")
        for i, line in enumerate(lines):
            for m in ANY_DATE_RE.finditer(line):
                raw = m.group().strip()
                start = max(0, i - 2)
                end = min(len(lines), i + 3)
                context = " ".join(l.strip() for l in lines[start:end] if l.strip())[:400]
                ctx_lower = context.lower()
                is_range = bool(RANGE_RE.fullmatch(raw)) or bool(RANGE_RE.search(raw))
                candidates.append({
                    "raw": raw,
                    "is_range": is_range,
                    "context": context,
                    "source_url": url,
                    "near_deadline_kw": any(k in ctx_lower for k in DEADLINE_KEYWORDS),
                    "near_open_kw": any(k in ctx_lower for k in OPEN_PERIOD_KEYWORDS),
                    "stale_year_mention": _has_stale_year_mention(context),
                })
    return candidates


def _has_stale_year_mention(context: str) -> bool:
    for m in STALE_YEAR_RE.finditer(context):
        y1 = int(m.group(1))
        if y1 not in TARGET_YEARS and y1 + 1 not in TARGET_YEARS:
            return True
    return False


def _years_in(raw: str) -> list[int]:
    return [int(y) for y in re.findall(r'\b(20\d{2})\b', raw)]


# ---- phase 3: filtering ----

def filter_candidates(candidates: list[dict]) -> list[dict]:
    out = []
    for c in candidates:
        years = _years_in(c["raw"])
        if not years or not any(y in TARGET_YEARS for y in years):
            continue
        if c["stale_year_mention"]:
            continue
        out.append(c)
    return out


# ---- phase 4: decision (rule-based default; LLM optional, not wired here) ----

def _end_of_range(raw: str) -> str:
    """Best-effort: return the tail (end date) portion of a range string,
    used when picking a deadline out of a 'period' range."""
    parts = re.split(r'[-–—]', raw)
    return parts[-1].strip() if len(parts) > 1 else raw


def decide_deadline_rule_based(candidates: list[dict]) -> dict | None:
    if not candidates:
        return None
    # Priority 1: explicit deadline-keyword candidates (prefer non-range, or
    # the end of a range if it is one).
    deadline_tagged = [c for c in candidates if c["near_deadline_kw"]]
    # Priority 2: a period/range candidate near an open-period keyword — the
    # end of that range is the implicit deadline (e.g. "Application period:
    # April 1 - June 30, 2026").
    period_tagged = [c for c in candidates if c["is_range"] and c["near_open_kw"]]
    # Priority 3: any range candidate at all (take its end).
    any_range = [c for c in candidates if c["is_range"]]
    # Priority 4: anything left.
    pool = deadline_tagged or period_tagged or any_range or candidates
    chosen = pool[0]
    value = _end_of_range(chosen["raw"]) if chosen["is_range"] else chosen["raw"]
    return {
        "deadline": value,
        "deadline_source_url": chosen["source_url"],
        "deadline_context": chosen["context"],
        "found_via": (
            "deadline_keyword" if chosen in deadline_tagged else
            "period_range" if chosen in period_tagged else
            "range" if chosen in any_range else "fallback"
        ),
    }


def decide_with_llm(programme: dict, candidates: list[dict], graph_config, SmartScraperGraph) -> dict | None:
    """Optional: constrained LLM decision over the literal candidate list
    only (never raw page text). Left unwired by default — pass in your own
    graph_config / SmartScraperGraph (e.g. from scraper.py's get_graph_config)
    to enable. Not used unless explicitly called."""
    if not candidates:
        return None
    listing = "\n".join(
        f"{i+1}. \"{c['raw']}\" — context: \"{c['context']}\" — source: {c['source_url']}"
        for i, c in enumerate(candidates)
    )
    prompt = (
        f"Programme: {programme.get('name_en')}\n\n"
        "Below is a numbered list of date candidates extracted from the programme's "
        "website and documents. Pick the ONE candidate that is the application "
        "DEADLINE (final date to submit an application). Respond with JSON: "
        '{"index": <number from the list, or null if none is a deadline>}.\n\n'
        f"{listing}"
    )
    result = SmartScraperGraph(prompt=prompt, source=f"data:text/plain,{prompt}", config=graph_config).run()
    try:
        idx = result.get("index") if isinstance(result, dict) else None
        if idx and 1 <= idx <= len(candidates):
            c = candidates[idx - 1]
            value = _end_of_range(c["raw"]) if c["is_range"] else c["raw"]
            return {
                "deadline": value,
                "deadline_source_url": c["source_url"],
                "deadline_context": c["context"],
                "found_via": "llm_choice",
            }
    except Exception as e:
        log.warning(f"decide_with_llm parse failed: {e}")
    return None


# ---- orchestration ----

def process_programme(prog: dict) -> dict:
    log.info(f"Processing {prog['id']}: {prog.get('name_en','')[:60]}")
    pool = gather_text_pool(prog["website"])
    candidates = extract_candidates(pool)
    filtered = filter_candidates(candidates)
    decision = decide_deadline_rule_based(filtered)

    out = {
        "id": prog["id"],
        "name_en": prog.get("name_en"),
        "website": prog.get("website"),
        "candidates_found": len(filtered),
        "needs_review": decision is None,
        "scraped_at": datetime.now(timezone.utc).isoformat(),
    }
    if decision:
        out.update(decision)
    else:
        out.update({"deadline": None, "deadline_source_url": None,
                    "deadline_context": None, "found_via": None})
    return out


def load_programmes(ids: set[str] | None) -> list[dict]:
    progs = []
    with open(CSV_PATH, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            if ids and row["id"] not in ids:
                continue
            if not row.get("website"):
                continue
            progs.append(row)
    return progs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ids", type=str, default=None, help="comma-separated programme ids to limit to")
    ap.add_argument("--workers", type=int, default=1)
    args = ap.parse_args()

    ids = set(args.ids.split(",")) if args.ids else None
    progs = load_programmes(ids)
    log.info(f"Loaded {len(progs)} programmes")

    OUT_DIR.mkdir(exist_ok=True)
    results = []
    for prog in progs:
        try:
            results.append(process_programme(prog))
        except Exception as e:
            log.error(f"Failed on {prog['id']}: {e}")
            results.append({"id": prog["id"], "error": str(e), "needs_review": True})

    out_path = OUT_DIR / "deadline_results.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    log.info(f"Wrote {out_path}")


if __name__ == "__main__":
    main()

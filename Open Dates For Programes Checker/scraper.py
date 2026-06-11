"""
Scrapes MSc programme websites and collects application dates,
deadlines and application links.

Outputs are written to the results/ folder and can be loaded
by the dashboard.
"""

import json
import time
import csv
import os
import sys
import logging
import argparse
import requests
import re
import signal
import pdfplumber
import io
from datetime import datetime, timezone
from docx import Document
from pathlib import Path
from urllib.parse import urlparse, urljoin
from concurrent.futures import ThreadPoolExecutor, as_completed
from dotenv import load_dotenv
from xml.etree.ElementTree import fromstring
from bs4 import BeautifulSoup


load_dotenv()

# Configuration

DEFAULT_CSV     = "programmes-api_msc-programmes.csv"
DEFAULT_MODEL   = "openai/gpt-4.1-nano"
OUTPUT_DIR      = Path("results")
OUTPUT_DIR.mkdir(exist_ok=True)
ERRORS_JSON = OUTPUT_DIR / "errors.json"

PASS1_JSON      = OUTPUT_DIR / "pass1_results.json"
PASS2_JSON      = OUTPUT_DIR / "pass2_results.json"
RESULTS_JSON    = OUTPUT_DIR / "results.json"
RESULTS_CSV     = OUTPUT_DIR / "results.csv"
LOG_FILE        = OUTPUT_DIR / "scraper.log"

MAX_RETRIES     = 3
MAX_RETRIES_PER_DOMAIN = 15
MAX_SITEMAP_ATTEMPTS = 15
MAX_SUFFIX_ATTEMPTS = 50
RETRY_DELAY     = 4
REQUEST_DELAY   = 1.0
DEFAULT_WORKERS = 3

_sitemap_cache: dict = {}

DOMAIN_ANNOUNCEMENT_HUBS = {
    "eap.gr": "https://www.eap.gr/en/tag/invitation-en/",
}

_hub_cache: dict = {}

# ---- date regex patterns ----

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

DATE_PATTERNS = [
    rf'\b\d{{1,2}}[/\-\.]\d{{1,2}}[/\-\.]\d{{2,4}}\b',        # 15/06/2025 or 15-06-25
    rf'\b\d{{4}}[/\-\.]\d{{1,2}}[/\-\.]\d{{1,2}}\b',           # 2025-06-15
    rf'\b\d{{1,2}}\s+{_MONTHS_EN}\s+\d{{4}}\b',                 # 15 June 2025
    rf'\b{_MONTHS_EN}\s+\d{{1,2}},?\s+\d{{4}}\b',               # June 15, 2025
    rf'\b\d{{1,2}}\s+{_MONTHS_EN_SHORT}\.?\s+\d{{4}}\b',        # 15 Jun 2025
    rf'\b\d{{1,2}}(?:st|nd|rd|th)\s+(?:of\s+)?{_MONTHS_EN}\s+\d{{4}}\b',  # 15th of June 2025
    rf'\b\d{{1,2}}\s+{_MONTHS_GR}\s+\d{{4}}\b',                 # 15 Ιουνίου 2025
    rf'\b{_MONTHS_GR}\s+\d{{4}}\b',                              # Ιούνιος 2025
]

DATE_RE = re.compile("|".join(DATE_PATTERNS), re.IGNORECASE)

# keywords that signal a date is application-related
CONTEXT_KEYWORDS = [
    "deadline", "προθεσμ", "application", "αιτησ", "αίτησ", "αιτήσε",
    "submit", "υποβολ", "open", "close", "📅", "έως", "εως", "μέχρι",
    "until", "by", "due", "from", "start", "begin", "end",
    "εγγραφ", "registration", "admission", "call", "πρόσκλ", "προσκλ",
    "ανακοίν", "ανακοιν", "apply", "intake", "semester",
]

# Common admissions/application paths
SUBPAGE_SUFFIXES = [
    # English
    "/category/announcements", "/category/news", "/category/calls",
    "/category/admission", "/category/applications","/admissions", "/admission", "/apply", "/apply-now", "/applications", "/application",
    "/application-form", "/application-process", "/application-portal",
    "/how-to-apply", "/how-to-apply-for", "/apply-here",
    "/prospective-students", "/prospective", "/future-students",
    "/entry-requirements", "/requirements", "/eligibility",
    "/registration", "/register", "/enroll", "/enrolment", "/enrollment",
    "/postgraduate", "/postgraduate-admissions", "/postgraduate-applications",
    "/masters", "/msc", "/graduate", "/graduate-admissions",
    "/fees-and-funding", "/funding", "/scholarships", "/scholarship",
    "/deadlines", "/dates", "/intake", "/open-days", "/academic-calendar",
    "/announcements", "/news", "/calls", "/announcement", "/call",
    "/news-events", "/news-and-events", "/latest-news",
    "/events", "/updates", "/notices", "/notice",
    "/open-calls", "/call-for-applications", "/call-for-interest",
    "/category/announcements", "/category/news", "/category/calls",
    "/category/admission", "/category/applications",

    # Greek (with accents)
    "/αιτηση", "/αίτηση", "/αιτησεις", "/αιτήσεις",
    "/εγγραφη", "/εγγραφή", "/εγγραφες", "/εγγραφές",
    "/προθεσμια", "/προθεσμία", "/προθεσμιες", "/προθεσμίες",
    "/προκηρυξη", "/προκήρυξη", "/προκηρυξεις", "/προκηρύξεις",
    "/ανακοινωση", "/ανακοίνωση", "/ανακοινωσεις", "/ανακοινώσεις",
    "/υποβολη-αιτησης", "/υποβολή-αίτησης",
    "/μεταπτυχιακο", "/μεταπτυχιακό",
    "/μεταπτυχιακα", "/μεταπτυχιακά",
    "/εισαγωγη", "/εισαγωγή",
    "/υποψηφιοι", "/υποψήφιοι",
    "/νεα", "/νέα", "/εκδηλωσεις", "/εκδηλώσεις",
    "/ενδιαφερον", "/ενδιαφέρον",
    "/προσκληση", "/πρόσκληση",

    # Greeklish
    "/aitisi", "/aitisi-eggrafi", "/aithsh", "/aithseis",
    "/eggrafi", "/eggrafes", "/egrafes", "/eggrafe",
    "/foitisi", "/foitites", "/foithsh", "/phoitisi", "/phitisi", "/fitiths",
    "/metaptixiako", "/metaptixiaka", "/metaptyxiako", "/metaptyxiaka",
    "/prokiriksi", "/prokiryxi", "/prokiriksh",
    "/anakoinosi", "/anakoinoseis", "/anakoinwsh",
    "/prothesmia", "/prothesmies",
    "/eisagogi", "/eisagwgh",
    "/ypopsifioi", "/ypopsifioi-foitites", "/nea", "/anakoinoseis", "/anakoinwseis",
    "/prokirixeis", "/prokiryxeis",
    "/ekdiloseis", "/prosklisi",
]

PDF_KEYWORDS = [
    # English
    "apply", "application", "applications", "admission", "admissions",
    "deadline", "deadlines", "dates", "open", "opening", "intake",
    "brochure", "prospectus", "announcement", "call", "calls",
    "registration", "enroll", "enrolment", "requirements", "eligibility",
    "scholarship", "funding", "academic-calendar","admission procedure", "admission calendar", "applications are open",
    "application calls", "application period", "application submission",
    "call for admission", "application deadline", "application for admission",
    "new application deadline", "application deadline extension",
    "online application", "extended deadline", "applications are now open",
    "applications until", "announcement for applications",
    "applications opening date", "application closing date",
    "application deadline until", "opening of applications",
    "new call", "deadline for submitting applications",
    "deadline for applications",

    # Greek (with accents)
    "αιτηση", "αίτηση", "αιτησεις", "αιτήσεις",
    "εγγραφη", "εγγραφή", "εγγραφες", "εγγραφές",
    "προθεσμια", "προθεσμία", "προκηρυξη", "προκήρυξη",
    "ανακοινωση", "ανακοίνωση", "μεταπτυχιακο", "μεταπτυχιακό",
    "υποβολη", "υποβολή", "εισαγωγη", "εισαγωγή", "προκήρυξη αιτήσεων", "υποβολή αίτησης",
    "προθεσμία υποβολής αιτήσεων",
    "παράταση προθεσμίας υποβολής αιτήσεων",
    "νέα παράταση προθεσμίας υποβολής αιτήσεων",
    "πρόσκληση εκδήλωσης ενδιαφέροντος",

    # Greeklish
    "aitisi", "aithsh", "aithseis", "eggrafi", "eggrafes",
    "foitisi", "foithsh", "phoitisi", "phitisi", "fitiths",
    "metaptixiako", "metaptyxiako", "prokiriksi", "prokiryxi",
    "anakoinosi", "anakoinwsh", "prothesmia", "eisagogi", "eisagwgh",
]

ANNOUNCEMENT_KEYWORDS = [
    # English
    "invitation","call-for-admission","call for applications", "call for admission", "application deadline",
    "deadline extension", "extended deadline", "applications open",
    "applications are open", "applications now open", "admission announcement",
    "new deadline", "application period", "intake",
    # Greek
    "προκήρυξη", "προκηρυξη", "παράταση", "παραταση",
    "αιτήσεις", "αιτησεις", "υποβολή αίτησης", "υποβολη αιτησης",
    "ανακοίνωση", "ανακοινωση",
    # Greeklish
    "prokiriksi", "prokiryxi", "paratasi", "anakoinosi",
]

DEADLINE_CSS_CLASSES = [
    "ew-deadline-text", "deadline", "application-deadline",
    "closing-date", "submission-deadline",
]

# Logging

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)
logging.getLogger("scrapegraphai").setLevel(logging.ERROR)
logging.getLogger("playwright").setLevel(logging.ERROR)
logging.getLogger("httpx").setLevel(logging.ERROR)
logging.getLogger("httpcore").setLevel(logging.ERROR)
#logging.getLogger("openai").setLevel(logging.ERROR)

# Shutdown handling

_shutdown = False

def _handle_signal(signum, frame):
    global _shutdown
    _shutdown = True
    log.info("Shutdown requested. Saving progress...")

signal.signal(signal.SIGINT,  _handle_signal)
signal.signal(signal.SIGTERM, _handle_signal)

def now_utc() -> str:
    return datetime.now(timezone.utc).isoformat()

# Extraction prompts

PASS1_PROMPT = """
You are extracting structured admissions data from a university MSc programme webpage. 
Return ONLY a valid JSON object. No markdown, no code fences, no explanation.

{
  "application_open_date": "Date when applications open, any format found, or null",
  "application_deadline": "Application deadline / closing date. Look for dates in ANY format including prose sentences like 'submit by 31st August 2026' or 'applications accepted until March 2026'. If multiple rounds, list all as one string. Or null.",
  "intake_semester": "e.g. Fall 2025, September 2025, Academic Year 2025-26, or null",
  "apply_button_url": "The full absolute URL of any Apply / Apply Now / Apply Here / Start Application / Admissions button or link. Resolve relative URLs using the page domain. Return null if none found.",
  "has_dates_on_page": true,
  "application_status": "open / closed / rolling / coming_soon / not_mentioned",
  "notes": "Any useful admissions note: rolling admissions, closed, contact department, waitlist, etc. Or null."
}

Rules:
- has_dates_on_page must be the boolean true if at least one real calendar date is found for open/deadline, otherwise false.
- application_status values: open, closed, rolling, coming_soon, not_mentioned
- For apply_button_url look for <a> tags with text: Apply, Apply Now, Apply Here, Start Application, Εγγραφή, Αίτηση, Υποβολή Αίτησης. If URL is relative prepend the page domain.
- Return null for any field with no data. Never return "NA" or "N/A".
- Extract dates even when written in plain sentences or paragraphs, not just labelled fields.
- has_dates_on_page must be true if ANY date related to applications or deadlines appears anywhere on the page, even in prose.
- Today is June 2026. Only extract dates that are in 2026 or later. Ignore any dates from 2025 or earlier.
- When multiple dates exist, prioritise dates labelled as application open, deadline, closing date, or submission — NOT results announcements, events, or publication dates.
"""

PASS2_PROMPT = """
You are extracting application timeline data from a university application portal page.
Return ONLY a valid JSON object. No markdown, no code fences, no explanation.

{
  "application_open_date": "Date when applications open or null",
  "application_deadline": "Application deadline. Look for dates in ANY format including prose sentences like 'submit by 31st August 2026' or 'applications accepted until March 2026'. Or null.",
  "intake_semester": "e.g. Fall 2025 or null",
  "portal_name": "Portal system name: Apply Texas / Slate / Embark / Common App / custom / or null",
  "requires_login": true,
  "application_status": "open / closed / rolling / coming_soon / not_mentioned",
  "notes": "Any timeline or admissions note or null"
}

Rules:
- requires_login must be boolean true if page is a login wall or auth redirect, otherwise false.
- Look for: deadline, opens, closes, due date, submit by, Προθεσμία, Αιτήσεις.
- Return null for any field with no data. Never return "NA" or "N/A".
- Extract dates even when written in plain sentences or paragraphs, not just labelled fields.
- Today is June 2026. Only extract dates that are in 2026 or later. Ignore any dates from 2025 or earlier.
- When multiple dates exist, prioritise dates labelled as application open, deadline, closing date, or submission — NOT results announcements, events, or publication dates.

"""

# CSV loading

def load_programmes(csv_path: str, active_only: bool) -> list:
    path = Path(csv_path)
    if not path.exists():
        log.error(f"CSV not found: {csv_path}")
        sys.exit(1)

    programmes = []
    with open(path, encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if active_only and row.get("is_archived") == "true":
                continue
            website = row.get("website", "").strip()
            if not website:
                continue

            has_fees = row.get("has_tuition_fees") == "true"
            fee_from = row.get("tuition_fees_from", "").strip()
            fee_to   = row.get("tuition_fees_to", "").strip()
            if has_fees and fee_from and fee_to:
                tuition = f"EUR {fee_from}" if fee_from == fee_to else f"EUR {fee_from}-{fee_to}"
            elif has_fees:
                tuition = "Yes (amount unknown)"
            else:
                tuition = "Free"

            programmes.append({
                "id":                  row.get("id", ""),
                "url":                 website,
                "name_en":             row.get("name_en", "").strip(),
                "name_gr":             row.get("name_gr", "").strip(),
                "university":          row.get("department.university.name_en", "").strip(),
                "university_gr":       row.get("department.university.name_gr", "").strip(),
                "department":          row.get("department.name_en", "").strip(),
                "city":                row.get("city_en", "").strip(),
                "topics":              row.get("topics", "").strip(),
                "languages":           row.get("languages", "").strip(),
                "ects":                row.get("ects", "").strip(),
                "semesters":           row.get("semesters", "").strip(),
                "tuition":             tuition,
                "study_modes":         row.get("study_modes", "").replace(";", " / "),
                "email":               row.get("email", "").strip(),
                "phone":               row.get("phone", "").strip(),
                "apply_status_db":     row.get("apply_status", "").strip(),
                "is_archived":         row.get("is_archived") == "true",
                "scholarship_offered": row.get("scholarship_offered") == "true",
                "university_image_url":row.get("department.university.image_url", "").strip(),
            })

    log.info(f"Loaded {len(programmes)} programmes from {csv_path}")
    return programmes

# Scraping helpers

def get_graph_config(model: str) -> dict:
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        log.error("OPENAI_API_KEY not set. Add to .env file.")
        sys.exit(1)
    return {
        "llm": {"api_key": api_key, "model": model},
        "headless": True,
        "verbose": False,
    }


def import_scraper():
    try:
        from scrapegraphai.graphs import SmartScraperGraph
        return SmartScraperGraph
    except ImportError:
        log.error("Run: pip install scrapegraphai playwright && playwright install chromium")
        sys.exit(1)


def strip_json_fences(text: str) -> str:
    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    return text.strip()


def unwrap_content(raw: dict) -> dict:
    """Some model responses are nested under 'content'."""
    if isinstance(raw, dict) and "content" in raw and len(raw) == 1:
        content = raw["content"]
        if isinstance(content, str):
            try:
                cleaned = strip_json_fences(content)
                # fix unescaped inner quotes if needed
                return json.loads(cleaned)
            except json.JSONDecodeError:
                # try extracting just the JSON object
                match = re.search(r'\{.*\}', content, re.DOTALL)
                if match:
                    try:
                        return json.loads(match.group())
                    except Exception:
                        pass
        elif isinstance(content, dict):
            return content
    return raw


def clean_result(raw: dict) -> dict:
    """Convert placeholder values to None and normalize lists to strings."""
    cleaned = {}
    for k, v in raw.items():
        if isinstance(v, list):
            v = ", ".join(str(i) for i in v if i)
        if isinstance(v, str) and v.strip().upper() in ("NA", "N/A", "NONE", "NULL", ""):
            cleaned[k] = None
        else:
            cleaned[k] = v
    return cleaned


def scrape_url(url: str, prompt: str, graph_config: dict, SmartScraperGraph) -> dict:
    last_error = None
    for attempt in range(1, MAX_RETRIES + 1):
        if _shutdown:
            return {"status": "cancelled", "error": "shutdown"}
        try:
            graph = SmartScraperGraph(prompt=prompt, source=url, config=graph_config)
            raw = graph.run()

            if isinstance(raw, dict):
                raw = unwrap_content(raw)
                if raw is None:
                    raise ValueError("unwrap_content returned None")
                if list(raw.keys()) == ["result"] and isinstance(raw.get("result"), dict):
                    raw = raw["result"]
                raw = clean_result(raw)
                raw["status"] = "ok"
                return raw
            elif isinstance(raw, str):
                parsed = json.loads(strip_json_fences(raw))
                parsed = clean_result(parsed)
                parsed["status"] = "ok"
                return parsed
            elif raw is None:
                raise ValueError("Scraper returned None")
            else:
                raise ValueError(f"Unexpected type: {type(raw)}")

        except json.JSONDecodeError as e:
            last_error = f"JSON parse error: {e}"
            log.warning(f"[attempt {attempt}/{MAX_RETRIES}] JSON error on {url}: {e}")
        except Exception as e:
            last_error = str(e)
            # Try to recover valid JSON from LangChain "Invalid json output" errors
            if "Invalid json output:" in last_error:
                try:
                    err = last_error.split("Invalid json output:", 1)[1].strip()
                    err = err.split("\nFor troubleshooting")[0].strip()
                    # Strip outer {"content": "..."} wrapper — inner JSON has unescaped quotes
                    prefix = '{"content": "'
                    if err.startswith(prefix):
                        err = err[len(prefix):]
                        if err.endswith('"}'):
                            err = err[:-2]
                    inner = json.loads(err)
                    if isinstance(inner, dict):
                        inner = clean_result(inner)
                        inner["status"] = "ok"
                        log.info(f"  Recovered JSON from error on {url}")
                        return inner
                except Exception:
                    pass
            log.warning(f"[attempt {attempt}/{MAX_RETRIES}] Error on {url}: {e}")

        if attempt < MAX_RETRIES:
            if "429" in str(last_error) or "rate_limit" in str(last_error):
                wait = 60
                log.info(f"  Rate limit hit, waiting {wait}s before retry...")
                time.sleep(wait)
            else:
                time.sleep(RETRY_DELAY * attempt)

    log.error(f"FAILED after {MAX_RETRIES} attempts: {url} — {last_error}")
    return {"status": "error", "error": last_error}


def try_subpages(base_url: str, prompt: str, graph_config: dict, SmartScraperGraph) -> dict:
    parsed = urlparse(base_url)
    base = f"{parsed.scheme}://{parsed.netloc}"

    sitemap_urls = fetch_sitemap_urls(base_url)
    suffix_urls = [base.rstrip("/") + s for s in SUBPAGE_SUFFIXES]

    all_urls = list(dict.fromkeys(sitemap_urls + suffix_urls))

    sitemap_count = 0
    suffix_count = 0
    total_count = 0
    consecutive_failures = 0

    for url in all_urls:
        if _shutdown:
            break
        if url == base_url:
            continue
        if total_count >= MAX_RETRIES_PER_DOMAIN:
            log.info(f"  Domain cap reached for {base}, stopping")
            break
        if consecutive_failures >= 5:
            log.info(f"  Too many consecutive failures for {base}, giving up")
            break

        is_sitemap_url = url in sitemap_urls
        if is_sitemap_url and sitemap_count >= MAX_SITEMAP_ATTEMPTS:
            continue
        if not is_sitemap_url and suffix_count >= MAX_SUFFIX_ATTEMPTS:
            continue

        log.info(f"  Trying sub-page: {url}")
        result = scrape_url(url, prompt, graph_config, SmartScraperGraph)

        if is_sitemap_url:
            sitemap_count += 1
        else:
            suffix_count += 1
        total_count += 1

        if result.get("status") == "ok" and result.get("has_dates_on_page") is True:
            log.info(f"  Found dates on sub-page: {url}")
            result["found_on_subpage"] = url
            return result
        if result.get("status") == "error":
            consecutive_failures += 1
        else:
            consecutive_failures = 0
        time.sleep(0.5)

    return {}

#RAW_HTML_HELPERS
def fetch_html(url: str) -> str | None:
    try:
        headers = {"User-Agent": "Mozilla/5.0 (compatible; MScScraper/1.0)"}
        response = requests.get(url, timeout=10, headers=headers)
        response.raise_for_status()
        return response.text
    except Exception as e:
        log.warning(f"fetch_html failed for {url}: {e}")
        return None
    
def fetch_sitemap_urls(base_url: str) -> list:
    parsed = urlparse(base_url)
    base = f"{parsed.scheme}://{parsed.netloc}"
    sitemap_url = base.rstrip("/") + "/sitemap.xml"
    
    keywords = ["apply", "admission", "deadline", "application", "dates", "registration", "enroll", "αιτηση", "αιτήση", "εγγραφη", "εγγραφή", "προθεσμια", "προθεσμία", "προκηρυξη", "προκήρυξη"]
    
    if sitemap_url in _sitemap_cache:
        return _sitemap_cache[sitemap_url]
    html = fetch_html(sitemap_url)
    if not html:
        _sitemap_cache[sitemap_url] = []
        return []
    
    try:
        root = fromstring(html)
        locs = [el.text.strip() for el in root.iter() if el.tag.endswith("loc") and el.text]
        
        # handle sitemap index
        all_urls = []
        for loc in locs:
            if loc.endswith(".xml"):
                child = fetch_html(loc)
                if child:
                    child_root = fromstring(child)
                    all_urls += [el.text.strip() for el in child_root.iter() if el.tag.endswith("loc") and el.text]
            else:
                all_urls.append(loc)
        
        filtered = [u for u in all_urls if any(k in u.lower() for k in keywords)]
        # Sort newest first — prefer URLs containing higher years (2026 > 2025 > 2022 etc.)
        def _url_year(url):
            years = re.findall(r'/(20\d{2})[-/]', url)
            return max(int(y) for y in years) if years else 0
        filtered.sort(key=_url_year, reverse=True)
        _sitemap_cache[sitemap_url] = filtered[:10]
        return filtered[:10]

    except Exception as e:
        log.warning(f"fetch_sitemap_urls failed for {base_url}: {e}")
        _sitemap_cache[sitemap_url] = []
        return []
    
def find_pdf_links(html: str, base_url: str) -> list:
    keywords = PDF_KEYWORDS
    try:
        soup = BeautifulSoup(html, "html.parser")
        links = []
        for a in soup.find_all("a", href=True):
            href = a["href"].strip()
            if href.lower().endswith(".pdf") or href.lower().endswith(".docx"):
                absolute = urljoin(base_url, href)
                link_text = a.get_text().lower()
                href_lower = href.lower()
                if any(k in href_lower or k in link_text for k in keywords):
                    links.append(absolute)
        return links[:5]
    except Exception as e:
        log.warning(f"find_pdf_links failed: {e}")
        return []
    
def find_announcement_links(html: str, base_url: str) -> list:
    try:
        soup = BeautifulSoup(html, "html.parser")
        links = []
        for a in soup.find_all("a", href=True):
            text = a.get_text().strip().lower()
            href = a["href"].strip()
            if not href or href.startswith("#"):
                continue
            if any(k in text for k in ANNOUNCEMENT_KEYWORDS):
                absolute = urljoin(base_url, href)
                if absolute not in links:
                    links.append(absolute)
        return links[:5]
    except Exception as e:
        log.warning(f"find_announcement_links failed: {e}")
        return []
    
def extract_pdf_text(url: str) -> str | None:
    try:
        headers = {"User-Agent": "Mozilla/5.0 (compatible; MScScraper/1.0)"}
        response = requests.get(url, timeout=15, headers=headers)
        response.raise_for_status()
        with pdfplumber.open(io.BytesIO(response.content)) as pdf:
            text = "\n".join(
                page.extract_text() or "" for page in pdf.pages
            )
        return text[:3000] if text.strip() else None
    except Exception as e:
        log.warning(f"extract_pdf_text failed for {url}: {e}")
        return None
    
def extract_docx_text(url: str) -> str | None:
    try:
        headers = {"User-Agent": "Mozilla/5.0 (compatible; MScScraper/1.0)"}
        response = requests.get(url, timeout=15, headers=headers)
        response.raise_for_status()
        doc = Document(io.BytesIO(response.content))
        text = "\n".join(para.text for para in doc.paragraphs if para.text.strip())
        return text[:3000] if text.strip() else None
    except Exception as e:
        log.warning(f"extract_docx_text failed for {url}: {e}")
        return None

def extract_dates_with_context(text: str, source_label: str) -> list[dict]:
    results = []
    seen_dates = set()
    lines = text.split("\n")

    for i, line in enumerate(lines):
        for match in DATE_RE.finditer(line):
            raw = match.group().strip()
            if raw in seen_dates:
                continue
            seen_dates.add(raw)

            # grab surrounding lines for context
            start = max(0, i - 2)
            end = min(len(lines), i + 3)
            context = " ".join(lines[start:end]).strip()

            ctx_lower = context.lower()
            has_keyword = any(k in ctx_lower for k in CONTEXT_KEYWORDS)

            results.append({
                "raw": raw,
                "context": context[:300],
                "source": source_label,
                "has_keyword": has_keyword,
            })

    return results

def validate_dates(dates: list[dict]) -> list[dict]:
    from dateutil import parser as dateparser
    valid = []
    for d in dates:
        try:
            parsed = dateparser.parse(d["raw"], dayfirst=True, fuzzy=False)
            if parsed is None:
                continue
            if parsed.year < 2026 or parsed.year > 2027:
                continue
            d["parsed_iso"] = parsed.strftime("%Y-%m-%d")
            d["year"] = parsed.year
            valid.append(d)
        except Exception:
            continue  # discard unparseable strings
    return valid


def build_date_list(prog: dict, dates: list[dict]) -> str:
    header = f"TARGET PROGRAMME: {prog['name_en']}"
    if prog.get("name_gr"):
        header += f" / {prog['name_gr']}"
    header += "\n\nDATES FOUND:\n"

    lines = []
    for i, d in enumerate(dates, 1):
        keyword_flag = " [KEYWORD MATCH]" if d.get("has_keyword") else ""
        lines.append(
            f"{i}. \"{d['raw']}\" — context: \"{d['context']}\" — "
            f"source: {d['source']}{keyword_flag}"
        )

    return header + "\n".join(lines)

def extract_date_contexts(html: str) -> str | None:
    """Extract date-adjacent and keyword-adjacent lines from raw HTML."""
    try:
        soup = BeautifulSoup(html, "html.parser")
        text = soup.get_text(separator="\n")

        context_keywords = [
            "deadline", "application", "admission", "apply", "open", "closing",
            "submit", "registration", "intake", "start date",
            "προθεσμία", "προθεσμια", "αιτήσεις", "αιτησεις",
            "υποβολή", "υποβολη", "εγγραφή", "εγγραφη",
        ]

        lines = text.split("\n")
        chunks = []
        seen = set()
        for i, line in enumerate(lines):
            line_s = line.strip()
            if not line_s:
                continue
            has_date = bool(DATE_RE.search(line_s))
            has_keyword = any(k in line_s.lower() for k in context_keywords)
            if has_date or has_keyword:
                start = max(0, i - 2)
                end = min(len(lines), i + 3)
                chunk = "\n".join(l.strip() for l in lines[start:end] if l.strip())
                if chunk and chunk not in seen:
                    seen.add(chunk)
                    chunks.append(chunk)
        return "\n---\n".join(chunks)[:3000] if chunks else None
    except Exception as e:
        log.warning(f"extract_date_contexts failed: {e}")
        return None
    

# Workers
def has_current_dates(result: dict) -> bool:
    """Returns True if result contains at least one date from 2026 or later."""
    for field in ["application_deadline", "application_open_date", "notes"]:
        val = result.get(field) or ""
        years = re.findall(r'\b(20\d{2})\b', str(val))
        if any(int(y) >= 2026 for y in years):
            return True
    return False

def worker_pass1(args):
    idx, total, prog, prompt, graph_config, SmartScraperGraph = args
    if _shutdown:
        return {"prog_id": prog["id"], "status": "cancelled", "scraped_at": now_utc()}

    log.info(f"[Pass1 {idx}/{total}] {prog['university']} — {prog['name_en'][:55]}")

    # Fetch raw HTML first
    html = fetch_html(prog["url"])

    # Check for domain-specific announcement hub (e.g. EAP posts all dates centrally)
    prog_domain = urlparse(prog["url"]).netloc
    hub_url = next((v for k, v in DOMAIN_ANNOUNCEMENT_HUBS.items() if k in prog_domain), None)
    if hub_url:
        if hub_url in _hub_cache:
            log.info(f"  Hub cache hit: {hub_url}")
            result = _hub_cache[hub_url].copy()
            result["prog_id"] = prog["id"]
            result["scraped_at"] = now_utc()
            result["found_in_announcement"] = hub_url
            return result
        log.info(f"  Domain hub: {hub_url}")
        hub_html = fetch_html(hub_url)
        hub_url_to_scrape = hub_url
        if hub_html:
            ann_links = find_announcement_links(hub_html, hub_url)
            if ann_links:
                log.info(f"  Following hub announcement: {ann_links[0]}")
                hub_url_to_scrape = ann_links[0]
        hub_result = scrape_url(hub_url_to_scrape, prompt, graph_config, SmartScraperGraph)
        if hub_result.get("has_dates_on_page") is True:
            hub_result["found_in_announcement"] = hub_url_to_scrape
            _hub_cache[hub_url] = hub_result
            result = hub_result.copy()
            result["prog_id"] = prog["id"]
            result["scraped_at"] = now_utc()
            return result

    # Extract and filter dates, then send structured list to AI
    if html:
        soup_text = BeautifulSoup(html, "html.parser").get_text(separator="\n")
        raw_dates = extract_dates_with_context(soup_text, prog["url"])
        valid_dates = validate_dates(raw_dates)
        keyword_dates = [d for d in valid_dates if d.get("has_keyword")]
        dates_to_use = keyword_dates if keyword_dates else valid_dates
    else:
        dates_to_use = []

    if dates_to_use:
        log.info(f"  Found {len(dates_to_use)} valid dates, building structured list")
        date_list = build_date_list(prog, dates_to_use)
        result = scrape_url(f"data:text/plain,{date_list}", prompt, graph_config, SmartScraperGraph)
        if not (result.get("has_dates_on_page") is True and has_current_dates(result)):
            result = scrape_url(prog["url"], prompt, graph_config, SmartScraperGraph)
    else:
        result = scrape_url(prog["url"], prompt, graph_config, SmartScraperGraph)
    # Detect external apply portal
    apply_url = result.get("apply_button_url")
    if apply_url and urlparse(apply_url).netloc != urlparse(prog["url"]).netloc:
        log.info(f"  External portal detected: {apply_url}")
        result["external_portal"] = apply_url

    # If no dates found, check document links (PDF/DOCX) on the page
    if html and result.get("status") == "ok" and not (result.get("has_dates_on_page") is True and has_current_dates(result)):
        doc_links = find_pdf_links(html, prog["url"])
        for doc_url in doc_links:
            log.info(f"  Trying document: {doc_url}")
            if doc_url.lower().endswith(".docx"):
                doc_text = extract_docx_text(doc_url)
            else:
                doc_text = extract_pdf_text(doc_url)
            if doc_text:
                doc_result = scrape_url(
                    f"data:text/plain,{doc_text}",
                    prompt, graph_config, SmartScraperGraph
                )
                if doc_result.get("has_dates_on_page") is True:
                    log.info(f"  Found dates in document: {doc_url}")
                    doc_result["found_in_pdf"] = doc_url
                    result = doc_result
                    break

    # If still no dates, follow announcement links on the page
    if html and result.get("status") == "ok" and not (result.get("has_dates_on_page") is True and has_current_dates(result)):
        announcement_links = find_announcement_links(html, prog["url"])
        for ann_url in announcement_links:
            if _shutdown:
                break
            log.info(f"  Trying announcement: {ann_url}")
            ann_result = scrape_url(ann_url, prompt, graph_config, SmartScraperGraph)
            if ann_result.get("has_dates_on_page") is True:
                log.info(f"  Found dates in announcement: {ann_url}")
                ann_result["found_in_announcement"] = ann_url
                result = ann_result
                break
            time.sleep(1.0)

    # No useful date info found, check sitemap and subpages
    if (result.get("status") == "ok"
            and not (result.get("has_dates_on_page") is True and has_current_dates(result))
            and not result.get("apply_button_url")):
        sub = try_subpages(prog["url"], prompt, graph_config, SmartScraperGraph)
        if sub.get("has_dates_on_page") is True:
            result = sub

    result["prog_id"]    = prog["id"]
    result["scraped_at"] = now_utc()
    time.sleep(REQUEST_DELAY)
    return result

def worker_pass2(args):
    idx, total, prog_id, source_url, apply_url, prompt, graph_config, SmartScraperGraph = args
    if _shutdown:
        return {"prog_id": prog_id, "status": "cancelled", "scraped_at": now_utc()}

    # If external portal (different domain), try sitemap first to find a better entry page
    from urllib.parse import urlparse
    src_domain  = urlparse(source_url).netloc.lstrip("www.")
    apply_domain = urlparse(apply_url).netloc.lstrip("www.")
    if apply_domain and apply_domain != src_domain:
        sitemap_hits = fetch_sitemap_urls(apply_url)
        if sitemap_hits:
            log.info(f"  External portal sitemap found {len(sitemap_hits)} relevant pages")
            apply_url = sitemap_hits[0]

    log.info(f"[Pass2 {idx}/{total}] {apply_url[:80]}")
    result = scrape_url(apply_url, prompt, graph_config, SmartScraperGraph)
    result["prog_id"]    = prog_id
    result["apply_url"]  = apply_url
    result["source_url"] = source_url
    result["scraped_at"] = now_utc()
    time.sleep(REQUEST_DELAY)
    return result

# Pass 1

def run_pass1(programmes, graph_config, SmartScraperGraph, workers, resume_ids):
    to_scrape = [p for p in programmes if p["id"] not in resume_ids]
    skipped = len(programmes) - len(to_scrape)
    if skipped:
        log.info(f"Resume: skipping {skipped} already-processed programmes")
    if not to_scrape:
        log.info("Pass 1: nothing new to scrape")
        return []

    log.info(f"Pass 1: {len(to_scrape)} programmes | {workers} workers")
    tasks = [
        (i+1, len(to_scrape), prog, PASS1_PROMPT, graph_config, SmartScraperGraph)
        for i, prog in enumerate(to_scrape)
    ]

    results = []
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {pool.submit(worker_pass1, t): t for t in tasks}
        try:
            for future in as_completed(futures):
                try:
                    results.append(future.result())
                except Exception as e:
                    task = futures[future]
                    log.error(f"Worker crash: {e}")
                    results.append({
                        "prog_id":    task[2]["id"],
                        "status":     "error",
                        "error":      str(e),
                        "scraped_at": now_utc(),
                    })
                if _shutdown:
                    log.info("Shutdown — cancelling remaining tasks...")
                    for f in futures:
                        f.cancel()
                    break
        except KeyboardInterrupt:
            for f in futures:
                f.cancel()
            pool.shutdown(wait=False, cancel_futures=True)

    return results

# Pass 2

def _needs_pass2(r: dict) -> bool:
    """True if pass2 should follow the apply URL for this pass1 result."""
    if r.get("has_dates_on_page") is not True:
        return True
    # Dates found but deadline is still missing — apply page may have more detail
    if not r.get("application_deadline"):
        return True
    return False

def run_pass2(pass1_results, prog_map, graph_config, SmartScraperGraph, workers, done_ids):
    candidates = [
        r for r in pass1_results
        if _needs_pass2(r)
        and str(r.get("apply_button_url") or "").startswith("http")
        and r.get("status") == "ok"
        and r.get("prog_id") not in done_ids
    ]

    if not candidates:
        log.info("Pass 2: no candidates")
        return []

    log.info(f"Pass 2: following apply links for {len(candidates)} programmes")
    tasks = [
        (i+1, len(candidates),
         r["prog_id"],
         prog_map.get(r["prog_id"], {}).get("url", ""),
         r.get("external_portal") or r["apply_button_url"],
         PASS2_PROMPT, graph_config, SmartScraperGraph)
        for i, r in enumerate(candidates)
    ]

    results = []
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {pool.submit(worker_pass2, t): t for t in tasks}
        try:
            for future in as_completed(futures):
                try:
                    results.append(future.result())
                except Exception as e:
                    log.error(f"Worker crash: {e}")
                    results.append({"status": "error", "error": str(e)})
                if _shutdown:
                    for f in futures:
                        f.cancel()
                    break
        except KeyboardInterrupt:
            for f in futures:
                f.cancel()
            pool.shutdown(wait=False, cancel_futures=True)

    return results

# Date extraction
def extract_date_from_notes(notes: str) -> str | None:
    if not notes:
        return None
    m = DATE_RE.search(notes)
    return m.group().strip() if m else None

# To more accurately find relevant dates 
def date_relevance_score(text) -> int:
    if not text:
        return 0
    if isinstance(text, list):
        text = " ".join(str(t) for t in text)
    text = str(text)
    years = re.findall(r'\b(20\d{2})\b', text)
    if not years:
        return 0
    max_year = max(int(y) for y in years)
    if max_year >= 2026:
        return 3
    elif max_year == 2025:
        return 2
    elif max_year >= 2023:
        return 1
    return 0

## BEST RECORDS CHANGED 1900
def best_record(records):
    return max(records, key=lambda r: (
        bool(r.get("found_in_pdf")),
        r.get("has_dates_on_page") is True,
        date_relevance_score(r.get("application_deadline") or ""),
        date_relevance_score(r.get("application_open_date") or ""),
        bool(r.get("application_deadline")),
        bool(r.get("application_open_date")),
        bool(r.get("notes")),
    ))

# Merge and export
def merge(programmes, pass1, pass2):
    
    p1_map = {}
    for r in pass1:
        pid = r["prog_id"]
        if pid not in p1_map:
            p1_map[pid] = r
        else:
            p1_map[pid] = best_record([p1_map[pid], r])
    p2_map = {r["prog_id"]: r for r in pass2 if r.get("prog_id")}

    rows = []
    for prog in programmes:
        pid = prog["id"]
        p1  = p1_map.get(pid, {})
        p2  = p2_map.get(pid, {})

        # Prefer PDF-sourced dates as they are more reliable than HTML
        if p1.get("found_in_pdf"):
            deadline  = p1.get("application_deadline") or p2.get("application_deadline")
            open_date = p1.get("application_open_date") or p2.get("application_open_date")
        else:
            deadline  = max(
                [p1.get("application_deadline"), p2.get("application_deadline")],
                key=lambda d: date_relevance_score(d or "")
            ) or None
            open_date = max(
                [p1.get("application_open_date"), p2.get("application_open_date")],
                key=lambda d: date_relevance_score(d or "")
            ) or None
        intake     = p1.get("intake_semester")        or p2.get("intake_semester")
        app_status = p1.get("application_status")     or p2.get("application_status") or "not_mentioned"
        apply_url  = p1.get("apply_button_url")       or p2.get("apply_url") or ""
        notes      = p1.get("notes") or p2.get("notes") or ""

        # Sometimes the model puts dates in notes instead of the date fields
        if not deadline and not open_date and notes:
            extracted = extract_date_from_notes(notes)
            if extracted:
                deadline = extracted

        found = "yes" if (deadline or open_date) else "no"
        found_in_announcement = p1.get("found_in_announcement") or ""

        rows.append({
            "id":                  pid,
            "name_en":             prog["name_en"],
            "name_gr":             prog["name_gr"],
            "university":          prog["university"],
            "university_gr":       prog["university_gr"],
            "department":          prog["department"],
            "city":                prog["city"],
            "topics":              prog["topics"],
            "languages":           prog["languages"],
            "ects":                prog["ects"],
            "semesters":           prog["semesters"],
            "tuition":             prog["tuition"],
            "study_modes":         prog["study_modes"],
            "email":               prog["email"],
            "scholarship":         "yes" if prog["scholarship_offered"] else "no",
            "university_image_url":prog["university_image_url"],
            "programme_url":       prog["url"],
            "open_date":           open_date or "",
            "deadline":            deadline or "",
            "intake":              intake or "",
            "apply_url":           apply_url,
            "application_status":  app_status,
            "notes":               notes,
            "portal":              p2.get("portal_name") or "",
            "requires_login":      str(p2.get("requires_login", "")).lower(),
            "found":               found,
            "found_in_announcement": found_in_announcement,
            "scrape_status":       p1.get("status", "not_scraped"),
            "pass2_status":        p2.get("status", "skipped"),
            "scraped_at":          p1.get("scraped_at", ""),
        })

    return rows


def export(rows):
    with open(RESULTS_JSON, "w", encoding="utf-8") as f:
        json.dump(rows, f, indent=2, ensure_ascii=False)
    log.info(f"JSON -> {RESULTS_JSON}")

    if rows:
        with open(RESULTS_CSV, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=rows[0].keys())
            writer.writeheader()
            writer.writerows(rows)
        log.info(f"CSV  -> {RESULTS_CSV}")

    total   = len(rows)
    found   = sum(1 for r in rows if r["found"] == "yes")
    errors  = sum(1 for r in rows if r["scrape_status"] == "error")
    missing = total - found - errors

    log.info(f"\n{'='*55}")
    log.info(f"Total   : {total}")
    log.info(f"Found   : {found}  ({round(found/total*100) if total else 0}%)")
    log.info(f"Missing : {missing}")
    log.info(f"Errors  : {errors}")
    log.info(f"{'='*55}\n")

# File helpers

def load_json_safe(path):
    if path.exists():
        try:
            with open(path, encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            log.warning(f"Could not load {path}: {e}")
    return []


def save_json(data, path):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    log.info(f"Saved {len(data)} records -> {path}")

def save_errors(pass1_results, prog_map):
    errors = [
        {
            "prog_id":   r.get("prog_id"),
            "url":       prog_map.get(r.get("prog_id"), {}).get("url", ""),
            "error":     r.get("error"),
            "scraped_at": r.get("scraped_at"),
        }
        for r in pass1_results
        if r.get("status") == "error"
    ]
    with open(ERRORS_JSON, "w", encoding="utf-8") as f:
        json.dump(errors, f, indent=2, ensure_ascii=False)
    log.info(f"Errors -> {ERRORS_JSON} ({len(errors)} records)")

# Command-line options
def parse_args():
    p = argparse.ArgumentParser(description="Study in Greece MSc dates scraper")
    p.add_argument("--csv",              default=DEFAULT_CSV)
    p.add_argument("--model",            default=DEFAULT_MODEL)
    p.add_argument("--resume",           action="store_true")
    p.add_argument("--workers",          type=int, default=DEFAULT_WORKERS)
    p.add_argument("--pass2-only",       action="store_true")
    p.add_argument("--include-archived", action="store_false", dest="active_only")
    p.add_argument("--limit", type=int, default=None)
    p.add_argument("--offset", type=int, default=0)
    p.add_argument("--missing-only", action="store_true")
    p.add_argument("--ids", type=str, default=None)
    p.set_defaults(active_only=True)
    return p.parse_args()

# Main run
def main():
    args = parse_args()

    all_programmes = load_programmes(args.csv, args.active_only)
    programmes = all_programmes[args.offset:]
    if args.limit:
        programmes = programmes[:args.limit]
    prog_map = {p["id"]: p for p in all_programmes}

    graph_config      = get_graph_config(args.model)
    SmartScraperGraph = import_scraper()

    log.info(f"Programmes : {len(programmes)} | Workers : {args.workers} | Model: {args.model}")

    existing_p1 = load_json_safe(PASS1_JSON)

    p1_by_id = {}
    for r in existing_p1:
        pid = r["prog_id"]
        if pid not in p1_by_id:
            p1_by_id[pid] = r
        else:
            p1_by_id[pid] = best_record([p1_by_id[pid], r])
    existing_p1 = list(p1_by_id.values())
    resume_ids  = {r["prog_id"] for r in existing_p1} if args.resume else set()
    
    if args.missing_only:
        existing_results = load_json_safe(RESULTS_JSON)
        found_ids = {r["id"] for r in existing_results if r.get("found") == "yes"}
        programmes = [p for p in programmes if p["id"] not in found_ids]
        log.info(f"--missing-only: {len(programmes)} programmes without dates")

    if args.ids:
        target_ids = set(args.ids.split(","))
        # Only wipe pass1 cache for target IDs when actually re-running pass1
        if not args.pass2_only:
            existing_p1 = [r for r in existing_p1 if r.get("prog_id") not in target_ids]
        programmes = [p for p in all_programmes if p["id"] in target_ids]
        log.info(f"--ids: targeting {len(programmes)} specific programmes")
    else:
        target_ids = set()

    if not args.pass2_only:
        new_p1 = run_pass1(programmes, graph_config, SmartScraperGraph, args.workers, resume_ids)
        all_p1 = existing_p1 + new_p1
        save_json(all_p1, PASS1_JSON)
        if _shutdown:
            log.info("Saved partial results. Run with --resume to continue.")
            rows = merge(all_programmes, all_p1, load_json_safe(PASS2_JSON))
            export(rows)
            save_errors(all_p1, prog_map)
            sys.exit(0)
    else:
        all_p1 = existing_p1
        log.info(f"--pass2-only: {len(all_p1)} existing pass1 results")

    if not all_p1:
        log.error("No pass1 results. Exiting.")
        sys.exit(1)

    existing_p2 = [r for r in load_json_safe(PASS2_JSON) if r.get("prog_id") not in target_ids]
    done_p2_ids = {r["prog_id"] for r in existing_p2 if r.get("prog_id")}
    new_p2      = run_pass2(all_p1, prog_map, graph_config, SmartScraperGraph, args.workers, done_p2_ids)
    all_p2      = existing_p2 + new_p2
    save_json(all_p2, PASS2_JSON)

    rows = merge(all_programmes, all_p1, all_p2)
    export(rows)
    save_errors(all_p1, prog_map)


if __name__ == "__main__":
    main()
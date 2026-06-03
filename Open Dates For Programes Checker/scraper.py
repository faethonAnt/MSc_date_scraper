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

PASS1_JSON      = OUTPUT_DIR / "pass1_results.json"
PASS2_JSON      = OUTPUT_DIR / "pass2_results.json"
RESULTS_JSON    = OUTPUT_DIR / "results.json"
RESULTS_CSV     = OUTPUT_DIR / "results.csv"
LOG_FILE        = OUTPUT_DIR / "scraper.log"

MAX_RETRIES     = 3
RETRY_DELAY     = 4
REQUEST_DELAY   = 2.0
DEFAULT_WORKERS = 3

# Common admissions/application paths
SUBPAGE_SUFFIXES = [
    "/admissions", "/apply", "/applications", "/how-to-apply",
    "/prospective-students", "/entry-requirements", "/registration",
    "/apply-now", "/application-form", "/application-process",
    "/how-to-apply", "/fees-and-funding", "/scholarships",
    "/open-days", "/intake", "/deadlines", "/dates",
    "/postgraduate", "/postgraduate-admissions",
    "/msc", "/masters", "/graduate",
    "/αιτησεις", "/εγγραφες", "/υποβολη-αιτησης",
    "/αιτηση", "/εγγραφη", "/προθεσμιες",
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
  "application_deadline": "Application deadline / closing date. If multiple rounds, list all as one string. Or null.",
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
"""

PASS2_PROMPT = """
You are extracting application timeline data from a university application portal page.
Return ONLY a valid JSON object. No markdown, no code fences, no explanation.

{
  "application_open_date": "Date when applications open or null",
  "application_deadline": "Application deadline or null",
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
    if isinstance(raw, dict) and set(raw.keys()) <= {"content"} and "content" in raw:
        content = raw["content"]
        if isinstance(content, str):
            try:
                return json.loads(strip_json_fences(content))
            except Exception:
                pass
    return raw


def clean_result(raw: dict) -> dict:
    """Convert placeholder values to None."""
    cleaned = {}
    for k, v in raw.items():
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
            log.warning(f"[attempt {attempt}/{MAX_RETRIES}] Error on {url}: {e}")

        if attempt < MAX_RETRIES:
            time.sleep(RETRY_DELAY * attempt)

    log.error(f"FAILED after {MAX_RETRIES} attempts: {url} — {last_error}")
    return {"status": "error", "error": last_error}


def try_subpages(base_url: str, prompt: str, graph_config: dict, SmartScraperGraph) -> dict:
    parsed = urlparse(base_url)
    base = f"{parsed.scheme}://{parsed.netloc}"

    sitemap_urls = fetch_sitemap_urls(base_url)
    suffix_urls = [base.rstrip("/") + s for s in SUBPAGE_SUFFIXES]
    
    all_urls = list(dict.fromkeys(sitemap_urls + suffix_urls))

    for url in all_urls:
        if _shutdown:
            break
        if url == base_url:
            continue
        log.info(f"  Trying sub-page: {url}")
        result = scrape_url(url, prompt, graph_config, SmartScraperGraph)
        if result.get("status") == "ok" and result.get("has_dates_on_page") is True:
            log.info(f"  Found dates on sub-page: {url}")
            result["found_on_subpage"] = url
            return result
        time.sleep(1.0)

    return {}

#raw_html
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
    
    html = fetch_html(sitemap_url)
    if not html:
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
        return filtered[:10]
    
    except Exception as e:
        log.warning(f"fetch_sitemap_urls failed for {base_url}: {e}")
        return []
    
def find_pdf_links(html: str, base_url: str) -> list:
    keywords = [
        # English
        "apply", "application", "admission", "deadline", "dates",
        "open", "brochure", "prospectus", "announcement", "call",
        # Greek
        "αιτηση", "αιτήση", "αιτησεις", "αιτήσεις", "εγγραφη",
        "εγγραφή", "εγγραφες", "εγγραφές", "προθεσμια", "προθεσμία",
        "ανακοινωση", "ανακοίνωση", "προκηρυξη", "προκήρυξη"
    ]
    try:
        soup = BeautifulSoup(html, "html.parser")
        links = []
        for a in soup.find_all("a", href=True):
            href = a["href"].strip()
            if href.lower().endswith(".pdf"):
                absolute = urljoin(base_url, href)
                links.append(absolute)
        return links[:5]
    except Exception as e:
        log.warning(f"find_pdf_links failed: {e}")
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

# Workers

def worker_pass1(args):
    idx, total, prog, prompt, graph_config, SmartScraperGraph = args
    if _shutdown:
        return {"prog_id": prog["id"], "status": "cancelled", "scraped_at": now_utc()}

    log.info(f"[Pass1 {idx}/{total}] {prog['university']} — {prog['name_en'][:55]}")
    
    # Fetch raw HTML first
    html = fetch_html(prog["url"])
    
    result = scrape_url(prog["url"], prompt, graph_config, SmartScraperGraph)

    # If no dates found, check PDF links on the page
    if html and result.get("status") == "ok" and result.get("has_dates_on_page") is not True:
        pdf_links = find_pdf_links(html, prog["url"])
        for pdf_url in pdf_links:
            log.info(f"  Trying PDF: {pdf_url}")
            pdf_text = extract_pdf_text(pdf_url)
            if pdf_text:
                pdf_result = scrape_url(
                    f"data:text/plain,{pdf_text}",
                    prompt, graph_config, SmartScraperGraph
                )
                if pdf_result.get("has_dates_on_page") is True:
                    log.info(f"  Found dates in PDF: {pdf_url}")
                    pdf_result["found_in_pdf"] = pdf_url
                    result = pdf_result
                    break

    # No useful date info found, check sitemap and subpages
    if (result.get("status") == "ok"
            and result.get("has_dates_on_page") is not True
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

def run_pass2(pass1_results, prog_map, graph_config, SmartScraperGraph, workers, done_ids):
    candidates = [
        r for r in pass1_results
        if r.get("has_dates_on_page") is not True
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
         r["apply_button_url"],
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

DATE_PATTERNS = [
    r'\b\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}\b',
    r'\b\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}\b',
    r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+\d{4}\b',
    r'\b\d{1,2}(?:st|nd|rd|th)\s+(?:of\s+)?(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}\b',
]
DATE_RE = re.compile("|".join(DATE_PATTERNS), re.IGNORECASE)

def extract_date_from_notes(notes: str) -> str | None:
    if not notes:
        return None
    m = DATE_RE.search(notes)
    return m.group().strip() if m else None

# Merge and export

def merge(programmes, pass1, pass2):
    p1_map = {r["prog_id"]: r for r in pass1}
    p2_map = {r["prog_id"]: r for r in pass2}

    rows = []
    for prog in programmes:
        pid = prog["id"]
        p1  = p1_map.get(pid, {})
        p2  = p2_map.get(pid, {})

        deadline   = p1.get("application_deadline")  or p2.get("application_deadline")
        open_date  = p1.get("application_open_date") or p2.get("application_open_date")
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

# Command-line options

def parse_args():
    p = argparse.ArgumentParser(description="Study in Greece MSc dates scraper")
    p.add_argument("--csv",              default=DEFAULT_CSV)
    p.add_argument("--model",            default=DEFAULT_MODEL)
    p.add_argument("--resume",           action="store_true")
    p.add_argument("--workers",          type=int, default=DEFAULT_WORKERS)
    p.add_argument("--pass2-only",       action="store_true")
    p.add_argument("--include-archived", action="store_false", dest="active_only")
    p.add_argument("--limit",type=int, default=None)
    p.set_defaults(active_only=True)
    return p.parse_args()

# Main run

def main():
    args = parse_args()

    programmes = load_programmes(args.csv, args.active_only)
    if args.limit:
        programmes = programmes[:args.limit]
    prog_map   = {p["id"]: p for p in programmes}

    graph_config      = get_graph_config(args.model)
    SmartScraperGraph = import_scraper()

    log.info(f"Programmes : {len(programmes)} | Workers : {args.workers} | Model: {args.model}")

    existing_p1 = load_json_safe(PASS1_JSON)
    resume_ids  = {r["prog_id"] for r in existing_p1} if args.resume else set()

    if not args.pass2_only:
        new_p1 = run_pass1(programmes, graph_config, SmartScraperGraph, args.workers, resume_ids)
        all_p1 = existing_p1 + new_p1
        save_json(all_p1, PASS1_JSON)
        if _shutdown:
            log.info("Saved partial results. Run with --resume to continue.")
            rows = merge(programmes, all_p1, load_json_safe(PASS2_JSON))
            export(rows)
            sys.exit(0)
    else:
        all_p1 = existing_p1
        log.info(f"--pass2-only: {len(all_p1)} existing pass1 results")

    if not all_p1:
        log.error("No pass1 results. Exiting.")
        sys.exit(1)

    existing_p2 = load_json_safe(PASS2_JSON)
    done_p2_ids = {r["prog_id"] for r in existing_p2}
    new_p2      = run_pass2(all_p1, prog_map, graph_config, SmartScraperGraph, args.workers, done_p2_ids)
    all_p2      = existing_p2 + new_p2
    save_json(all_p2, PASS2_JSON)

    rows = merge(programmes, all_p1, all_p2)
    export(rows)


if __name__ == "__main__":
    main()
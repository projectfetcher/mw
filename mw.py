import os
import re
import csv
import sys
import time
import json
import base64
import hashlib
import logging
from datetime import datetime, timedelta
from urllib.parse import urljoin, urlparse, urlsplit, urlunsplit, parse_qsl, urlencode

import requests
import urllib3
from bs4 import BeautifulSoup

# Suppress InsecureRequestWarning that fires on every verify=False call —
# we accept the risk for internal WP API calls on self-hosted servers.
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

try:
    import openpyxl
    _XLSX_AVAILABLE = True
except ImportError:
    _XLSX_AVAILABLE = False

try:
    import language_tool_python
    from sentence_transformers import SentenceTransformer, util as st_util
    _NLP_AVAILABLE = True
except ImportError:
    _NLP_AVAILABLE = False

# =============================================================================
#  CONFIG
# =============================================================================

BASE_URL        = "https://jobsearchmalawi.com"
SITEMAP_URL     = "https://jobsearchmalawi.com/job_listing-sitemap.xml"

MAX_JOBS        = int(os.environ.get("MAX_JOBS", "0"))
REQUEST_DELAY   = float(os.environ.get("REQUEST_DELAY", "1.5"))
REQUEST_TIMEOUT = int(os.environ.get("REQUEST_TIMEOUT", "25"))

RESOLVE_APPLY_URLS = os.environ.get("RESOLVE_APPLY_URLS", "1") != "0"
RESOLVE_DELAY      = float(os.environ.get("RESOLVE_DELAY", "0.5"))

OUTPUT_FILE          = "jobs_output.xlsx"
PROCESSED_IDS_FILE   = "processed_jobs.csv"
FLAGGED_FILE         = "flagged_no_apply.csv"

_TRACKER_FIELDS = ["Job ID", "Job URL", "Job Title", "Company Name",
                   "Status", "Timestamp", "WP ID"]

_FLAGGED_FIELDS = ["Job ID", "Job URL", "Job Title", "Company Name",
                   "Reason", "Timestamp"]

# ── WordPress ─────────────────────────────────────────────────────────────────
# Set WP_BASE_URL to the full URL of your WordPress site, e.g.:
#   export WP_BASE_URL="https://malawi.mimusjobs.com"
# The REST API base and endpoint paths are derived from it automatically.
# probe_wp_api() in main() will test the endpoint and correct it if needed
# (e.g. if the REST API lives on the root domain instead of the subdomain).
WP_URL      = os.environ.get("WP_BASE_URL", "").rstrip("/")
WP_USER     = os.environ.get("WP_USERNAME", "")
WP_PASSWORD = os.environ.get("WP_APP_PASSWORD", "")
WP_BASE      = WP_URL                          # alias used throughout
WP_API_BASE  = f"{WP_BASE}/wp-json/wp/v2"     # reassigned by probe_wp_api() if needed
WP_JOBS_URL    = f"{WP_BASE}/job-listings"
WP_MEDIA_URL = f"{WP_API_BASE}/media"

# ── Mistral ───────────────────────────────────────────────────────────────────
MISTRAL_API_KEY = os.environ.get("MISTRAL_API_KEY", "")
MISTRAL_MODEL   = "mistral-small-latest"
MISTRAL_URL     = "https://api.mistral.ai/v1/chat/completions"

ENABLE_PARAPHRASE = True

JOB_TYPE_MAPPING = {
    "full-time": "full-time", "full time": "full-time",
    "part-time": "part-time", "part time": "part-time",
    "contract":  "contract",  "temporary": "temporary",
    "internship":"internship","freelance": "freelance",
    "volunteer": "volunteer",
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml",
    "Accept-Charset": "utf-8",
}

SESSION = requests.Session()
SESSION.headers.update(HEADERS)

_apply_url_cache = {}

# =============================================================================
#  LOGGING / COLOUR
# =============================================================================

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log_ = logging.getLogger(__name__)

_USE_COLOUR = sys.stdout.isatty()

def _c(code, text):
    return f"\033[{code}m{text}\033[0m" if _USE_COLOUR else text

C_HEADER  = lambda t: _c("1;36",  t)
C_LABEL   = lambda t: _c("1;33",  t)
C_VALUE   = lambda t: _c("97",    t)
C_DIM     = lambda t: _c("2",     t)
C_GREEN   = lambda t: _c("1;32",  t)
C_RED     = lambda t: _c("1;31",  t)
C_BLUE    = lambda t: _c("1;34",  t)
C_DIVIDER = lambda: _c("2", "─" * 80)

def log(msg):
    print(msg, flush=True)

TRACKING_PARAM_PREFIXES = ("utm_",)
TRACKING_PARAM_EXACT = {
    "fbclid", "gclid", "msclkid", "mc_cid", "mc_eid", "ref", "referrer",
}

EMAIL_PATTERN    = re.compile(r"[A-Za-z0-9.+_-]+@[A-Za-z0-9-]+\.[A-Za-z0-9.-]+")
META_REFRESH_PAT = re.compile(
    r'<meta[^>]+http-equiv=["\']refresh["\'][^>]*content=["\'][^"\'>]*url=([^"\'>]+)', re.I)
JS_REDIRECT_PAT  = re.compile(
    r'(?:window\.)?location(?:\.href)?\s*(?:=\s*|\.replace\(\s*)["\']([^"\']+)["\']', re.I)

BOILERPLATE_PATTERNS = [
    re.compile(r"go to method of application\s*[»>]*", re.I),
    re.compile(r"Read more about this company", re.I),
    re.compile(r"Subscribe to our newsletter.*", re.I),
]

# Deadline extraction from body text
# Matches all label variants seen on jobsearchmalawi.com in the wild:
#   "Closing Date:", "Closing date for receiving applications is", "Closure Date :",
#   "The closing date for receiving applications is", "Application Deadline:", etc.
_DEADLINE_LABEL_RE = re.compile(
    r'(?:the\s+)?'
    r'(?:closing\s+date(?:\s+for\s+(?:receipt|receiving)\s+(?:of\s+)?applications)?'
    r'|closure\s+date'
    r'|application\s+deadline'
    r'|apply\s+by'
    r'|deadline(?:\s+for\s+applications)?'
    r')'
    r'\s*(?:is\s+(?:(?:on\s+or\s+before|on)\s+)?|:\s*(?:(?:on\s+or\s+before|on)\s+)?|[:\-–]\s*)',
    re.I,
)
# Strip a repeated label from the remainder (e.g. "Closing Date: The closing date ... is")
_DEADLINE_NOISE_RE = re.compile(
    r'^(?:the\s+)?(?:closing\s+date|closure\s+date|application\s+deadline)'
    r'(?:\s+for\s+(?:receipt|receiving)\s+(?:of\s+)?applications)?'
    r'\s*(?:is\s+(?:(?:on\s+or\s+before|on)\s+)?|:\s*)',
    re.I,
)
# Extract a date value after the label has been removed
_DEADLINE_DATE_RE = re.compile(
    r'^((?:[A-Za-z]+,?\s+)?'                       # optional day name
    r'(?:close\s+of\s+business[, ]+)?'             # "Close of Business, "
    r'(?:applications?\s+close\s+\S+\s+\S+\s+\S+\s+on\s+)?'  # "Applications close ... on "
    r'(?:\d{1,2}(?:st|nd|rd|th)?\s+\w+\s+\d{4}'  # 23rd June 2026
    r'|\d{1,2}/\w+/\d{4}'                          # 24/Apr/2026
    r'|\d{1,2}/\d{1,2}/\d{4}'                      # 28/04/2026
    r'|\d{4}-\d{2}-\d{2}'                          # 2026-06-23
    r'|\w+\s+\d{1,2},?\s+\d{4}'                   # April 25, 2026
    r'|\d{1,2}\s+\w+\s+\d{4}'                     # 15 April 2026
    r')(?:[, ]+(?:at\s+)?[\d:]+\s*(?:am|pm|CAT|EAT|GMT|UTC)?)?)',
    re.I,
)

_MOJIBAKE = [
    ("Â", ""), ("â€™", "'"), ("â€œ", '"'), ("â€\x9d", '"'), ("â€", '"'),
    ("â€¢", "•"), ("â„¢", "™"), ("\u00a0", " "), ("\u200b", ""), ("\ufeff", ""),
]

# FIX: expanded to catch WP Image resize suffixes like -100x100 or ?resize=100,100
WP_RESIZE_PARAM_RE = re.compile(r"[?&](?:resize|w|h|fit|crop)=[^&]+")
WP_RESIZE_SUFFIX_RE = re.compile(r"-\d+x\d+(\.[a-zA-Z]{2,5})$")

# =============================================================================
#  TEXT HELPERS
# =============================================================================

def _fix_mojibake(text):
    for pattern, replacement in _MOJIBAKE:
        text = text.replace(pattern, replacement)
    text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "", text)
    return text

def sanitize_text(text, is_url=False):
    if not isinstance(text, str):
        text = str(text) if (text is not None and str(text) not in ("nan","None","NaN")) else ""
    text = text.strip()
    if text in ("nan","None","NaN","","N/A","n/a","NA","na"):
        return ""
    text = _fix_mojibake(text)
    if is_url:
        return re.sub(r"[ \t\r\n\f\v]+", " ", text).strip()
    text = re.sub(r"#+\s*", "", text)
    text = re.sub(r"\*\*", "", text)
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()

def clean_text(el):
    if el is None:
        return ""
    return re.sub(r"\s+", " ", el.get_text(" ", strip=True)).strip()

def clean_description(text):
    if not text:
        return text
    for pattern in BOILERPLATE_PATTERNS:
        text = pattern.sub("", text)
    return re.sub(r"[ \t]+", " ", text).strip()

def extract_deadline_from_text(text):
    """
    Scan body text for a deadline sentence and return the date value only.
    Handles all formats observed on jobsearchmalawi.com:
      "Closing Date: 28/04/2026"
      "Closing date for receiving applications is Friday, 28th April, 2026"
      "The closing date for receiving applications is Thursday 30th April 2026, at 5 pm"
      "Closure Date : 24/Apr/2026, 11:59:00 PM"
      "Application Deadline: 25th March 2026"
      "Closing Date: Close of Business, 29 May 2026"
    Returns '' if nothing found.
    """
    if not text:
        return ""
    m = _DEADLINE_LABEL_RE.search(text)
    if not m:
        return ""
    remainder = text[m.end():].strip()
    # Strip a repeated label clause (e.g. label was "Closing Date:" but value starts
    # with "The closing date for receiving applications is 20th June 2026")
    remainder = _DEADLINE_NOISE_RE.sub("", remainder).strip()
    # Try to extract a clean date token
    dm = _DEADLINE_DATE_RE.match(remainder)
    if dm:
        return dm.group(1).strip().rstrip(",")
    # Fallback: everything up to the first full stop or newline, capped at 80 chars
    fallback = re.split(r"[.\n]", remainder)[0].strip()
    return fallback[:80] if fallback else ""


    if not url:
        return url
    parts = urlsplit(url)
    if not parts.query:
        return url
    kept = [(k, v) for k, v in parse_qsl(parts.query, keep_blank_values=True)
            if not k.lower().startswith(TRACKING_PARAM_PREFIXES) and k.lower() not in TRACKING_PARAM_EXACT]
    return urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(kept), parts.fragment))

def _same_site(url):
    try:
        return urlparse(url).netloc == urlparse(BASE_URL).netloc
    except Exception:
        return False

def extract_email(text):
    if not text:
        return ""
    m = EMAIL_PATTERN.search(text)
    return m.group(0) if m else ""

def absolute_url(href):
    if not href:
        return ""
    return href if href.startswith("http") else urljoin(BASE_URL, href)

# FIX: strip WP image resize query params and -NNNxNNN suffixes so the full
#      original image URL is used, which loads correctly and uploads cleanly.
def normalize_logo_url(url):
    if not url:
        return url
    # Strip query-string resize params
    url = WP_RESIZE_PARAM_RE.sub("", url).rstrip("?&")
    # Strip -100x100 style suffixes before extension
    url = WP_RESIZE_SUFFIX_RE.sub(r"\1", url)
    return url

# FIX: parse relative dates like "3 days ago", "1 week ago", "2 months ago"
def resolve_relative_date(text):
    """Convert '3 days ago', '1 week ago', etc. to YYYY-MM-DD. Returns original text if not matched."""
    if not text:
        return text
    text = text.strip()
    m = re.match(r"(\d+)\s+(day|week|month)s?\s+ago", text, re.I)
    if not m:
        return text
    n, unit = int(m.group(1)), m.group(2).lower()
    delta = {"day": timedelta(days=n), "week": timedelta(weeks=n),
             "month": timedelta(days=n * 30)}[unit]
    return (datetime.now() - delta).strftime("%Y-%m-%d")

# =============================================================================
#  HTTP HELPERS
# =============================================================================

def get_soup(url):
    resp = SESSION.get(url, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()
    resp.encoding = resp.encoding or "utf-8"
    return BeautifulSoup(resp.text, "lxml")

def _find_html_redirect(html, current_url):
    m = META_REFRESH_PAT.search(html) or JS_REDIRECT_PAT.search(html)
    if not m:
        return ""
    target = m.group(1).strip().strip("'\"")
    return urljoin(current_url, target)

def resolve_apply_url(raw):
    if not raw:
        return ""
    if raw in _apply_url_cache:
        return _apply_url_cache[raw]
    resolved = ""
    try:
        resp = SESSION.get(raw, timeout=REQUEST_TIMEOUT, allow_redirects=True)
        final = resp.url
        if final and not _same_site(final):
            resolved = final
        else:
            html_target = _find_html_redirect(resp.text, final)
            resolved = html_target if html_target else (final if final != raw else "")
    except requests.RequestException as e:
        log(f"    WARNING: could not resolve apply URL {raw}: {e}")
    resolved = strip_tracking_params(resolved)
    _apply_url_cache[raw] = resolved
    if RESOLVE_DELAY:
        time.sleep(RESOLVE_DELAY)
    return resolved

def resolve_application_contact(raw_apply_url, description, raw_anchor_text=""):
    """
    FIX: also check raw_anchor_text for emails (e.g. mailto: links rendered as text).
    Returns dict with apply_url and/or apply_email populated.
    """
    result = {"apply_url": "", "apply_email": "", "apply_raw": raw_apply_url}

    # If raw_apply_url is already an email address
    if raw_apply_url and re.match(r"^[A-Za-z0-9.+_-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$", raw_apply_url):
        result["apply_email"] = raw_apply_url
        return result

    if raw_apply_url and RESOLVE_APPLY_URLS:
        result["apply_url"] = resolve_apply_url(raw_apply_url)

    if not result["apply_url"]:
        # FIX: try anchor text first (may contain visible email), then description
        email = extract_email(raw_anchor_text) or extract_email(description)
        result["apply_email"] = email

    return result

# =============================================================================
#  LOGO EXTRACTION
# =============================================================================

LOGO_KW_RE      = re.compile(r"logo", re.I)
# FIX: added "company-placeholder" which is WP Job Manager's default logo class
PLACEHOLDER_RE  = re.compile(r"default|placeholder|avatar|no-?image|blank|generic|company-placeholder", re.I)
SITE_BRAND_RE   = re.compile(r"jobsearchmalawi", re.I)

def clean_logo_url(raw):
    if not raw:
        return ""
    raw = raw.strip()
    if not raw.startswith("http"):
        raw = absolute_url(raw)
    raw = re.sub(r"[\"')\s]+$", "", raw)
    # FIX: strip WP resize params so we get the full-size original
    return normalize_logo_url(raw)

def is_placeholder_logo(url):
    return not url or bool(PLACEHOLDER_RE.search(url))

def extract_company_logo(soup):
    for img in soup.find_all("img"):
        if img.find_parent(["header", "nav", "footer"]):
            continue
        blob = " ".join(filter(None, [
            " ".join(img.get("class", []) or []),
            img.get("id",""), img.get("alt",""), img.get("src",""),
        ]))
        if LOGO_KW_RE.search(blob) or re.search(r"compan", blob, re.I):
            src = img.get("src") or img.get("data-src") or ""
            cand = clean_logo_url(src)
            if cand and not is_placeholder_logo(cand) and not SITE_BRAND_RE.search(cand):
                return cand
    og = soup.find("meta", property="og:image") or soup.find("meta", attrs={"name":"og:image"})
    if og:
        cand = clean_logo_url(og.get("content",""))
        if cand and not is_placeholder_logo(cand) and not SITE_BRAND_RE.search(cand):
            return cand
    return ""

# =============================================================================
#  NLP (optional)
# =============================================================================

_grammar_tool = None
_sim_model    = None

def _get_sim_model():
    global _sim_model
    if _sim_model is None and _NLP_AVAILABLE:
        try:
            _sim_model = SentenceTransformer("all-MiniLM-L6-v2", device="cpu")
        except Exception as e:
            log_.warning(f"SentenceTransformer init failed: {e}")
    return _sim_model

def similarity_score(a, b):
    model = _get_sim_model()
    if model:
        try:
            emb = model.encode([a, b], convert_to_tensor=True)
            return float(st_util.pytorch_cos_sim(emb[0], emb[1]))
        except Exception:
            pass
    def tokens(s):
        return set(re.sub(r"[^a-z0-9 ]", " ", s.lower()).split())
    ta, tb = tokens(a), tokens(b)
    if not ta or not tb:
        return 0.0
    return len(ta & tb) / max(len(ta), len(tb))

def clean_output(text):
    text = _fix_mojibake(text)
    for pat in [r"\[/?INST\]", r"</?s>",
                r"(?i)(rewritten?|rephrased?|output|paraphrase[d]?)[:\s]+",
                r"\*\*", r"###", r"---"]:
        text = re.sub(pat, "", text)
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()

# =============================================================================
#  MISTRAL
# =============================================================================

def mistral_generate(prompt, max_tokens=400, temperature=0.7):
    if not MISTRAL_API_KEY:
        return ""
    try:
        response = requests.post(
            MISTRAL_URL,
            headers={"Authorization": f"Bearer {MISTRAL_API_KEY}", "Content-Type": "application/json"},
            json={"model": MISTRAL_MODEL, "messages": [{"role":"user","content":prompt}],
                  "max_tokens": max_tokens, "temperature": temperature},
            timeout=30,
        )
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"].strip()
    except Exception as e:
        log_.error(f"Mistral API error: {e}")
        return ""

def paraphrase_title(title):
    if not ENABLE_PARAPHRASE or not MISTRAL_API_KEY:
        return title
    clean = sanitize_text(title)
    if not clean:
        return title
    best_result, best_sim = None, 0.0
    for attempt in range(4):
        temp = round(0.68 + attempt * 0.06, 2)
        prompt = (f"Rewrite this job title professionally using different words. "
                  f"Output ONLY the rewritten title, nothing else. "
                  f"Keep it between 4 and 12 words.\n\nJob title: {clean}")
        raw    = mistral_generate(prompt, max_tokens=50, temperature=temp)
        result = clean_output(raw).split("\n")[0].strip().strip('"').strip("'")
        wc     = len(result.split()) if result else 0
        sim    = similarity_score(clean, result) if result else 0.0
        is_dup = result.lower().strip() == clean.lower().strip()
        valid  = bool(result) and 4 <= wc <= 14 and sim >= 0.55 and not is_dup
        if valid and sim > best_sim:
            best_sim, best_result = sim, result
        time.sleep(1)
    return best_result if best_result else clean

def paraphrase_description(text):
    if not ENABLE_PARAPHRASE or not MISTRAL_API_KEY:
        return text
    clean = sanitize_text(text)
    if not clean:
        return text
    paragraphs = [p.strip() for p in re.split(r"\n+", clean) if p.strip()] or [clean]
    rewritten  = []
    for para in paragraphs:
        prompt = (f"Rewrite this job description paragraph professionally. "
                  f"Keep ALL facts, requirements, and responsibilities. "
                  f"Use different sentence structure and vocabulary. "
                  f"Output ONLY the rewritten paragraph — no labels, no explanation.\n\n"
                  f"Original:\n{para}")
        best_result, best_sim, accepted = None, 0.0, None
        for attempt in range(3):
            raw    = mistral_generate(prompt, max_tokens=500, temperature=round(0.65 + attempt*0.08, 2))
            result = clean_output(raw).strip()
            rw     = len(result.split()) if result else 0
            sim    = similarity_score(para, result) if result and rw >= 5 else 0.0
            valid  = bool(result) and rw >= 8 and sim >= 0.48
            if valid:
                accepted = result
                break
            if result and sim > best_sim:
                best_sim, best_result = sim, result
            time.sleep(1)
        rewritten.append(accepted or (best_result if best_result and best_sim >= 0.40 else para))
    return "\n\n".join(rewritten)

def paraphrase_company(text):
    if not ENABLE_PARAPHRASE or not MISTRAL_API_KEY:
        return text
    clean = sanitize_text(text)
    if not clean:
        return text
    prompt = (f"Rewrite this company description professionally. "
              f"Preserve all facts. Use different wording. "
              f"Output ONLY the rewritten description.\n\nOriginal:\n{clean}")
    raw    = mistral_generate(prompt, max_tokens=600, temperature=0.68)
    result = clean_output(raw)
    rw     = len(result.split()) if result else 0
    time.sleep(1)
    return result if result and rw >= 10 else clean

# =============================================================================
#  TRACKER (stdlib csv only)
# =============================================================================

def _init_tracker():
    if not os.path.exists(PROCESSED_IDS_FILE):
        try:
            with open(PROCESSED_IDS_FILE, "w", newline="", encoding="utf-8") as f:
                csv.writer(f).writerow(_TRACKER_FIELDS)
        except Exception as e:
            log_.error(f"Could not create tracker: {e}")
    # FIX: also init the flagged CSV
    if not os.path.exists(FLAGGED_FILE):
        try:
            with open(FLAGGED_FILE, "w", newline="", encoding="utf-8") as f:
                csv.writer(f).writerow(_FLAGGED_FIELDS)
        except Exception as e:
            log_.error(f"Could not create flagged file: {e}")

def load_processed_ids():
    _init_tracker()
    ids, urls = set(), set()
    try:
        with open(PROCESSED_IDS_FILE, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                if row.get("Job ID"):   ids.add(row["Job ID"].strip())
                if row.get("Job URL"):  urls.add(row["Job URL"].strip())
    except Exception as e:
        log_.error(f"Could not read tracker: {e}")
    return ids, urls

def _upsert_row(job_id, updates):
    _init_tracker()
    rows = []
    try:
        with open(PROCESSED_IDS_FILE, newline="", encoding="utf-8") as f:
            rows = list(csv.DictReader(f))
    except Exception:
        pass
    found = False
    for row in rows:
        if row.get("Job ID","").strip() == str(job_id):
            row.update(updates)
            row["Timestamp"] = datetime.now().isoformat()
            found = True
            break
    if not found:
        new_row = {k:"" for k in _TRACKER_FIELDS}
        new_row["Job ID"]    = str(job_id)
        new_row["Timestamp"] = datetime.now().isoformat()
        new_row.update(updates)
        rows.append(new_row)
    try:
        with open(PROCESSED_IDS_FILE, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=_TRACKER_FIELDS, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(rows)
    except Exception as e:
        log_.error(f"Tracker write error: {e}")

# FIX: write flagged jobs (no apply contact) to dedicated CSV, never post them
def flag_job(job_id, job_url, title, company, reason="no_apply_contact"):
    log_.warning(f"FLAGGED [{reason}]: {title} | {job_url}")
    rows = []
    try:
        with open(FLAGGED_FILE, newline="", encoding="utf-8") as f:
            rows = list(csv.DictReader(f))
    except Exception:
        pass
    # Avoid duplicate flagged entries
    if any(r.get("Job ID","").strip() == str(job_id) for r in rows):
        return
    rows.append({
        "Job ID": job_id, "Job URL": job_url, "Job Title": title,
        "Company Name": company, "Reason": reason,
        "Timestamp": datetime.now().isoformat(),
    })
    try:
        with open(FLAGGED_FILE, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=_FLAGGED_FIELDS, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(rows)
    except Exception as e:
        log_.error(f"Flagged file write error: {e}")

def make_job_id(job_url, title="", company=""):
    seed = job_url or f"{title}{company}"
    return hashlib.md5(seed.encode()).hexdigest()[:16]

def mark_scraped(job_id, job_url, title, company):
    _upsert_row(job_id, {"Job URL": job_url, "Job Title": title,
                         "Company Name": company, "Status": "scraped", "WP ID": ""})

def mark_paraphrased(job_id):
    _upsert_row(job_id, {"Status": "paraphrased"})

def mark_posted(job_id, wp_id):
    _upsert_row(job_id, {"Status": "posted", "WP ID": str(wp_id)})

def mark_failed(job_id, reason):
    _upsert_row(job_id, {"Status": f"failed|{reason}"})

# =============================================================================
#  STEP 1 — COLLECT URLS FROM SITEMAP
# =============================================================================

def collect_job_urls():
    """Fetch the XML sitemap and extract all job listing URLs."""
    log(f"\n{'='*80}\nFETCHING SITEMAP: {SITEMAP_URL}\n{'='*80}")
    try:
        resp = SESSION.get(SITEMAP_URL, timeout=REQUEST_TIMEOUT)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "xml")
        urls = [loc.get_text(strip=True) for loc in soup.find_all("loc")]
        log(f"  Found {len(urls)} job URLs in sitemap")
        return urls
    except Exception as e:
        log(f"  ERROR fetching sitemap: {e}")
        return []

# =============================================================================
#  STEP 2 — PARSE INDIVIDUAL JOB PAGE
# =============================================================================

def parse_job_page(url):
    """
    Parse a single job listing page from jobsearchmalawi.com.

    FIX — title: WP Job Manager renders the title as
        <h1 class="entry-title job_title"> — the old selector '.entry-title'
        works only if the theme keeps that class alone.  We now try multiple
        selectors in priority order and fall back to <meta og:title>.

    FIX — company: 'p.name' doesn't exist on most themes; WP Job Manager
        puts the company in <div class="company"><p>…</p></div> or
        <strong class="name"> or an <li> with class "company".  Added a
        cascade of selectors plus an og:site_name fallback.

    FIX — date posted: may be "12 days ago" or "1 week ago" — resolved to
        an actual date via resolve_relative_date().

    FIX — deadline: stripped with a broader regex covering all common label
        variants ("Closing Date:", "Deadline:", "Application Deadline:", etc.)

    FIX — apply email: raw anchor text (visible email text inside <a>) is
        now passed to resolve_application_contact() so emails rendered as
        link text are captured.
    """
    soup = get_soup(url)

    # ── Title ─────────────────────────────────────────────────────────────────
    # FIX: try multiple selectors; WP Job Manager uses h1 with multiple classes
    title = ""
    for sel in [
        "h1.entry-title.job_title",
        "h1.job_title",
        "h1.entry-title",
        ".job_title",
        "h1",
    ]:
        el = soup.select_one(sel)
        if el:
            title = clean_text(el)
            if title:
                break
    # Final fallback: og:title meta tag
    if not title:
        og_title = soup.find("meta", property="og:title")
        if og_title:
            title = og_title.get("content", "").strip()
            # Strip trailing " - JobSearchMalawi" style suffixes
            title = re.sub(r"\s*[-|–]\s*[A-Za-z ]+$", "", title).strip()

    # ── Job type ──────────────────────────────────────────────────────────────
    job_type = ""
    for sel in ["li.job-type", "li.type", ".job-type"]:
        el = soup.select_one(sel)
        if el:
            job_type = clean_text(el)
            # Strip label prefix: "Job Type: Full Time" → "Full Time"
            job_type = re.sub(r"^(job\s*type|type)[:\s]+", "", job_type, flags=re.I).strip()
            if job_type:
                break

    # ── Location ──────────────────────────────────────────────────────────────
    location = ""
    for sel in ["li.location", ".job-location", "li.job-location"]:
        el = soup.select_one(sel)
        if el:
            location = clean_text(el)
            location = re.sub(r"^(location)[:\s]+", "", location, flags=re.I).strip()
            if location:
                break

    # ── Date posted ───────────────────────────────────────────────────────────
    date_posted = ""
    for sel in ["li.date-posted", ".date-posted", "li.posted-date", "time[datetime]"]:
        el = soup.select_one(sel)
        if el:
            # Prefer machine-readable datetime attribute if present
            if el.name == "time" and el.get("datetime"):
                date_posted = el["datetime"][:10]
            else:
                date_posted = clean_text(el)
                date_posted = re.sub(
                    r"^(date\s*posted|posted|date)[:\s]+", "", date_posted, flags=re.I
                ).strip()
            if date_posted:
                break
    # FIX: resolve relative dates like "3 days ago"
    date_posted = resolve_relative_date(date_posted)

    # ── Deadline ──────────────────────────────────────────────────────────────
    # TWO-STAGE strategy:
    #   Stage 1 — WP meta sidebar field (only set when employer fills it in).
    #   Stage 2 — Regex scan of the full description body text, which is where
    #             the MAJORITY of deadlines appear on jobsearchmalawi.com
    #             (employers paste the closing date inside the vacancy text).
    deadline = ""
    # Stage 1: dedicated meta field / sidebar elements
    for sel in ["div.cfwjm_output", ".job-deadline", ".deadline", "li.closing-date",
                "li.job-expiry", ".job-expiry"]:
        el = soup.select_one(sel)
        if el:
            val = clean_text(el)
            val = re.sub(
                r"^(closes?\s*:?|closing\s*date\s*:?|deadline\s*:?|"
                r"application\s*deadline\s*:?|apply\s*by\s*:?|expir(?:y|es)\s*:?)\s*",
                "", val, flags=re.I,
            ).strip()
            if val:
                deadline = val
                break

    # Stage 2 deadline scan happens after description is extracted — see below.

    # ── Company ───────────────────────────────────────────────────────────────
    # FIX: WP Job Manager renders company in several ways depending on theme.
    #      Added a cascade of selectors and an og:site_name fallback.
    company = ""
    for sel in [
        ".company .company_name",   # some themes wrap it
        ".company strong.name",
        ".company p.name",
        ".company_name",
        "strong.name",
        "p.name",                   # original (often fails)
        ".company-name",
        "li.company",
        ".single-job-listing .company h3",
        ".single-job-listing .company p",
        ".job-company",
    ]:
        el = soup.select_one(sel)
        if el:
            candidate = clean_text(el)
            # Sanity check: skip if it looks like a job title or location
            if candidate and len(candidate) < 120 and not re.search(
                r"\b(required|responsibilities|qualifications|apply|salary)\b",
                candidate, re.I
            ):
                company = candidate
                break
    # Fallback: check og:site_name (often set to company name on WP Job Manager)
    if not company:
        og_site = soup.find("meta", property="og:site_name")
        if og_site:
            val = og_site.get("content","").strip()
            if val and val.lower() not in ("jobsearchmalawi", "jobsearch malawi",
                                           "jobs in malawi", "malawi jobs"):
                company = val

    # ── Company logo ──────────────────────────────────────────────────────────
    logo = ""
    company_section = soup.select_one(".company")
    if company_section:
        img_el = company_section.select_one("img")
        if img_el:
            src = img_el.get("src") or img_el.get("data-src") or ""
            logo = clean_logo_url(src)   # FIX: clean_logo_url now strips resize params
    if not logo or is_placeholder_logo(logo):
        logo = extract_company_logo(soup)

    # ── Company details / blurb ───────────────────────────────────────────────
    company_details = ""
    for sel in [".company-description", ".about-company", "#company-description",
                ".company-profile", "div.company"]:
        el = soup.select_one(sel)
        if el:
            txt = clean_text(el)
            # Don't accidentally grab the company name-only div
            if txt and len(txt) > 40:
                company_details = txt
                break
    if not company_details:
        for h in soup.find_all(re.compile(r"^h[2-4]$")):
            if re.search(r"about", h.get_text(), re.I):
                sibling = h.find_next_sibling()
                if sibling:
                    company_details = clean_text(sibling)
                break

    # ── Job description ───────────────────────────────────────────────────────
    desc_el = soup.select_one(".job_description") or soup.select_one("div.job_description")
    if desc_el:
        description = clean_description(clean_text(desc_el))
    else:
        description = ""
        for div in soup.find_all("div"):
            t = clean_text(div)
            if len(t) > len(description) and len(t) > 200:
                description = t
        description = clean_description(description)

    # Stage 2 deadline scan: most employers on jobsearchmalawi.com embed the
    # closing date inside the vacancy body text rather than a dedicated field.
    # Scan description only (not full page) to avoid nav/footer false matches.
    if not deadline:
        deadline = extract_deadline_from_text(description)

    # ── Application link ──────────────────────────────────────────────────────
    raw_apply      = ""
    raw_anchor_txt = ""  # FIX: capture visible anchor text for email extraction

    raw_html = str(soup)

    # Pattern 1: external URL "To apply for this job please visit <a href="...">..."
    url_pattern = re.compile(
        r'To apply for this job please visit\s*<a href="([^"]+)"',
        re.I,
    )
    # Pattern 2: email link
    email_pattern = re.compile(
        r'To apply for this job.*?email.*?to\s*<a[^>]*class=["\']job_application_email["\'][^>]*>(.+?)</a>',
        re.I | re.S,
    )

    m_url   = url_pattern.search(raw_html)
    m_email = email_pattern.search(raw_html)

    if m_url:
        raw_apply = m_url.group(1).strip()
    elif m_email:
        raw_anchor_txt = re.sub(r"<[^>]+>", "", m_email.group(1)).strip()
        raw_apply      = raw_anchor_txt

    # Direct element fallback
    if not raw_apply:
        a_el = (
            soup.select_one("a.job_application_email")
            or soup.select_one(".application-link a")
            or soup.select_one(".application a")
        )
        if a_el:
            href = a_el.get("href", "")
            if href.startswith("mailto:"):
                raw_apply      = re.sub(r"^mailto:", "", href).strip()
                raw_anchor_txt = clean_text(a_el)
            else:
                raw_apply      = href
                raw_anchor_txt = clean_text(a_el)

    log(f"    Resolving apply link for '{title}' …")
    # FIX: pass raw_anchor_txt so visible email text is also checked
    application = resolve_application_contact(raw_apply, description, raw_anchor_txt)

    return {
        "title":           title,
        "job_url":         url,
        "job_type":        job_type,
        "location":        location,
        "posted_date":     date_posted,
        "deadline":        deadline,
        "description":     description,
        "apply_url":       application["apply_url"],
        "apply_email":     application["apply_email"],
        "apply_raw":       application["apply_raw"],
        "company_name":    company,
        "company_details": company_details,
        "company_logo":    logo,
    }

# =============================================================================
#  STEP 3 — DEDUPLICATE + PARAPHRASE
# =============================================================================

def process_job(raw_job, processed_ids, processed_urls, seen_content):
    job_url = raw_job.get("job_url", "")
    title   = raw_job.get("title", "")
    company = raw_job.get("company_name", "")

    job_id = make_job_id(job_url, title, company)

    if job_id in processed_ids or job_url in processed_urls:
        log(C_DIM(f"  ⧳ Already processed — skipped: {job_url}"))
        return None

    fp = (title.lower().strip(), company.lower().strip(), raw_job.get("location","").lower().strip())
    if fp in seen_content:
        log(C_DIM(f"  ⧳ Duplicate this run — skipped: {title}"))
        return None
    seen_content.add(fp)

    # FIX: enforce hard rule — no apply contact = flag to CSV, never post
    apply_url   = raw_job.get("apply_url", "")
    apply_email = raw_job.get("apply_email", "")
    if not apply_url and not apply_email:
        log(C_RED(f"  ⚑ FLAGGED (no apply contact): {title}"))
        flag_job(job_id, job_url, title, company, reason="no_apply_contact")
        # Still mark as seen so we don't reprocess it on the next run
        mark_scraped(job_id, job_url, title, company)
        _upsert_row(job_id, {"Status": "flagged|no_apply_contact"})
        processed_ids.add(job_id)
        processed_urls.add(job_url)
        return None

    mark_scraped(job_id, job_url, title, company)
    processed_ids.add(job_id)
    processed_urls.add(job_url)

    description = raw_job.get("description", "")
    blurb       = raw_job.get("company_details", "")

    para_title = title
    para_desc  = description
    para_blurb = blurb

    if ENABLE_PARAPHRASE and MISTRAL_API_KEY:
        print(C_BLUE(f"\n  ✍️  Paraphrasing '{title}' …"))
        para_title = paraphrase_title(title)
        para_desc  = paraphrase_description(description)
        if blurb:
            para_blurb = paraphrase_company(blurb)
        mark_paraphrased(job_id)
    else:
        print(C_DIM("  ⚠️  Paraphrasing skipped"))

    application = apply_url or apply_email

    company_website = ""
    if apply_url:
        try:
            parts = urlsplit(apply_url)
            if parts.scheme and parts.netloc and "jobsearchmalawi" not in parts.netloc.lower():
                company_website = f"{parts.scheme}://{parts.netloc}"
        except Exception:
            pass

    apply_method = "resolved_redirect" if apply_url else ("description_email" if apply_email else "not_found")

    return {
        "jobTitle":          para_title,
        "jobDescription":    para_desc,
        "companyDetails":    para_blurb,
        "originalTitle":     title,
        "originalDesc":      description,
        "jobType":           raw_job.get("job_type", ""),
        "jobLocation":       raw_job.get("location", ""),
        "datePosted":        raw_job.get("posted_date", ""),
        "deadline":          raw_job.get("deadline", ""),
        "application":       application,
        "companyName":       company,
        "companyLogo":       raw_job.get("company_logo", ""),
        "companyWebsite":    company_website,
        "jobUrl":            job_url,
        "_jobId":            job_id,
        "_apply_method":     apply_method,
        "_apply_raw":        raw_job.get("apply_raw", ""),
    }

# =============================================================================
#  VERBOSE PRINTER
# =============================================================================

def print_job_verbose(index, job):
    desc = job.get("jobDescription", "")
    desc_preview = (desc[:400] + " [...]") if len(desc) > 400 else desc
    print()
    print(C_DIVIDER())
    print(C_HEADER(f"  JOB #{index}"))
    print(C_DIVIDER())
    print(f"  {C_LABEL('Title (original)')}    : {C_VALUE(job.get('originalTitle',''))}")
    print(f"  {C_LABEL('Title (paraphrased)')} : {C_GREEN(job.get('jobTitle',''))}")
    print(f"  {C_LABEL('Job Type')}             : {job.get('jobType','') or C_DIM('—')}")
    print(f"  {C_LABEL('Location')}             : {job.get('jobLocation','') or C_DIM('—')}")
    print(f"  {C_LABEL('Posted')}               : {job.get('datePosted','') or C_DIM('—')}")
    print(f"  {C_LABEL('Deadline')}             : {job.get('deadline','') or C_DIM('—')}")
    application = job.get("application","")
    print(f"  {C_LABEL('Apply')}                : {C_GREEN(application) if application else C_DIM('— not found —')}")
    print(f"  {C_LABEL('Apply Method')}         : {C_DIM(job.get('_apply_method',''))}")
    print()
    print(f"  {C_BLUE('── COMPANY ──────────────────────────────────────────')}")
    print(f"  {C_LABEL('Name')}      : {C_VALUE(job.get('companyName','') or C_DIM('—'))}")
    print(f"  {C_LABEL('Website')}   : {job.get('companyWebsite','') or C_DIM('—')}")
    print(f"  {C_LABEL('Logo')}      : {job.get('companyLogo','') or C_DIM('— none —')}")
    about = job.get("companyDetails","")
    if about:
        preview = (about[:200] + " [...]") if len(about) > 200 else about
        print(f"  {C_LABEL('About')}     : {preview}")
    print()
    print(f"  {C_BLUE('── DESCRIPTION PREVIEW ─────────────────────────────')}")
    print(desc_preview if desc_preview else C_DIM("   — no description —"))
    print(f"  {C_LABEL('Job URL')}   : {job.get('jobUrl','')}")
    print(C_DIVIDER())

# =============================================================================
#  WORDPRESS POSTING
# =============================================================================

def _wp_auth_headers():
    token = base64.b64encode(f"{WP_USER}:{WP_PASSWORD}".encode()).decode()
    return {"Authorization": f"Basic {token}", "Content-Type": "application/json"}

def _wp_get(path, params=None):
    """Authenticated GET to the WP REST API."""
    return requests.get(
        path, params=params,
        headers=_wp_auth_headers(),
        auth=(WP_USER, WP_PASSWORD),
        timeout=15, verify=False,
    )

def _wp_post(path, json_data):
    """Authenticated POST to the WP REST API."""
    return requests.post(
        path, json=json_data,
        headers=_wp_auth_headers(),
        auth=(WP_USER, WP_PASSWORD),
        timeout=20, verify=False,
    )

def _wp_json(r, context=""):
    """
    Safely parse a WP REST API response.
    Raises a clear error (logged) if the body is empty or not JSON —
    which happens on 401 Unauthorized, 403 Forbidden, or OLS/nginx redirects
    that return an HTML page instead of JSON.
    Returns None on failure so callers can skip gracefully.
    """
    if not r.content:
        log_.error(f"WP API empty response [{r.status_code}] {context} — check credentials/endpoint")
        return None
    try:
        return r.json()
    except Exception:
        snippet = r.text[:200].replace("\n", " ")
        log_.error(f"WP API non-JSON response [{r.status_code}] {context}: {snippet}")
        return None

def probe_wp_api():
    """
    Discover the correct WP REST API base URL for this site.
    Tries the configured WP_API_BASE first, then the root domain (in case
    WP_BASE_URL is a multisite subdomain but REST API is on the network root).

    Common reasons for 404 on /wp-json/wp/v2/posts:
      - WP_BASE_URL is a multisite subdomain; REST API served from root domain
      - Pretty permalinks are disabled (REST API requires them)
      - A security plugin or .htaccess rule is blocking /wp-json/
    Returns the working api base URL, or empty string.
    """
    root_domain = re.sub(r"^https?://[^.]+\.", lambda m: m.group(0).split("//")[0] + "//", WP_BASE)
    candidates = list(dict.fromkeys(filter(None, [
        WP_API_BASE,                              # as configured
        f"{root_domain.rstrip('/')}/wp-json/wp/v2",  # root domain fallback
        f"{WP_BASE.rstrip('/')}/?rest_route=/wp/v2",  # no-permalink fallback
    ])))
    for base in candidates:
        test_url = f"{base}/posts"
        try:
            r = requests.get(test_url, params={"per_page": 1},
                             headers=_wp_auth_headers(),
                             auth=(WP_USER, WP_PASSWORD),
                             timeout=10, verify=False)
            if r.status_code in (200, 401):
                log_.info(f"✅ WP REST API reachable: {base} [{r.status_code}]")
                return base
            log_.debug(f"WP probe {r.status_code}: {test_url}")
        except Exception as e:
            log_.debug(f"WP probe error {test_url}: {e}")
    log_.error(
        "❌ WP REST API unreachable. Check:\n"
        "   1. WP_BASE_URL is correct (currently: %s)\n"
        "   2. Pretty permalinks are enabled in WP Settings → Permalinks\n"
        "   3. /wp-json/ is not blocked by a security plugin or .htaccess",
        WP_BASE,
    )
    return ""


def get_or_create_term(taxonomy_url, name):
    if not name or not name.strip():
        return None
    slug = re.sub(r"[^a-z0-9-]", "-", name.lower().strip())
    slug = re.sub(r"-{2,}", "-", slug).strip("-")
    try:
        r = _wp_get(taxonomy_url, params={"slug": slug})
        data = _wp_json(r, f"GET {taxonomy_url}?slug={slug}")
        if isinstance(data, list) and data:
            return data[0]["id"]
    except Exception as e:
        log_.debug(f"Term lookup failed '{name}': {e}")
    try:
        r = _wp_post(taxonomy_url, {"name": name, "slug": slug})
        data = _wp_json(r, f"POST {taxonomy_url} name={name}")
        if data and isinstance(data, dict):
            return data.get("id")
    except Exception as e:
        log_.error(f"Term create error '{name}': {e}")
    return None

def build_jsonld(job):
    """Build a JSON-LD JobPosting schema block to embed in post content."""
    apply_url   = job.get("application","")
    apply_email = ""
    if re.match(r"^[A-Za-z0-9.+_-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$", apply_url):
        apply_email = apply_url
        apply_url   = ""

    schema = {
        "@context":    "https://schema.org",
        "@type":       "JobPosting",
        "title":       sanitize_text(job.get("jobTitle","")),
        "description": sanitize_text(job.get("jobDescription","")),
        "datePosted":  sanitize_text(job.get("datePosted","")),
        "hiringOrganization": {
            "@type": "Organization",
            "name":  sanitize_text(job.get("companyName","")),
        },
        "jobLocation": {
            "@type":   "Place",
            "address": {"@type": "PostalAddress", "addressLocality": sanitize_text(job.get("jobLocation",""))},
        },
    }
    if job.get("deadline"):
        schema["validThrough"] = sanitize_text(job["deadline"])
    if job.get("jobType"):
        schema["employmentType"] = sanitize_text(job["jobType"]).upper().replace(" ","-")
    if apply_url:
        schema["url"] = apply_url
    if apply_email:
        schema["applicationContact"] = {"@type": "ContactPoint", "email": apply_email}
    if job.get("companyWebsite"):
        schema["hiringOrganization"]["sameAs"] = sanitize_text(job["companyWebsite"], is_url=True)
    return f'\n\n<script type="application/ld+json">\n{json.dumps(schema, indent=2, ensure_ascii=False)}\n</script>\n'

def post_job_to_wordpress(job):
    if not WP_USER or not WP_PASSWORD:
        return None, None
    h = _wp_auth_headers()
    title       = sanitize_text(job.get("jobTitle",""))
    description = sanitize_text(job.get("jobDescription",""))
    if not title or not description:
        log_.warning("Skipping WP post — missing title or description")
        return None, None

    slug = re.sub(r"[^a-z0-9-]", "-", title.lower())
    slug = re.sub(r"-{2,}", "-", slug).strip("-")[:80]

    # Duplicate check by slug
    try:
        r = _wp_get(WP_JOBS_URL, params={"slug": slug, "status": "publish"})
        posts = _wp_json(r, f"duplicate check slug={slug}")
        if isinstance(posts, list) and posts:
            log_.info(f"⏭ Already on WP: {title}")
            return posts[0]["id"], posts[0].get("link")
    except Exception:
        pass

    logo_url    = sanitize_text(job.get("companyLogo",""), is_url=True)
    location    = sanitize_text(job.get("jobLocation",""))
    raw_type    = sanitize_text(job.get("jobType","")) or "Full-time"
    job_type_s  = JOB_TYPE_MAPPING.get(raw_type.lower().strip(), "full-time")
    company     = sanitize_text(job.get("companyName",""))
    application = sanitize_text(job.get("application",""), is_url=True)
    deadline    = sanitize_text(job.get("deadline",""))
    co_website  = sanitize_text(job.get("companyWebsite",""), is_url=True)
    about       = sanitize_text(job.get("companyDetails",""))
    date_posted = sanitize_text(job.get("datePosted",""))

    is_email_app = bool(re.match(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", application))
    is_url_app   = bool(re.match(r"^https?://[^\s]+$", application))
    if not (is_email_app or is_url_app):
        application = ""

    # ── Upload logo ───────────────────────────────────────────────────────────
    attachment_id = None
    if logo_url:
        try:
            img_r = SESSION.get(logo_url, timeout=15)
            if img_r.status_code == 200:
                ct  = img_r.headers.get("Content-Type","image/jpeg")
                # FIX: detect png/gif/webp properly
                if "png" in ct:
                    ext = "png"
                elif "gif" in ct:
                    ext = "gif"
                elif "webp" in ct:
                    ext = "webp"
                else:
                    ext = "jpg"
                fn  = re.sub(r"[^a-z0-9]", "-", company.lower()) + "-logo." + ext
                fn  = re.sub(r"-{2,}", "-", fn).strip("-")
                up_h = dict(_wp_auth_headers())
                up_h["Content-Disposition"] = f'attachment; filename="{fn}"'
                up_h["Content-Type"] = ct
                up_r = requests.post(
                    WP_MEDIA_URL, headers=up_h, data=img_r.content,
                    auth=(WP_USER, WP_PASSWORD), timeout=20, verify=False,
                )
                if up_r.status_code in (200, 201):
                    up_data = _wp_json(up_r, "media upload")
                    if up_data:
                        attachment_id = up_data.get("id")
        except Exception as e:
            log_.warning(f"Logo upload failed: {e}")

    # ── Taxonomy terms ────────────────────────────────────────────────────────
    cat_term_id = get_or_create_term(f"{WP_API_BASE}/categories", location) if location else None

    # Build tag names: job type, company name, source
    tag_names = list(filter(None, [
        job_type_s.replace("-"," ").title() if job_type_s else None,
        company if company else None,
        "JobSearchMalawi",
    ]))
    tag_ids = []
    for tag_name in tag_names:
        tid = get_or_create_term(f"{WP_API_BASE}/tags", tag_name)
        if tid:
            tag_ids.append(tid)

    # ── Build post content ────────────────────────────────────────────────────
    # Add hidden expiry comment + JSON-LD schema
    expiry_comment = f"<!-- job-expiry: {deadline} -->" if deadline else ""
    content = description + "\n\n" + expiry_comment + build_jsonld(job)

    payload = {
        "title":          title,
        "content":        content,
        "slug":           slug,
        "status":         "publish",
        "featured_media": attachment_id or 0,
        "meta": {
            "_job_location":    location,
            "_job_type":        job_type_s,
            "_application":     application,
            "_job_expires":     deadline,
            "_company_name":    company,
            "_company_website": co_website,
            "_company_logo":    str(attachment_id) if attachment_id else "",
            "_company_details": about,
            "_date_posted":     date_posted,
        },
    }
    if cat_term_id:
        payload["categories"] = [cat_term_id]
    if tag_ids:
        payload["tags"] = tag_ids

    for attempt in range(3):
        try:
            r = requests.post(
                WP_JOBS_URL, json=payload, headers=h,
                auth=(WP_USER, WP_PASSWORD), timeout=20, verify=False,
            )
            r.raise_for_status()
            post = _wp_json(r, f"create post '{title}'")
            if not post:
                raise ValueError("Empty/invalid JSON in post response")
            log_.info(f"✅ Posted: '{title}' → WP ID {post.get('id')}")
            return post.get("id"), post.get("link")
        except Exception as e:
            log_.error(f"WP post attempt {attempt+1} failed: {e}")
            if attempt < 2:
                time.sleep(2 ** attempt)
    return None, None

# =============================================================================
#  EXCEL EXPORT
# =============================================================================

EXCEL_HEADERS = [
    "Job Title", "Job Type", "Job Location", "Date Posted", "Deadline",
    "Job Description", "Application", "Apply Method", "Company Name",
    "Company Logo", "Company Website", "Company Details", "Job URL",
]

def _save_excel(jobs):
    if not _XLSX_AVAILABLE:
        log_.warning("openpyxl not installed — skipping Excel export")
        return
    if not jobs:
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(EXCEL_HEADERS)
    for job in jobs:
        ws.append([
            job.get("jobTitle",""),     job.get("jobType",""),
            job.get("jobLocation",""),  job.get("datePosted",""),
            job.get("deadline",""),     job.get("jobDescription",""),
            job.get("application",""),  job.get("_apply_method",""),
            job.get("companyName",""),  job.get("companyLogo",""),
            job.get("companyWebsite",""),job.get("companyDetails",""),
            job.get("jobUrl",""),
        ])
    wb.save(OUTPUT_FILE)
    log_.info(f"Saved {len(jobs)} rows → {OUTPUT_FILE}")

# =============================================================================
#  MAIN
# =============================================================================

def main():
    start_time = datetime.now()
    print()
    print(C_HEADER("=" * 80))
    print(C_HEADER("  JOBSEARCHMALAWI.COM SCRAPER + MISTRAL PARAPHRASE + WORDPRESS POSTING"))
    print(C_HEADER("=" * 80))
    print(f"  Source sitemap  : {SITEMAP_URL}")
    print(f"  Request delay   : {REQUEST_DELAY}s")
    print(f"  Max new jobs    : {'unlimited' if not MAX_JOBS else MAX_JOBS}")
    print(f"  Resolve apply   : {'✅ enabled' if RESOLVE_APPLY_URLS else '❌ disabled'}")
    print(f"  Paraphrase      : {'✅ enabled' if (ENABLE_PARAPHRASE and MISTRAL_API_KEY) else '❌ disabled'}")
    print(f"  WordPress post  : {'✅ enabled' if (WP_USER and WP_PASSWORD) else '❌ disabled'}")
    print(f"  Excel export    : {'✅ enabled' if _XLSX_AVAILABLE else '❌ disabled (pip install openpyxl)'}")
    print(f"  WP API base     : {WP_API_BASE or '(not configured)'}")
    print(f"  Started         : {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(C_HEADER("=" * 80))

    # Probe and auto-correct WP REST API endpoint before doing any work
    global WP_API_BASE, WP_JOBS_URL, WP_MEDIA_URL
    if WP_USER and WP_PASSWORD:
        discovered = probe_wp_api()
        if discovered and discovered != WP_API_BASE:
            log(C_BLUE(f"  ℹ️  WP API base corrected to: {discovered}"))
            WP_API_BASE  = discovered
            WP_JOBS_URL  = f"{WP_API_BASE}/posts"
            WP_MEDIA_URL = f"{WP_API_BASE}/media"
        elif not discovered:
            log(C_RED("  ⚠️  WP posting will be skipped — REST API unreachable"))

    _init_tracker()
    processed_ids, processed_urls = load_processed_ids()
    print(f"  Tracker loaded: {len(processed_ids)} previously processed job IDs\n")

    job_urls = collect_job_urls()
    if not job_urls:
        print(C_RED("  No URLs collected — exiting."))
        return

    jobs_out      = []
    seen_content  = set()
    posted_count  = 0
    flagged_count = 0
    errors        = 0

    for i, url in enumerate(job_urls, start=1):
        log(f"\nScraping job {i}/{len(job_urls)}: {url}")
        try:
            raw_job = parse_job_page(url)
        except Exception as e:
            errors += 1
            log(C_RED(f"  ERROR scraping {url}: {e}"))
            time.sleep(REQUEST_DELAY)
            continue

        try:
            job = process_job(raw_job, processed_ids, processed_urls, seen_content)
        except Exception as e:
            errors += 1
            log(C_RED(f"  ERROR processing job: {e}"))
            time.sleep(REQUEST_DELAY)
            continue

        if job is None:
            # Could be dedup skip OR flagged — check if it was flagged
            apply_url   = raw_job.get("apply_url","")
            apply_email = raw_job.get("apply_email","")
            if not apply_url and not apply_email:
                flagged_count += 1
            time.sleep(REQUEST_DELAY)
            continue

        jobs_out.append(job)
        print_job_verbose(len(jobs_out), job)

        print(C_BLUE("\n  📤 Posting to WordPress …"))
        wp_id, wp_url = post_job_to_wordpress(job)
        if wp_id:
            mark_posted(job["_jobId"], wp_id)
            posted_count += 1
            print(C_GREEN(f"  ✅ WP ID={wp_id}  🔗 {wp_url}"))
        else:
            mark_failed(job["_jobId"], "wp_post_failed_or_skipped")
            print(C_RED("  ❌ WordPress post failed / skipped"))

        if len(jobs_out) % 25 == 0:
            _save_excel(jobs_out)

        if MAX_JOBS and len(jobs_out) >= MAX_JOBS:
            log(f"\nMAX_JOBS limit ({MAX_JOBS}) reached, stopping.")
            break

        time.sleep(REQUEST_DELAY)

    _save_excel(jobs_out)

    end_time = datetime.now()
    duration = (end_time - start_time).total_seconds() / 60.0
    print()
    print(C_HEADER("=" * 80))
    print(C_HEADER("  SCRAPE COMPLETE"))
    print(C_HEADER("=" * 80))
    print(f"  {C_LABEL('URLs in sitemap')}       : {len(job_urls)}")
    print(f"  {C_LABEL('New jobs processed')}    : {C_GREEN(str(len(jobs_out)))}")
    print(f"  {C_LABEL('Posted to WordPress')}   : {C_GREEN(str(posted_count))}")
    print(f"  {C_LABEL('Flagged (no apply)')}    : {C_RED(str(flagged_count)) if flagged_count else '0'}")
    print(f"  {C_LABEL('Errors')}                : {C_RED(str(errors)) if errors else '0'}")
    print(f"  {C_LABEL('Duration')}              : ~{duration:.1f} min")
    print(f"  {C_LABEL('Output file')}           : {OUTPUT_FILE}")
    print(f"  {C_LABEL('Tracker file')}          : {PROCESSED_IDS_FILE}")
    print(f"  {C_LABEL('Flagged file')}          : {FLAGGED_FILE}")
    print(C_HEADER("=" * 80))


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(1)

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
mw.py — Multi-source Malawi job scraper + (optional) paraphrase + WordPress posting

SOURCES (each one is isolated: if a source is blocked / down / returns nothing
(403, 404, 5xx, timeouts, empty listings) it is logged and SKIPPED, and the run
continues with the next source):

    jobsearchmalawi  -> https://jobsearchmalawi.com   (XML sitemap)
    careersmw        -> https://careersmw.com         (WP REST API / sitemap / homepage)
    careeradmw       -> https://www.careeradmw.com    (paginated /jobs listing)
    linkedin         -> LinkedIn guest jobs API       (location = Malawi by default)

ENV VARS
    WP_BASE_URL, WP_USERNAME, WP_APP_PASSWORD   WordPress target (posting skipped if user/pass missing)
    INTERNAL_BOT_KEY                            Cloudflare allow-list header (X-Internal-Auth)
    MISTRAL_API_KEY                             optional; paraphrasing only runs when set
    ENABLE_PARAPHRASE=0                         force paraphrasing off even if key is set
    SOURCES=jobsearchmalawi,careersmw,careeradmw,linkedin
    MAX_JOBS=0                                  total new jobs cap for the run (0 = unlimited)
    MAX_JOBS_PER_SOURCE=0                       new jobs cap per source (0 = unlimited)
    REQUEST_DELAY=1.5   REQUEST_TIMEOUT=25
    RESOLVE_APPLY_URLS=1   RESOLVE_DELAY=0.5
    SOURCE_MAX_CONSECUTIVE_ERRORS=8             abort a source after N errors in a row
    CAREERADMW_MAX_PAGES=60   CAREERSMW_MAX_PAGES=40
    LINKEDIN_LOCATION=Malawi
    LINKEDIN_KEYWORDS=                          comma separated, empty = all jobs
    LINKEDIN_TIME_RANGE=r604800                 r86400 = 24h, r604800 = 7d, r2592000 = 30d
    LINKEDIN_MAX_PAGES=0                        0 = unlimited (LinkedIn caps at ~40 pages)
    LINKEDIN_DEFAULT_EXPIRY_DAYS=30             estimated deadline when none is found

CLI
    python mw.py --sources careersmw,careeradmw
"""

import os
import re
import csv
import sys
import time
import json
import base64
import hashlib
import logging
import argparse
from collections import Counter
from datetime import datetime, timedelta
from html import escape as html_escape
from urllib.parse import (urljoin, urlparse, urlsplit, urlunsplit, parse_qsl,
                          urlencode, unquote, quote_plus)

import requests
import urllib3
from bs4 import BeautifulSoup

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
    import lxml  # noqa: F401
    _PARSER = "lxml"
except ImportError:
    _PARSER = "html.parser"

try:
    from sentence_transformers import SentenceTransformer, util as st_util
    _NLP_AVAILABLE = True
except Exception:
    _NLP_AVAILABLE = False

# =============================================================================
#  CONFIG
# =============================================================================

MAX_JOBS            = int(os.environ.get("MAX_JOBS", "0"))
MAX_JOBS_PER_SOURCE = int(os.environ.get("MAX_JOBS_PER_SOURCE", "0"))
REQUEST_DELAY       = float(os.environ.get("REQUEST_DELAY", "1.5"))
REQUEST_TIMEOUT     = int(os.environ.get("REQUEST_TIMEOUT", "25"))
RESOLVE_APPLY_URLS  = os.environ.get("RESOLVE_APPLY_URLS", "1") != "0"
RESOLVE_DELAY       = float(os.environ.get("RESOLVE_DELAY", "0.5"))
MAX_CONSEC_ERRORS   = int(os.environ.get("SOURCE_MAX_CONSECUTIVE_ERRORS", "8"))
DEFAULT_SOURCES     = "jobsearchmalawi,careersmw,careeradmw,linkedin"

JSM_BASE     = "https://jobsearchmalawi.com"
JSM_SITEMAP  = f"{JSM_BASE}/job_listing-sitemap.xml"
CMW_BASE     = "https://careersmw.com"
CMW_MAX_PAGES = int(os.environ.get("CAREERSMW_MAX_PAGES", "40"))
CAD_BASE     = "https://www.careeradmw.com"
CAD_MAX_PAGES = int(os.environ.get("CAREERADMW_MAX_PAGES", "60"))

LI_LOCATION  = os.environ.get("LINKEDIN_LOCATION", "Malawi")
LI_KEYWORDS  = [k.strip() for k in os.environ.get("LINKEDIN_KEYWORDS", "").split(",")] if os.environ.get("LINKEDIN_KEYWORDS") else [""]
LI_TIME_RANGE = os.environ.get("LINKEDIN_TIME_RANGE", "r604800")
LI_MAX_PAGES = int(os.environ.get("LINKEDIN_MAX_PAGES", "0"))
LI_EXPIRY_DAYS = int(os.environ.get("LINKEDIN_DEFAULT_EXPIRY_DAYS", "30"))
LI_DELAY     = max(REQUEST_DELAY, 2.0)

OUTPUT_FILE        = "jobs_output.xlsx"
PROCESSED_IDS_FILE = "processed_jobs.csv"
FLAGGED_FILE       = "flagged_no_apply.csv"

_TRACKER_FIELDS = ["Job ID", "Job URL", "Job Title", "Company Name",
                   "Status", "Timestamp", "WP ID", "Source"]
_FLAGGED_FIELDS = ["Job ID", "Job URL", "Job Title", "Company Name",
                   "Reason", "Timestamp", "Source"]

# ── WordPress ────────────────────────────────────────────────────────────────
WP_URL      = os.environ.get("WP_BASE_URL", "").rstrip("/")
WP_USER     = os.environ.get("WP_USERNAME", "")
WP_PASSWORD = os.environ.get("WP_APP_PASSWORD", "")
INTERNAL_BOT_KEY = os.environ.get("INTERNAL_BOT_KEY", "")
WP_BASE      = WP_URL
WP_API_BASE  = f"{WP_BASE}/wp-json/wp/v2"     # reassigned by probe_wp_api() if needed
WP_JOBS_URL  = f"{WP_API_BASE}/job-listings"
WP_MEDIA_URL = f"{WP_API_BASE}/media"

# Cloudflare Bot Fight Mode allow-list header, only attached to requests to the WP host
if INTERNAL_BOT_KEY:
    _wp_host_hint = WP_URL.split("//")[-1].split("/")[0]
    _orig_session_request = requests.Session.request

    def _internal_auth_patched_request(self, method, url, *args, **kwargs):
        if _wp_host_hint and _wp_host_hint in str(url):
            headers = dict(kwargs.get("headers") or {})
            headers.setdefault("X-Internal-Auth", INTERNAL_BOT_KEY)
            kwargs["headers"] = headers
        return _orig_session_request(self, method, url, *args, **kwargs)

    requests.Session.request = _internal_auth_patched_request

# ── Mistral (optional) ───────────────────────────────────────────────────────
MISTRAL_API_KEY = os.environ.get("MISTRAL_API_KEY", "")
MISTRAL_MODEL   = "mistral-small-latest"
MISTRAL_URL     = "https://api.mistral.ai/v1/chat/completions"
ENABLE_PARAPHRASE = os.environ.get("ENABLE_PARAPHRASE", "1") != "0"

JOB_TYPE_MAPPING = {
    "full-time": "full-time", "full time": "full-time", "fulltime": "full-time",
    "permanent": "full-time", "other": "full-time",
    "part-time": "part-time", "part time": "part-time", "parttime": "part-time",
    "contract": "contract", "fixed contract": "contract", "fixed-term": "contract",
    "fixed term": "contract", "consultancy": "contract",
    "temporary": "temporary", "internship": "internship", "intern": "internship",
    "freelance": "freelance", "volunteer": "volunteer",
}

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Charset": "utf-8",
}

SESSION = requests.Session()
SESSION.headers.update(HEADERS)

_apply_url_cache = {}

# =============================================================================
#  LOGGING / COLOUR
# =============================================================================

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log_ = logging.getLogger("mw")

_USE_COLOUR = sys.stdout.isatty()

def _c(code, text):
    return f"\033[{code}m{text}\033[0m" if _USE_COLOUR else text

C_HEADER  = lambda t: _c("1;36", t)
C_LABEL   = lambda t: _c("1;33", t)
C_VALUE   = lambda t: _c("97",   t)
C_DIM     = lambda t: _c("2",    t)
C_GREEN   = lambda t: _c("1;32", t)
C_RED     = lambda t: _c("1;31", t)
C_BLUE    = lambda t: _c("1;34", t)
C_DIVIDER = lambda: _c("2", "─" * 80)

def log(msg=""):
    print(msg, flush=True)

# =============================================================================
#  REGEXES / CONSTANTS
# =============================================================================

TRACKING_PARAM_PREFIXES = ("utm_",)
TRACKING_PARAM_EXACT = {"fbclid", "gclid", "msclkid", "mc_cid", "mc_eid", "ref", "referrer"}

EMAIL_PATTERN    = re.compile(r"[A-Za-z0-9.+_%-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}")
META_REFRESH_PAT = re.compile(
    r'<meta[^>]+http-equiv=["\']refresh["\'][^>]*content=["\'][^"\'>]*url=([^"\'>]+)', re.I)
JS_REDIRECT_PAT  = re.compile(
    r'(?:window\.)?location(?:\.href)?\s*(?:=\s*|\.replace\(\s*)["\']([^"\']+)["\']', re.I)

BOILERPLATE_PATTERNS = [
    re.compile(r"go to method of application\s*[»>]*", re.I),
    re.compile(r"Read more about this company", re.I),
    re.compile(r"Subscribe to our newsletter.*", re.I),
]

NOISE_EMAIL_DOMAINS = ("example.com", "sentry.io", "w3.org", "schema.org",
                       "wixpress.com", "squarespace.com", "domain.com")
FAKE_LOCAL_RE = re.compile(
    r"^(name|user|email|mail|yourname|your[-_.]?email|sample|test|hello"
    r"|noreply|no[-_.]?reply|webmaster|example)$", re.I)

BAD_DOMAINS = ("linkedin.com", "google.com", "youtube.com", "facebook.com", "twitter.com",
               "x.com", "instagram.com", "t.co", "example.com", "w3.org", "sentry.io",
               "schema.org", "wa.me", "whatsapp.com", "tiktok.com")

_DEADLINE_LABEL_RE = re.compile(
    r'(?:the\s+)?'
    r'(?:closing\s+date(?:\s+for\s+(?:receipt|receiving)\s+(?:of\s+)?applications)?'
    r'|closure\s+date|application\s+deadline|apply\s+by'
    r'|deadline(?:\s+for\s+applications)?)'
    r'\s*(?:is\s+(?:(?:on\s+or\s+before|on)\s+)?|:\s*(?:(?:on\s+or\s+before|on)\s+)?|[:\-–]\s*)',
    re.I)
_DEADLINE_NOISE_RE = re.compile(
    r'^(?:the\s+)?(?:closing\s+date|closure\s+date|application\s+deadline)'
    r'(?:\s+for\s+(?:receipt|receiving)\s+(?:of\s+)?applications)?'
    r'\s*(?:is\s+(?:(?:on\s+or\s+before|on)\s+)?|:\s*)', re.I)
_DEADLINE_DATE_RE = re.compile(
    r'^((?:[A-Za-z]+,?\s+)?'
    r'(?:close\s+of\s+business[, ]+)?'
    r'(?:applications?\s+close\s+\S+\s+\S+\s+\S+\s+on\s+)?'
    r'(?:\d{1,2}(?:st|nd|rd|th)?\s+\w+,?\s+\d{4}'
    r'|\d{1,2}/\w+/\d{4}|\d{1,2}/\d{1,2}/\d{4}|\d{4}-\d{2}-\d{2}'
    r'|\w+\s+\d{1,2},?\s+\d{4}|\d{1,2}\s+\w+\s+\d{4})'
    r'(?:[, ]+(?:at\s+)?[\d:]+\s*(?:am|pm|CAT|EAT|GMT|UTC)?)?)', re.I)

_MOJIBAKE = [
    ("Â", ""), ("â€™", "'"), ("â€œ", '"'), ("â€\x9d", '"'), ("â€", '"'),
    ("â€¢", "•"), ("â„¢", "™"), ("\u00a0", " "), ("\u200b", ""), ("\ufeff", ""),
]

WP_RESIZE_PARAM_RE  = re.compile(r"[?&](?:resize|w|h|fit|crop)=[^&]+")
WP_RESIZE_SUFFIX_RE = re.compile(r"-\d+x\d+(\.[a-zA-Z]{2,5})$")
LOGO_KW_RE     = re.compile(r"logo", re.I)
PLACEHOLDER_RE = re.compile(r"default|placeholder|avatar|no-?image|blank|generic|company-placeholder|ghost|favicon", re.I)

_MONTHS = {m: i + 1 for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"])}

# =============================================================================
#  TEXT HELPERS
# =============================================================================

def _fix_mojibake(text):
    for pattern, replacement in _MOJIBAKE:
        text = text.replace(pattern, replacement)
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "", text)

def sanitize_text(text, is_url=False):
    if not isinstance(text, str):
        text = str(text) if (text is not None and str(text) not in ("nan", "None", "NaN")) else ""
    text = text.strip()
    if text in ("nan", "None", "NaN", "", "N/A", "n/a", "NA", "na"):
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

def html_to_text(el):
    """Block-aware text extraction (keeps paragraph/list line breaks)."""
    if el is None:
        return ""
    txt = el.get_text("\n", strip=True)
    txt = _fix_mojibake(txt)
    txt = re.sub(r"[ \t]+", " ", txt)
    return re.sub(r"\n{3,}", "\n\n", txt).strip()

def clean_description(text):
    if not text:
        return text
    for pattern in BOILERPLATE_PATTERNS:
        text = pattern.sub("", text)
    return re.sub(r"[ \t]+", " ", text).strip()

# ── Paragraph formatting ─────────────────────────────────────────────────────

_SECTION_HEADINGS = [
    "About the role", "About the Role", "About the job", "About us", "About Us",
    "About the company", "Job Summary", "Job summary", "Job Purpose", "Purpose of the job",
    "Duties and Responsibilities", "Duties & Responsibilities", "Key Responsibilities",
    "Responsibilities", "Main Duties", "Skills and Experience", "Skills & Experience",
    "Qualifications and Experience", "Qualifications", "Requirements", "Key Requirements",
    "Minimum Requirements", "Personal Attributes", "Competencies", "How to Apply",
    "How to apply", "Application Procedure", "Method of Application",
]
_HEAD_RE = re.compile(
    r"(?<![A-Za-z])(" + "|".join(re.escape(h) for h in
                                 sorted(_SECTION_HEADINGS, key=len, reverse=True)) + r")\b\s*:?\s*")

# Words that typically begin a new duty/requirement bullet in run-on text
_BULLET_STARTERS = (
    "Responsible", "Communicate", "Stock", "Reporting", "Report", "Training", "Train", "Sales",
    "Work", "Identify", "Regional", "Gather", "Market", "Other", "Experience", "Knowledge",
    "Strong", "High", "Excellent", "Enthusiastic", "Proven", "Qualified", "Ensure", "Manage",
    "Develop", "Prepare", "Conduct", "Coordinate", "Monitor", "Maintain", "Provide", "Support",
    "Lead", "Implement", "Oversee", "Assist", "Participate", "Perform", "Demonstrated",
    "Ability", "Bachelor", "Diploma", "Degree", "Certificate", "Minimum", "Must", "Good",
    "Proficiency", "Computer", "Fluency", "Self",
)
_BULLET_SPLIT_RE = re.compile(
    r"(?<=[a-z0-9\)\.,;:])\s+(?=(?:%s)\b)" % "|".join(_BULLET_STARTERS))
_SENT_SPLIT_RE = re.compile(r"(?<=[.!?])\s+(?=[A-Z•\-\d])")

def _is_heading(line):
    s = line.strip().rstrip(":").strip()
    return bool(s) and len(s.split()) <= 6 and s in _SECTION_HEADINGS

def format_paragraphs(text):
    """Turn a wall of text into paragraphs / headings / '• ' bullet lines separated by blank lines."""
    if not text:
        return text
    text = text.replace("\r", "")
    # headings get their own block (blank line before AND after)
    text = _HEAD_RE.sub(lambda m: f"\n\n{m.group(1)}\n\n", text)
    text = re.sub(r"[ \t]*[•●▪◦■]\s*", "\n• ", text)          # inline bullet glyphs
    blocks = [b.strip() for b in re.split(r"\n\s*\n", text) if b.strip()]
    out = []
    for block in blocks:
        lines = [l.strip() for l in block.split("\n") if l.strip()]
        if len(lines) > 1 or len(block) <= 350 or _is_heading(block):
            out.append("\n".join(lines))                      # already structured / short
            continue
        sentences = _SENT_SPLIT_RE.split(block)
        run_on = [s for s in sentences if len(s) > 250]
        if run_on and len(sentences) <= 3:                    # no punctuation: duty/requirement list
            items = []
            for s in sentences:
                items.extend(p.strip() for p in _BULLET_SPLIT_RE.split(s) if p.strip())
            if len(items) >= 3:
                out.append("\n".join("• " + i.lstrip("•- ").strip() for i in items))
                continue
        para, buf = [], ""                                    # normal prose: ~3 sentences per paragraph
        for s in sentences:
            if buf and (len(buf) + len(s) > 450):
                para.append(buf)
                buf = s
            else:
                buf = f"{buf} {s}".strip()
        if buf:
            para.append(buf)
        out.extend(para)
    return "\n\n".join(out)

def text_to_html(text):
    """Paragraph text -> <h3>/<p>/<ul><li> HTML for WordPress."""
    if not text:
        return ""
    parts = []
    for block in re.split(r"\n\s*\n", text.strip()):
        lines = [l.strip() for l in block.split("\n") if l.strip()]
        if not lines:
            continue
        if len(lines) == 1 and _is_heading(lines[0]):
            parts.append(f"<h3>{html_escape(lines[0].rstrip(':'))}</h3>")
        elif all(l.startswith("• ") for l in lines):
            parts.append("<ul>" + "".join(f"<li>{html_escape(l[2:].strip())}</li>" for l in lines) + "</ul>")
        else:
            parts.append("<p>" + "<br>".join(html_escape(l.lstrip("• ").strip()) for l in lines) + "</p>")
    return "\n".join(parts)

def extract_deadline_from_text(text):
    if not text:
        return ""
    m = _DEADLINE_LABEL_RE.search(text)
    if not m:
        return ""
    remainder = text[m.end():].strip()
    remainder = _DEADLINE_NOISE_RE.sub("", remainder).strip()
    dm = _DEADLINE_DATE_RE.match(remainder)
    if dm:
        return dm.group(1).strip().rstrip(",")
    fallback = re.split(r"[.\n]", remainder)[0].strip()
    return fallback[:80] if fallback else ""

def strip_tracking_params(url):
    if not url:
        return url
    parts = urlsplit(url)
    if not parts.query:
        return url
    kept = [(k, v) for k, v in parse_qsl(parts.query, keep_blank_values=True)
            if not k.lower().startswith(TRACKING_PARAM_PREFIXES) and k.lower() not in TRACKING_PARAM_EXACT]
    return urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(kept), parts.fragment))

def normalize_date(text):
    """Find a date inside free text and return YYYY-MM-DD ('' if none found)."""
    if not text:
        return ""
    t = re.sub(r"(\d)(st|nd|rd|th)\b", r"\1", str(text))
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", t)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    m = re.search(r"(\d{1,2})\s+([A-Za-z]{3,})\.?,?\s+(\d{4})", t)
    if m and m.group(2)[:3].lower() in _MONTHS:
        return f"{int(m.group(3)):04d}-{_MONTHS[m.group(2)[:3].lower()]:02d}-{int(m.group(1)):02d}"
    m = re.search(r"([A-Za-z]{3,})\.?\s+(\d{1,2}),?\s+(\d{4})", t)
    if m and m.group(1)[:3].lower() in _MONTHS:
        return f"{int(m.group(3)):04d}-{_MONTHS[m.group(1)[:3].lower()]:02d}-{int(m.group(2)):02d}"
    m = re.search(r"(\d{1,2})/([A-Za-z]{3})[a-z]*/(\d{4})", t)
    if m and m.group(2).lower() in _MONTHS:
        return f"{int(m.group(3)):04d}-{_MONTHS[m.group(2).lower()]:02d}-{int(m.group(1)):02d}"
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", t)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if mo > 12 and d <= 12:
            d, mo = mo, d
        if 1 <= mo <= 12 and 1 <= d <= 31:
            return f"{y:04d}-{mo:02d}-{d:02d}"
    return ""

def resolve_relative_date(text):
    """'3 days ago' / '2 weeks ago' / '5 hours ago' / 'Posted 1 month ago' -> YYYY-MM-DD."""
    if not text:
        return text
    t = text.strip()
    iso = normalize_date(t)
    if iso:
        return iso
    if re.search(r"just\s*now|today", t, re.I):
        return datetime.now().strftime("%Y-%m-%d")
    m = re.search(r"(\d+)\s*(minute|min|hour|hr|day|week|wk|month|year)s?\s+ago", t, re.I)
    if not m:
        return t
    n, unit = int(m.group(1)), m.group(2).lower()
    if unit in ("minute", "min", "hour", "hr"):
        delta = timedelta(hours=0)
    elif unit == "day":
        delta = timedelta(days=n)
    elif unit in ("week", "wk"):
        delta = timedelta(weeks=n)
    elif unit == "month":
        delta = timedelta(days=n * 30)
    else:
        delta = timedelta(days=n * 365)
    return (datetime.now() - delta).strftime("%Y-%m-%d")

def clean_email(raw):
    if not raw:
        return ""
    em = re.sub(r"^mailto:", "", raw.strip(), flags=re.I)
    em = unquote(em.split("?")[0]).strip().strip(".,;:()<>[]\"'").lower()
    if not EMAIL_PATTERN.fullmatch(em):
        return ""
    local, _, domain = em.partition("@")
    if FAKE_LOCAL_RE.match(local) or any(domain.endswith(d) for d in NOISE_EMAIL_DOMAINS):
        return ""
    return em

def extract_email(text):
    if not text:
        return ""
    for m in EMAIL_PATTERN.finditer(text):
        em = clean_email(m.group(0))
        if em:
            return em
    return ""

def is_linkedin_url(v):
    return bool(v) and bool(re.search(r"linkedin\.com", v, re.I))

def blank_if_linkedin(v):
    return "" if is_linkedin_url(v) else v

def is_bad_url(url):
    if not url or not url.startswith("http"):
        return True
    host = urlparse(url).netloc.lower()
    return any(host == d or host.endswith("." + d) for d in BAD_DOMAINS)

def host_of(url):
    try:
        return urlparse(url).netloc.lower().replace("www.", "")
    except Exception:
        return ""

def same_site(url, base):
    return host_of(url) == host_of(base)

def absolute_url(href, base):
    if not href:
        return ""
    return href if href.startswith("http") else urljoin(base, href)

# =============================================================================
#  HTTP HELPERS (all errors swallowed -> None, so sources can be skipped)
# =============================================================================

def fetch(url, retries=3, headers=None, params=None):
    """GET with retry/back-off. Returns Response (200) or None. Never raises."""
    last = "no attempt"
    for attempt in range(1, retries + 1):
        try:
            r = SESSION.get(url, timeout=REQUEST_TIMEOUT, headers=headers,
                            params=params, allow_redirects=True)
        except requests.RequestException as e:
            last = f"{type(e).__name__}: {e}"
            log(C_DIM(f"    fetch attempt {attempt}/{retries} failed: {last}"))
            time.sleep(min(2 * attempt, 8))
            continue
        sc = r.status_code
        if sc == 200:
            if not r.encoding or r.encoding.lower() == "iso-8859-1":
                r.encoding = "utf-8"
            return r
        last = f"HTTP {sc}"
        if sc in (400, 401, 403, 404, 410, 451, 999):
            log(C_RED(f"    SKIP {last}: {url}"))
            return None
        if sc == 429:
            wait = 15 * attempt
            log(C_RED(f"    rate-limited (429) — sleeping {wait}s: {url}"))
            time.sleep(wait)
            continue
        if sc >= 500:
            log(C_DIM(f"    {last} (attempt {attempt}/{retries}): {url}"))
            time.sleep(2 ** attempt)
            continue
        log(C_RED(f"    SKIP {last}: {url}"))
        return None
    log(C_RED(f"    GAVE UP after {retries} attempts ({last}): {url}"))
    return None

def get_soup(url, retries=3, headers=None):
    r = fetch(url, retries=retries, headers=headers)
    if r is None:
        return None
    try:
        return BeautifulSoup(r.text, _PARSER)
    except Exception as e:
        log(C_RED(f"    parse error {url}: {e}"))
        return None

def source_reachable(base_url):
    """Pre-flight check so a dead/blocked source is skipped immediately."""
    log(C_DIM(f"  pre-flight check: {base_url}"))
    return fetch(base_url, retries=2) is not None

def _find_html_redirect(html, current_url):
    m = META_REFRESH_PAT.search(html) or JS_REDIRECT_PAT.search(html)
    if not m:
        return ""
    return urljoin(current_url, m.group(1).strip().strip("'\""))

def resolve_apply_url(raw, base):
    if not raw:
        return ""
    if raw in _apply_url_cache:
        return _apply_url_cache[raw]
    resolved = ""
    try:
        resp = SESSION.get(raw, timeout=REQUEST_TIMEOUT, allow_redirects=True)
        final = resp.url
        if final and not same_site(final, base):
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

def resolve_application_contact(raw_apply_url, description, raw_anchor_text="", base=JSM_BASE):
    result = {"apply_url": "", "apply_email": "", "apply_raw": raw_apply_url}
    if raw_apply_url and re.match(r"^[A-Za-z0-9.+_-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$", raw_apply_url):
        result["apply_email"] = clean_email(raw_apply_url)
        return result
    if raw_apply_url and RESOLVE_APPLY_URLS:
        result["apply_url"] = resolve_apply_url(raw_apply_url, base)
    if not result["apply_url"]:
        result["apply_email"] = extract_email(raw_anchor_text) or extract_email(description)
    return result

def find_apply_in(el, text, base):
    """Generic apply finder: mailto link > external link > email in text."""
    if el is not None:
        for a in el.find_all("a", href=True):
            href = a["href"].strip()
            if href.lower().startswith("mailto:"):
                em = clean_email(href)
                if em:
                    return "", em, "mailto_link"
        for a in el.find_all("a", href=True):
            href = absolute_url(a["href"].strip(), base)
            if href.startswith("http") and not same_site(href, base) and not is_bad_url(href):
                return strip_tracking_params(href), "", "external_link"
    em = extract_email(text)
    if em:
        return "", em, "email_in_text"
    return "", "", "not_found"

# =============================================================================
#  LOGO HELPERS
# =============================================================================

def clean_logo_url(raw, base=""):
    if not raw:
        return ""
    raw = raw.strip()
    if raw.startswith("data:"):
        return ""
    if not raw.startswith("http"):
        raw = urljoin(base, raw) if base else ""
    raw = re.sub(r"[\"')\s]+$", "", raw)
    raw = WP_RESIZE_PARAM_RE.sub("", raw).rstrip("?&")
    return WP_RESIZE_SUFFIX_RE.sub(r"\1", raw)

def is_placeholder_logo(url):
    return not url or url.startswith("data:") or bool(PLACEHOLDER_RE.search(url))

def extract_logo_from_soup(soup, base_url, brand_re=None):
    for img in soup.find_all("img"):
        if img.find_parent(["nav", "footer"]):
            continue
        blob = " ".join(filter(None, [" ".join(img.get("class", []) or []),
                                      img.get("id", ""), img.get("alt", ""), img.get("src", "")]))
        if LOGO_KW_RE.search(blob):
            src = img.get("data-src") or img.get("data-lazy-src") or img.get("src") or ""
            cand = clean_logo_url(src, base_url)
            if cand and not is_placeholder_logo(cand) and not (brand_re and brand_re.search(cand)):
                return cand
    for prop in ("og:image", "twitter:image"):
        tag = soup.find("meta", attrs={"property": prop}) or soup.find("meta", attrs={"name": prop})
        if tag:
            cand = clean_logo_url(tag.get("content", ""), base_url)
            if cand and not is_placeholder_logo(cand):
                return cand
    link = soup.find("link", rel=lambda r: r and "apple-touch-icon" in " ".join(r if isinstance(r, list) else [r]).lower())
    if link and link.get("href"):
        cand = clean_logo_url(link["href"], base_url)
        if cand:
            return cand
    return ""

# =============================================================================
#  NLP / MISTRAL (optional paraphrasing — only runs if MISTRAL_API_KEY is set)
# =============================================================================

_sim_model = None

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
    for pat in [r"\[/?INST\]", r"</?s>", r"(?i)(rewritten?|rephrased?|output|paraphrase[d]?)[:\s]+",
                r"\*\*", r"###", r"---"]:
        text = re.sub(pat, "", text)
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()

def mistral_generate(prompt, max_tokens=400, temperature=0.7):
    if not MISTRAL_API_KEY:
        return ""
    try:
        response = requests.post(
            MISTRAL_URL,
            headers={"Authorization": f"Bearer {MISTRAL_API_KEY}", "Content-Type": "application/json"},
            json={"model": MISTRAL_MODEL, "messages": [{"role": "user", "content": prompt}],
                  "max_tokens": max_tokens, "temperature": temperature},
            timeout=30)
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
        log(C_DIM(f"      title attempt {attempt+1}: '{result}' words={wc} sim={sim:.3f} valid={valid}"))
        if valid and sim > best_sim:
            best_sim, best_result = sim, result
        time.sleep(1)
    return best_result if best_result else clean

def paraphrase_description(text):
    """Paraphrase block by block, keeping headings, bullets and paragraph breaks intact."""
    if not ENABLE_PARAPHRASE or not MISTRAL_API_KEY:
        return text
    clean = sanitize_text(text)
    if not clean:
        return text
    blocks = [b.strip() for b in re.split(r"\n\s*\n", clean) if b.strip()] or [clean]
    rewritten = []
    for i, block in enumerate(blocks, 1):
        if _is_heading(block):
            rewritten.append(block)
            continue
        has_bullets = block.lstrip().startswith("•")
        if has_bullets:
            shape = 'Keep the same bullet list: one item per line, each line starting with "• ". '
        else:
            shape = "Keep it as a single paragraph. "
        prompt = ("Rewrite this job description section professionally. "
                  "Keep ALL facts, requirements, and responsibilities. "
                  "Use different sentence structure and vocabulary. "
                  + shape +
                  "Output ONLY the rewritten text — no labels, no explanation.\n\n"
                  f"Original:\n{block}")
        best_result, best_sim, accepted = None, 0.0, None
        for attempt in range(3):
            raw    = mistral_generate(prompt, max_tokens=700, temperature=round(0.65 + attempt * 0.08, 2))
            result = clean_output(raw).strip()
            rw     = len(result.split()) if result else 0
            sim    = similarity_score(block, result) if result and rw >= 5 else 0.0
            valid  = bool(result) and rw >= 8 and sim >= 0.48
            if has_bullets and valid and not result.lstrip().startswith("•"):
                valid = False
            log(C_DIM(f"      block {i}/{len(blocks)} attempt {attempt+1}: words={rw} sim={sim:.3f} valid={valid}"))
            if valid:
                accepted = result
                break
            if result and sim > best_sim:
                best_sim, best_result = sim, result
            time.sleep(1)
        rewritten.append(accepted or (best_result if best_result and best_sim >= 0.40 else block))
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
    time.sleep(1)
    return result if result and len(result.split()) >= 10 else clean

# =============================================================================
#  TRACKER (stdlib csv only)
# =============================================================================

def _init_tracker():
    for path, fields in ((PROCESSED_IDS_FILE, _TRACKER_FIELDS), (FLAGGED_FILE, _FLAGGED_FIELDS)):
        if not os.path.exists(path):
            try:
                with open(path, "w", newline="", encoding="utf-8") as f:
                    csv.writer(f).writerow(fields)
            except Exception as e:
                log_.error(f"Could not create {path}: {e}")

def load_processed_ids():
    _init_tracker()
    ids, urls = set(), set()
    try:
        with open(PROCESSED_IDS_FILE, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                if row.get("Job ID"):
                    ids.add(row["Job ID"].strip())
                if row.get("Job URL"):
                    urls.add(row["Job URL"].strip())
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
        if row.get("Job ID", "").strip() == str(job_id):
            row.update(updates)
            row["Timestamp"] = datetime.now().isoformat()
            found = True
            break
    if not found:
        new_row = {k: "" for k in _TRACKER_FIELDS}
        new_row["Job ID"] = str(job_id)
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

def flag_job(job_id, job_url, title, company, source, reason="no_apply_contact"):
    log_.warning(f"FLAGGED [{reason}]: {title} | {job_url}")
    rows = []
    try:
        with open(FLAGGED_FILE, newline="", encoding="utf-8") as f:
            rows = list(csv.DictReader(f))
    except Exception:
        pass
    if any(r.get("Job ID", "").strip() == str(job_id) for r in rows):
        return
    rows.append({"Job ID": job_id, "Job URL": job_url, "Job Title": title,
                 "Company Name": company, "Reason": reason,
                 "Timestamp": datetime.now().isoformat(), "Source": source})
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

def mark_scraped(job_id, job_url, title, company, source):
    _upsert_row(job_id, {"Job URL": job_url, "Job Title": title, "Company Name": company,
                         "Status": "scraped", "WP ID": "", "Source": source})

def mark_paraphrased(job_id):
    _upsert_row(job_id, {"Status": "paraphrased"})

def mark_posted(job_id, wp_id):
    _upsert_row(job_id, {"Status": "posted", "WP ID": str(wp_id)})

def mark_failed(job_id, reason):
    _upsert_row(job_id, {"Status": f"failed|{reason}"})

# =============================================================================
#  SOURCE 1 — jobsearchmalawi.com
# =============================================================================

def collect_jobsearchmalawi():
    log(f"  Sitemap: {JSM_SITEMAP}")
    urls, seen = [], set()

    def _read(sm_url, depth=0):
        r = fetch(sm_url, retries=2)
        if r is None:
            return
        for loc in re.findall(r"<loc>\s*([^<\s]+)\s*</loc>", r.text):
            if loc in seen:
                continue
            seen.add(loc)
            if loc.endswith(".xml") and depth < 2:
                _read(loc, depth + 1)
            elif not loc.endswith(".xml"):
                urls.append(loc)

    _read(JSM_SITEMAP)
    if not urls:
        for alt in (f"{JSM_BASE}/wp-sitemap-posts-job_listing-1.xml", f"{JSM_BASE}/sitemap_index.xml"):
            _read(alt)
        urls = [u for u in urls if "/job/" in u or "/job-" in u or "/jobs/" in u] or urls
    log(f"  Found {len(urls)} job URLs")
    return urls

def parse_jobsearchmalawi(url):
    soup = get_soup(url)
    if soup is None:
        return None

    title = ""
    for sel in ["h1.entry-title.job_title", "h1.job_title", "h1.entry-title", ".job_title", "h1"]:
        el = soup.select_one(sel)
        if el:
            title = clean_text(el)
            if title:
                break
    if not title:
        og = soup.find("meta", property="og:title")
        if og:
            title = re.sub(r"\s*[-|–]\s*[A-Za-z ]+$", "", og.get("content", "").strip()).strip()

    job_type = ""
    for sel in ["li.job-type", "li.type", ".job-type"]:
        el = soup.select_one(sel)
        if el:
            job_type = re.sub(r"^(job\s*type|type)[:\s]+", "", clean_text(el), flags=re.I).strip()
            if job_type:
                break

    location = ""
    for sel in ["li.location", ".job-location", "li.job-location"]:
        el = soup.select_one(sel)
        if el:
            location = re.sub(r"^(location)[:\s]+", "", clean_text(el), flags=re.I).strip()
            if location:
                break

    date_posted = ""
    for sel in ["li.date-posted", ".date-posted", "li.posted-date", "time[datetime]"]:
        el = soup.select_one(sel)
        if el:
            if el.name == "time" and el.get("datetime"):
                date_posted = el["datetime"][:10]
            else:
                date_posted = re.sub(r"^(date\s*posted|posted|date)[:\s]+", "", clean_text(el), flags=re.I).strip()
            if date_posted:
                break
    date_posted = resolve_relative_date(date_posted)

    deadline = ""
    for sel in ["div.cfwjm_output", ".job-deadline", ".deadline", "li.closing-date", "li.job-expiry", ".job-expiry"]:
        el = soup.select_one(sel)
        if el:
            val = re.sub(r"^(closes?\s*:?|closing\s*date\s*:?|deadline\s*:?|application\s*deadline\s*:?|"
                         r"apply\s*by\s*:?|expir(?:y|es)\s*:?)\s*", "", clean_text(el), flags=re.I).strip()
            if val:
                deadline = val
                break

    company = ""
    for sel in [".company .company_name", ".company strong.name", ".company p.name", ".company_name",
                "strong.name", "p.name", ".company-name", "li.company",
                ".single-job-listing .company h3", ".single-job-listing .company p", ".job-company"]:
        el = soup.select_one(sel)
        if el:
            cand = clean_text(el)
            if cand and len(cand) < 120 and not re.search(
                    r"\b(required|responsibilities|qualifications|apply|salary)\b", cand, re.I):
                company = cand
                break
    if not company:
        og_site = soup.find("meta", property="og:site_name")
        if og_site:
            val = og_site.get("content", "").strip()
            if val and val.lower() not in ("jobsearchmalawi", "jobsearch malawi", "jobs in malawi", "malawi jobs"):
                company = val

    logo = ""
    company_section = soup.select_one(".company")
    if company_section:
        img_el = company_section.select_one("img")
        if img_el:
            logo = clean_logo_url(img_el.get("src") or img_el.get("data-src") or "", JSM_BASE)
    if not logo or is_placeholder_logo(logo):
        logo = extract_logo_from_soup(soup, JSM_BASE, re.compile(r"jobsearchmalawi", re.I))

    company_details = ""
    for sel in [".company-description", ".about-company", "#company-description", ".company-profile", "div.company"]:
        el = soup.select_one(sel)
        if el:
            txt = clean_text(el)
            if txt and len(txt) > 40:
                company_details = txt
                break
    if not company_details:
        for h in soup.find_all(re.compile(r"^h[2-4]$")):
            if re.search(r"about", h.get_text(), re.I):
                sib = h.find_next_sibling()
                if sib:
                    company_details = clean_text(sib)
                break

    # Description: keep block/line structure (html_to_text), NOT clean_text (which flattens it)
    desc_el = soup.select_one(".job_description") or soup.select_one("div.job_description")
    if desc_el:
        description = clean_description(html_to_text(desc_el))
    else:
        description, best = "", None
        for div in soup.find_all("div"):
            t = clean_text(div)
            if len(t) > len(description) and len(t) > 200:
                description, best = t, div
        description = clean_description(html_to_text(best)) if best else ""

    if not deadline:
        deadline = extract_deadline_from_text(description)

    raw_apply, raw_anchor_txt = "", ""
    raw_html = str(soup)
    m_url = re.search(r'To apply for this job please visit\s*<a href="([^"]+)"', raw_html, re.I)
    m_email = re.search(r'To apply for this job.*?email.*?to\s*<a[^>]*class=["\']job_application_email["\'][^>]*>(.+?)</a>',
                        raw_html, re.I | re.S)
    if m_url:
        raw_apply = m_url.group(1).strip()
    elif m_email:
        raw_anchor_txt = re.sub(r"<[^>]+>", "", m_email.group(1)).strip()
        raw_apply = raw_anchor_txt
    if not raw_apply:
        a_el = (soup.select_one("a.job_application_email") or soup.select_one(".application-link a")
                or soup.select_one(".application a"))
        if a_el:
            href = a_el.get("href", "")
            if href.startswith("mailto:"):
                raw_apply = re.sub(r"^mailto:", "", href).strip()
            else:
                raw_apply = href
            raw_anchor_txt = clean_text(a_el)

    log(f"    Resolving apply link for '{title}' …")
    application = resolve_application_contact(raw_apply, description, raw_anchor_txt, JSM_BASE)

    return {
        "title": title, "job_url": url, "job_type": job_type, "location": location,
        "posted_date": date_posted, "deadline": deadline, "description": description,
        "apply_url": application["apply_url"], "apply_email": application["apply_email"],
        "apply_raw": application["apply_raw"], "apply_method": "",
        "company_name": company, "company_details": company_details, "company_logo": logo,
        "company_website": "", "categories": [],
    }

# =============================================================================
#  SOURCE 2 — careersmw.com
# =============================================================================

_CMW_JOB_RE = re.compile(r"^https?://(?:www\.)?careersmw\.com/job/[^/?#]+/?$", re.I)

def collect_careersmw():
    urls, seen = [], set()

    def add(u):
        u = u.strip()
        if u and _CMW_JOB_RE.match(u) and u not in seen:
            seen.add(u)
            urls.append(u)

    # Strategy 1 — WP REST API
    log("  strategy 1: WP REST API (job_listing)")
    for page in range(1, CMW_MAX_PAGES + 1):
        r = fetch(f"{CMW_BASE}/wp-json/wp/v2/job_listing",
                  retries=1, headers=None, params={"per_page": 100, "page": page, "_fields": "link"})
        if r is None:
            break
        try:
            data = r.json()
        except Exception:
            break
        if not isinstance(data, list) or not data:
            break
        for it in data:
            add(it.get("link", "") if isinstance(it, dict) else "")
        log(C_DIM(f"    REST page {page}: total={len(urls)}"))
        if len(data) < 100:
            break
        time.sleep(REQUEST_DELAY)

    # Strategy 2 — sitemaps
    if not urls:
        log("  strategy 2: sitemaps")
        for sm in (f"{CMW_BASE}/job_listing-sitemap.xml", f"{CMW_BASE}/wp-sitemap-posts-job_listing-1.xml",
                   f"{CMW_BASE}/sitemap_index.xml", f"{CMW_BASE}/sitemap.xml"):
            r = fetch(sm, retries=1)
            if r is None:
                continue
            locs = re.findall(r"<loc>\s*([^<\s]+)\s*</loc>", r.text)
            for loc in locs:
                if loc.endswith(".xml") and ("job" in loc.lower() or "listing" in loc.lower()):
                    r2 = fetch(loc, retries=1)
                    if r2 is not None:
                        for l2 in re.findall(r"<loc>\s*([^<\s]+)\s*</loc>", r2.text):
                            add(l2)
                else:
                    add(loc)
            if urls:
                break

    # Strategy 3 — homepage / "all jobs" listing
    if not urls:
        log("  strategy 3: homepage listing")
        for page_url in (f"{CMW_BASE}/?search_keywords&search_location&search_category&job_all=1", CMW_BASE):
            soup = get_soup(page_url, retries=2)
            if soup is None:
                continue
            for a in soup.find_all("a", href=True):
                add(absolute_url(a["href"], CMW_BASE))
            if urls:
                break

    log(f"  Found {len(urls)} job URLs")
    return urls

def parse_careersmw(url):
    soup = get_soup(url)
    if soup is None:
        return None

    h1 = soup.select_one("h1.entry-title") or soup.select_one("h1")
    title = clean_text(h1)
    if not title:
        return None

    article = soup.select_one("article.job-single") or soup.select_one("article")
    classes = article.get("class", []) if article else []
    job_type = next((c.split("-", 1)[1].replace("-", " ") for c in classes if c.startswith("job_listing_type-")), "")
    categories = [c.split("-", 1)[1].replace("-", " ").title() for c in classes if c.startswith("job_listing_category-")]

    company, location = "", ""
    line = soup.select_one(".cmw-cjm-employer-line span")
    if line:
        txt = clean_text(line)
        if "·" in txt:
            company, location = [p.strip() for p in txt.split("·", 1)]
        else:
            company = txt

    logo = ""
    img = soup.select_one("img.cmw-cjm-employer-logo")
    if img:
        logo = clean_logo_url(img.get("data-src") or img.get("src") or "", CMW_BASE)

    posted, deadline = "", ""
    for sp in soup.select(".cmw-cjm-single-dates span"):
        t = clean_text(sp)
        if re.search(r"expire|closing|deadline", t, re.I):
            deadline = normalize_date(t) or t
        elif re.search(r"posted", t, re.I):
            posted = normalize_date(t) or t

    desc_el = soup.select_one(".cmw-cjm-description") or soup.select_one(".entry-content")
    description = clean_description(html_to_text(desc_el))
    poster_imgs = []
    if desc_el:
        for im in desc_el.find_all("img"):
            src = clean_logo_url(im.get("data-src") or im.get("src") or "", CMW_BASE)
            if src:
                poster_imgs.append(src)
    if len(description) < 60:
        # image-only advert: build a minimal description so the post is still valid
        poster_note = f" Full details are in the job advert image: {poster_imgs[0]}" if poster_imgs else ""
        description = (f"{title} at {company or 'the employer'}"
                       f"{' (' + location + ')' if location else ''}. "
                       f"Please refer to the official job advert for the full requirements and responsibilities."
                       f"{poster_note}").strip()

    apply_box = soup.select_one(".cmw-cjm-external-application")
    apply_url, apply_email, method = find_apply_in(apply_box, apply_box.get_text(" ") if apply_box else "", CMW_BASE)
    if not apply_url and not apply_email:
        apply_url, apply_email, method = find_apply_in(desc_el, description, CMW_BASE)
    if not deadline:
        deadline = extract_deadline_from_text(description)

    return {
        "title": title, "job_url": url, "job_type": job_type, "location": location,
        "posted_date": posted, "deadline": deadline, "description": description,
        "apply_url": apply_url, "apply_email": apply_email, "apply_raw": apply_url or apply_email,
        "apply_method": method, "company_name": company, "company_details": "",
        "company_logo": logo, "company_website": "", "categories": categories,
    }

# =============================================================================
#  SOURCE 3 — careeradmw.com
# =============================================================================

_CAD_JOB_RE = re.compile(r"^https?://(?:www\.)?careeradmw\.com/jobs/[^/?#]+/?$", re.I)

def collect_careeradmw():
    urls, seen = [], set()
    empty_streak = 0
    for page in range(1, CAD_MAX_PAGES + 1):
        page_url = f"{CAD_BASE}/jobs" if page == 1 else f"{CAD_BASE}/jobs?page={page}"
        soup = get_soup(page_url, retries=2)
        if soup is None:
            log(C_RED(f"  listing page {page} unavailable — stopping pagination"))
            break
        new = 0
        for a in soup.find_all("a", href=True):
            u = absolute_url(a["href"].strip(), CAD_BASE).split("#")[0]
            if _CAD_JOB_RE.match(u) and u not in seen:
                seen.add(u)
                urls.append(u)
                new += 1
        log(C_DIM(f"  listing page {page}: +{new} (total {len(urls)})"))
        if new == 0:
            empty_streak += 1
            if empty_streak >= 1:
                break
        else:
            empty_streak = 0
        time.sleep(REQUEST_DELAY)
    log(f"  Found {len(urls)} job URLs")
    return urls

def parse_careeradmw(url):
    soup = get_soup(url)
    if soup is None:
        return None

    h1 = soup.select_one("h1")
    title = clean_text(h1)
    if not title:
        return None

    company = ""
    el = soup.select_one("p.truncate.font-semibold") or soup.select_one("section p.text-lg")
    if el:
        company = clean_text(el)

    posted = ""
    for p in soup.select("p.text-xs"):
        t = clean_text(p)
        if t.lower().startswith("posted"):
            posted = normalize_date(t) or t
            break

    location, job_type, deadline = "", "", ""
    for i in soup.select("i.fa-map-marker-alt"):
        if i.parent is not None and i.parent.name == "span":
            location = clean_text(i.parent)
            break
    for i in soup.select("i.fa-briefcase"):
        if i.parent is not None and i.parent.name == "span":
            job_type = clean_text(i.parent)
            break
    for sp in soup.find_all("span"):
        t = clean_text(sp)
        m = re.match(r"apply by\s+(.+)", t, re.I)
        if m:
            deadline = normalize_date(m.group(1)) or m.group(1).strip()
            break

    desc_el = soup.select_one("div.prose")
    description = clean_description(html_to_text(desc_el))
    if not description:
        return None
    if not deadline:
        deadline = extract_deadline_from_text(description)

    logo = ""
    for im in soup.select("img"):
        src = im.get("src") or ""
        if src.startswith("http") and im.get("alt", "").strip() == company and company:
            cand = clean_logo_url(src, CAD_BASE)
            if cand and not is_placeholder_logo(cand):
                logo = cand
                break

    categories = [clean_text(a) for a in soup.select('a[href*="jobs?category="]') if clean_text(a)]

    apply_url, apply_email, method = find_apply_in(desc_el, description, CAD_BASE)

    return {
        "title": title, "job_url": url, "job_type": job_type, "location": location,
        "posted_date": posted, "deadline": deadline, "description": description,
        "apply_url": apply_url, "apply_email": apply_email, "apply_raw": apply_url or apply_email,
        "apply_method": method, "company_name": company, "company_details": "",
        "company_logo": logo, "company_website": "", "categories": categories[:4],
    }

# =============================================================================
#  SOURCE 4 — LinkedIn (guest API)
# =============================================================================

_LI_UAS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
]
_li_ua_idx = 0

def _li_headers():
    global _li_ua_idx
    ua = _LI_UAS[_li_ua_idx % len(_LI_UAS)]
    _li_ua_idx += 1
    return {"User-Agent": ua, "Accept-Language": "en-US,en;q=0.9",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"}

def li_canonical(url):
    m = re.search(r"/jobs/view/(?:[^/?#]*-)?(\d+)", url)
    if m:
        return f"https://www.linkedin.com/jobs/view/{m.group(1)}/"
    return re.sub(r"[?#].*$", "", url)

def li_decode_url(raw):
    """Unwrap LinkedIn redirect wrappers to the real external URL ('' if none)."""
    if not raw:
        return ""
    raw = raw.replace("&amp;", "&")
    if raw.startswith("http") and "linkedin.com" not in raw:
        return raw
    m = re.search(r"[?&]url=([^&]+)", raw)
    if m:
        d = unquote(m.group(1))
        if "%" in d:
            d = unquote(d)
        if d.startswith("http") and "linkedin.com" not in d:
            return d
    b = re.search(r"[?&]offsiteApplyUrl=([^&]+)", raw)
    if b:
        try:
            p = json.loads(base64.b64decode(unquote(b.group(1))).decode("utf-8"))
            if p and p.get("url"):
                return p["url"]
        except Exception:
            pass
    return ""

def collect_linkedin():
    seen, urls = set(), []
    for ki, keyword in enumerate(LI_KEYWORDS, 1):
        label = keyword or "(all)"
        log(f"  keyword {ki}/{len(LI_KEYWORDS)}: {label}  location={LI_LOCATION}")
        page, empty = 0, 0
        while True:
            if LI_MAX_PAGES and page >= LI_MAX_PAGES:
                break
            start = page * 25
            api = ("https://www.linkedin.com/jobs-guest/jobs/api/seeMoreJobPostings/search"
                   f"?location={quote_plus(LI_LOCATION)}&f_TPR={LI_TIME_RANGE}"
                   f"&keywords={quote_plus(keyword)}&start={start}")
            time.sleep(LI_DELAY)
            r = fetch(api, retries=3, headers=_li_headers())
            if r is None or not r.text.strip():
                log(C_DIM(f"    [{label}] page {page+1}: no more results / blocked"))
                break
            found = []
            for href in re.findall(r'href="(https?://[^"]*?/jobs/view/[^"]+?)"', r.text):
                c = li_canonical(href.replace("&amp;", "&"))
                if c not in seen:
                    seen.add(c)
                    found.append(c)
            log(C_DIM(f"    [{label}] page {page+1} (start={start}): +{len(found)} (total {len(urls) + len(found)})"))
            urls.extend(found)
            if found:
                empty = 0
            else:
                empty += 1
                if empty >= 2:
                    break
            if start >= 975:
                break
            page += 1
            if page % 10 == 0:
                log(C_DIM("    pausing 20s (every 10 pages) …"))
                time.sleep(20)
    log(f"  Found {len(urls)} job URLs")
    return urls

def _li_jsonld(html):
    for raw in re.findall(r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>', html, re.S | re.I):
        try:
            data = json.loads(raw.strip())
        except Exception:
            continue
        if isinstance(data, list):
            data = next((d for d in data if isinstance(d, dict) and d.get("@type") == "JobPosting"), {})
        if isinstance(data, dict) and data.get("@type") == "JobPosting":
            return data
    return {}

def _li_company_info(company_url):
    info = {"website": "", "logo": "", "about": "", "name": ""}
    if not company_url:
        return info
    base = re.sub(r"\?.*$", "", company_url.rstrip("/"))
    guest = base.replace("/company/", "/company-guest/")
    r = fetch(guest, retries=1, headers=_li_headers()) if guest != base else None
    if r is None:
        r = fetch(base, retries=1, headers=_li_headers())
    if r is None:
        return info
    soup = BeautifulSoup(r.text, _PARSER)
    ws = soup.select_one("a[data-tracking-control-name='about_website']")
    if ws:
        info["website"] = blank_if_linkedin(li_decode_url(ws.get("href", "")))
    og = soup.find("meta", property="og:image")
    if og:
        cand = clean_logo_url(og.get("content", ""))
        if cand and not is_placeholder_logo(cand):
            info["logo"] = cand
    ab = soup.select_one("section.about-us p, .core-section-container__content p, .org-about-us-organization-description__text")
    if ab:
        info["about"] = clean_text(ab)
    h1 = soup.select_one("h1")
    if h1:
        info["name"] = clean_text(h1).split(" | ")[0].strip()
    return info

def _crawl_company_site(website):
    """Light crawl: logo + a careers/contact email from the employer's own site."""
    out = {"logo": "", "email": ""}
    if not website or is_bad_url(website):
        return out
    root = website.rstrip("/")
    r = fetch(root, retries=1)
    if r is None:
        return out
    soup = BeautifulSoup(r.text, _PARSER)
    out["logo"] = extract_logo_from_soup(soup, root)
    def mail_from(s):
        for a in s.find_all("a", href=True):
            if a["href"].lower().startswith("mailto:"):
                em = clean_email(a["href"])
                if em:
                    return em
        return extract_email(s.get_text(" "))
    out["email"] = mail_from(soup)
    if not out["email"]:
        for a in soup.find_all("a", href=True):
            txt = a.get_text(" ", strip=True).lower()
            href = absolute_url(a["href"], root)
            if same_site(href, root) and re.search(r"career|vacanc|job|contact", txt + href.lower()):
                r2 = fetch(href, retries=1)
                if r2 is not None:
                    out["email"] = mail_from(BeautifulSoup(r2.text, _PARSER))
                    if out["email"]:
                        break
    return out

def parse_linkedin(url):
    r = fetch(url, retries=2, headers=_li_headers())
    if r is None:
        return None
    html = r.text
    soup = BeautifulSoup(html, _PARSER)

    def sel_text(*sels):
        for s in sels:
            el = soup.select_one(s)
            if el:
                t = el.get_text(strip=True)
                if t:
                    return t
        return ""

    title = sel_text(".top-card-layout__title", "h1.topcard__title",
                     ".job-details-jobs-unified-top-card__job-title", "h1")
    if not title or title.lower() in ("sign in", "linkedin"):
        log(C_RED("    LinkedIn returned a login wall / empty page"))
        return None

    ld = _li_jsonld(html)
    org = (ld.get("hiringOrganization") or {}) if isinstance(ld.get("hiringOrganization"), dict) else {}

    company = sel_text(".topcard__org-name-link", ".job-details-jobs-unified-top-card__company-name",
                       ".topcard__flavor") or org.get("name", "")
    if company.strip().lower() in ("sign in", "linkedin", "join now", ""):
        m = re.search(r"-at-([a-z0-9\-]+)-\d+/?", url, re.I)
        company = m.group(1).replace("-", " ").title() if m else ""
    company_el = (soup.select_one(".topcard__org-name-link")
                  or soup.select_one(".job-details-jobs-unified-top-card__company-name a"))
    company_url = company_el.get("href", "") if company_el else ""

    location = sel_text(".topcard__flavor--bullet", ".job-details-jobs-unified-top-card__bullet")
    time_el = soup.find("time")
    raw_posted = (time_el.get("datetime", "") if time_el else "") or \
        sel_text(".posted-time-ago__text", ".job-details-jobs-unified-top-card__posted-date")
    posted = resolve_relative_date(raw_posted) or ld.get("datePosted", "")[:10]

    desc_el = soup.select_one(".show-more-less-html__markup") or soup.select_one(".description__text")
    description = clean_description(html_to_text(desc_el))
    if not description and ld.get("description"):
        description = clean_description(html_to_text(BeautifulSoup(ld["description"], _PARSER)))
    if not description:
        log(C_RED("    LinkedIn job has no description — skipped"))
        return None

    job_type = ""
    for li in soup.select(".description__job-criteria-list > li"):
        h3 = li.find("h3")
        if h3 and "employment type" in h3.get_text().lower():
            sp = li.select(".description__job-criteria-text, span")
            if sp:
                job_type = sp[-1].get_text(strip=True)
    if not job_type:
        et = ld.get("employmentType", "")
        job_type = (et[0] if isinstance(et, list) and et else et) or ""
        job_type = str(job_type).replace("_", " ").title()

    deadline = normalize_date(ld.get("validThrough", "")) or extract_deadline_from_text(description)
    if not deadline and posted and re.match(r"\d{4}-\d{2}-\d{2}", posted):
        try:
            deadline = (datetime.strptime(posted[:10], "%Y-%m-%d") + timedelta(days=LI_EXPIRY_DAYS)).strftime("%Y-%m-%d")
        except Exception:
            pass

    # company enrichment (LinkedIn company page -> employer website)
    time.sleep(0.5)
    cinfo = _li_company_info(company_url)
    website = blank_if_linkedin(cinfo["website"]
                                or (org.get("sameAs", "") if isinstance(org.get("sameAs"), str) else "")
                                or org.get("url", ""))
    logo = cinfo["logo"]
    if not logo:
        lg = org.get("logo", "")
        lg = lg if isinstance(lg, str) else (lg.get("url", "") if isinstance(lg, dict) else "")
        logo = clean_logo_url(lg)
    if is_placeholder_logo(logo):
        logo = ""
    if is_bad_url(website):
        website = ""

    # apply discovery
    apply_url, apply_email, method = "", "", "not_found"
    cand = blank_if_linkedin(li_decode_url(ld.get("url", "")) if ld.get("url") else "")
    if cand and not is_bad_url(cand):
        apply_url, method = cand, "jsonld_url"
    if not apply_url:
        for a in soup.find_all("a", href=True):
            ctrl = (a.get("data-tracking-control-name") or "").lower()
            if "apply" in ctrl or "offsite" in ctrl:
                c = li_decode_url(a["href"])
                if c and not is_bad_url(c):
                    apply_url, method = c, "apply_button"
                    break
    if not apply_url:
        for script in soup.find_all("script"):
            txt = script.string or ""
            m = re.search(r'"(?:applyStartUrl|applicationUrl)"\s*:\s*"([^"]+)"', txt)
            if m:
                c = blank_if_linkedin(m.group(1).replace("\\u0026", "&").replace("\\", ""))
                if c.startswith("http") and not is_bad_url(c):
                    apply_url, method = c, "script_tag"
                    break
    if not apply_url:
        apply_url, apply_email, method = find_apply_in(desc_el, description, "https://www.linkedin.com")
    if not apply_url and not apply_email and website:
        log(f"    crawling employer site for apply contact: {website}")
        site = _crawl_company_site(website)
        if site["email"]:
            apply_email, method = site["email"], "company_site_email"
        if site["logo"] and not logo:
            logo = site["logo"]
    if not logo and website:
        logo = _crawl_company_site(website)["logo"]

    apply_url = blank_if_linkedin(apply_url)

    return {
        "title": title, "job_url": url, "job_type": job_type, "location": location,
        "posted_date": posted, "deadline": deadline, "description": description,
        "apply_url": apply_url, "apply_email": apply_email, "apply_raw": apply_url or apply_email,
        "apply_method": method, "company_name": company or cinfo["name"],
        "company_details": cinfo["about"], "company_logo": logo,
        "company_website": website, "categories": [],
    }

# =============================================================================
#  SOURCE REGISTRY
# =============================================================================

SOURCES = {
    "jobsearchmalawi": {"label": "JobSearchMalawi", "base": JSM_BASE,
                        "collect": collect_jobsearchmalawi, "parse": parse_jobsearchmalawi, "delay": REQUEST_DELAY},
    "careersmw":       {"label": "CareersMW", "base": CMW_BASE,
                        "collect": collect_careersmw, "parse": parse_careersmw, "delay": REQUEST_DELAY},
    "careeradmw":      {"label": "CareerAdMW", "base": CAD_BASE,
                        "collect": collect_careeradmw, "parse": parse_careeradmw, "delay": REQUEST_DELAY},
    "linkedin":        {"label": "LinkedIn", "base": "https://www.linkedin.com/jobs",
                        "collect": collect_linkedin, "parse": parse_linkedin, "delay": LI_DELAY},
}

# =============================================================================
#  DEDUPLICATE + PARAPHRASE
# =============================================================================

def process_job(raw_job, source_label, processed_ids, processed_urls, seen_content):
    """Returns (job_dict_or_None, status) where status in ok | dup | flagged."""
    job_url = raw_job.get("job_url", "")
    title   = raw_job.get("title", "")
    company = raw_job.get("company_name", "")
    job_id  = make_job_id(job_url, title, company)

    if job_id in processed_ids or job_url in processed_urls:
        log(C_DIM(f"  ⧳ Already processed — skipped: {job_url}"))
        return None, "dup"

    fp = (title.lower().strip(), company.lower().strip(), raw_job.get("location", "").lower().strip())
    if fp in seen_content:
        log(C_DIM(f"  ⧳ Duplicate (same title/company/location) — skipped: {title}"))
        mark_scraped(job_id, job_url, title, company, source_label)
        _upsert_row(job_id, {"Status": "duplicate"})
        processed_ids.add(job_id)
        processed_urls.add(job_url)
        return None, "dup"
    seen_content.add(fp)

    apply_url   = raw_job.get("apply_url", "")
    apply_email = raw_job.get("apply_email", "")
    fallback_to_source = False
    if not apply_url and not apply_email:
        if job_url:
            log(C_BLUE(f"  ↪ No apply contact — using source job URL as application link: {job_url}"))
            apply_url = job_url
            fallback_to_source = True
        else:
            log(C_RED(f"  ⚑ FLAGGED (no apply contact and no job URL): {title}"))
            flag_job(job_id, job_url, title, company, source_label)
            mark_scraped(job_id, job_url, title, company, source_label)
            _upsert_row(job_id, {"Status": "flagged|no_apply_contact"})
            processed_ids.add(job_id)
            processed_urls.add(job_url)
            return None, "flagged"

    mark_scraped(job_id, job_url, title, company, source_label)
    processed_ids.add(job_id)
    processed_urls.add(job_url)

    # Structure the description into headings / paragraphs / bullets BEFORE paraphrasing
    description = format_paragraphs(raw_job.get("description", ""))
    blurb       = raw_job.get("company_details", "")
    para_title, para_desc, para_blurb = title, description, blurb

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
    company_website = raw_job.get("company_website", "")
    if not company_website and apply_url and not fallback_to_source:
        try:
            parts = urlsplit(apply_url)
            if parts.scheme and parts.netloc and not any(
                    h in parts.netloc.lower() for h in ("jobsearchmalawi", "careersmw", "careeradmw")):
                company_website = f"{parts.scheme}://{parts.netloc}"
        except Exception:
            pass

    if fallback_to_source:
        apply_method = "source_job_url"
    else:
        apply_method = raw_job.get("apply_method") or (
            "resolved_redirect" if apply_url else ("description_email" if apply_email else "not_found"))

    return {
        "source":         source_label,
        "jobTitle":       para_title,
        "jobDescription": para_desc,
        "companyDetails": para_blurb,
        "originalTitle":  title,
        "originalDesc":   description,
        "jobType":        raw_job.get("job_type", ""),
        "jobLocation":    raw_job.get("location", ""),
        "datePosted":     raw_job.get("posted_date", ""),
        "deadline":       normalize_date(raw_job.get("deadline", "")) or raw_job.get("deadline", ""),
        "application":    application,
        "companyName":    company,
        "companyLogo":    raw_job.get("company_logo", ""),
        "companyWebsite": company_website,
        "categories":     raw_job.get("categories", []),
        "jobUrl":         job_url,
        "_jobId":         job_id,
        "_apply_method":  apply_method,
        "_apply_raw":     raw_job.get("apply_raw", ""),
    }, "ok"

# =============================================================================
#  VERBOSE PRINTER
# =============================================================================

def print_job_verbose(index, job):
    desc = job.get("jobDescription", "")
    desc_preview = (desc[:400] + " [...]") if len(desc) > 400 else desc
    print()
    print(C_DIVIDER())
    print(C_HEADER(f"  JOB #{index}   [{job.get('source','')}]"))
    print(C_DIVIDER())
    print(f"  {C_LABEL('Title (original)')}    : {C_VALUE(job.get('originalTitle',''))}")
    print(f"  {C_LABEL('Title (final)')}       : {C_GREEN(job.get('jobTitle',''))}")
    print(f"  {C_LABEL('Job Type')}             : {job.get('jobType','') or C_DIM('—')}")
    print(f"  {C_LABEL('Location')}             : {job.get('jobLocation','') or C_DIM('—')}")
    print(f"  {C_LABEL('Posted')}               : {job.get('datePosted','') or C_DIM('—')}")
    print(f"  {C_LABEL('Deadline')}             : {job.get('deadline','') or C_DIM('—')}")
    print(f"  {C_LABEL('Categories')}           : {', '.join(job.get('categories', [])) or C_DIM('—')}")
    application = job.get("application", "")
    print(f"  {C_LABEL('Apply')}                : {C_GREEN(application) if application else C_DIM('— not found —')}")
    print(f"  {C_LABEL('Apply Method')}         : {C_DIM(job.get('_apply_method',''))}")
    print()
    print(f"  {C_BLUE('── COMPANY ──────────────────────────────────────────')}")
    print(f"  {C_LABEL('Name')}      : {C_VALUE(job.get('companyName','') or C_DIM('—'))}")
    print(f"  {C_LABEL('Website')}   : {job.get('companyWebsite','') or C_DIM('—')}")
    print(f"  {C_LABEL('Logo')}      : {job.get('companyLogo','') or C_DIM('— none —')}")
    about = job.get("companyDetails", "")
    if about:
        print(f"  {C_LABEL('About')}     : {(about[:200] + ' [...]') if len(about) > 200 else about}")
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
    h = {"Authorization": f"Basic {token}", "Content-Type": "application/json",
         "User-Agent": HEADERS["User-Agent"]}
    if INTERNAL_BOT_KEY:
        h["X-Internal-Auth"] = INTERNAL_BOT_KEY
    return h

def _wp_get(path, params=None):
    return requests.get(path, params=params, headers=_wp_auth_headers(),
                        auth=(WP_USER, WP_PASSWORD), timeout=15, verify=False)

def _wp_post(path, json_data):
    return requests.post(path, json=json_data, headers=_wp_auth_headers(),
                         auth=(WP_USER, WP_PASSWORD), timeout=20, verify=False)

def _wp_json(r, context=""):
    if not r.content:
        log_.error(f"WP API empty response [{r.status_code}] {context} — check credentials/endpoint")
        return None
    try:
        return r.json()
    except Exception:
        log_.error(f"WP API non-JSON response [{r.status_code}] {context}: {r.text[:200].replace(chr(10), ' ')}")
        return None

def probe_wp_api():
    host = urlparse(WP_BASE).netloc
    parts = host.split(".")
    root_host = ".".join(parts[1:]) if len(parts) > 2 else host
    scheme = urlparse(WP_BASE).scheme or "https"
    candidates = list(dict.fromkeys(filter(None, [
        WP_API_BASE,
        f"{scheme}://{root_host}/wp-json/wp/v2",
        f"{WP_BASE.rstrip('/')}/?rest_route=/wp/v2",
    ])))
    for base in candidates:
        try:
            r = requests.get(f"{base}/posts", params={"per_page": 1}, headers=_wp_auth_headers(),
                             auth=(WP_USER, WP_PASSWORD), timeout=10, verify=False)
            if r.status_code in (200, 401):
                log_.info(f"✅ WP REST API reachable: {base} [{r.status_code}]")
                return base
        except Exception as e:
            log_.debug(f"WP probe error {base}: {e}")
    log_.error("❌ WP REST API unreachable. Check WP_BASE_URL, pretty permalinks, and /wp-json/ blocking.")
    return ""

def get_or_create_term(taxonomy_url, name):
    if not name or not name.strip():
        return None
    slug = re.sub(r"-{2,}", "-", re.sub(r"[^a-z0-9-]", "-", name.lower().strip())).strip("-")
    if not slug:
        return None
    try:
        data = _wp_json(_wp_get(taxonomy_url, params={"slug": slug}), f"GET {taxonomy_url}?slug={slug}")
        if isinstance(data, list) and data:
            return data[0]["id"]
    except Exception as e:
        log_.debug(f"Term lookup failed '{name}': {e}")
    try:
        data = _wp_json(_wp_post(taxonomy_url, {"name": name, "slug": slug}), f"POST {taxonomy_url} name={name}")
        if data and isinstance(data, dict):
            return data.get("id")
    except Exception as e:
        log_.error(f"Term create error '{name}': {e}")
    return None

def build_jsonld(job):
    apply_url, apply_email = job.get("application", ""), ""
    if re.match(r"^[A-Za-z0-9.+_-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$", apply_url):
        apply_email, apply_url = apply_url, ""
    schema = {
        "@context": "https://schema.org", "@type": "JobPosting",
        "title": sanitize_text(job.get("jobTitle", "")),
        "description": sanitize_text(job.get("jobDescription", "")),
        "datePosted": sanitize_text(job.get("datePosted", "")),
        "hiringOrganization": {"@type": "Organization", "name": sanitize_text(job.get("companyName", ""))},
        "jobLocation": {"@type": "Place", "address": {
            "@type": "PostalAddress", "addressLocality": sanitize_text(job.get("jobLocation", ""))}},
    }
    if job.get("deadline"):
        schema["validThrough"] = sanitize_text(job["deadline"])
    if job.get("jobType"):
        schema["employmentType"] = sanitize_text(job["jobType"]).upper().replace(" ", "-")
    if apply_url:
        schema["url"] = apply_url
    if apply_email:
        schema["applicationContact"] = {"@type": "ContactPoint", "email": apply_email}
    if job.get("companyWebsite"):
        schema["hiringOrganization"]["sameAs"] = sanitize_text(job["companyWebsite"], is_url=True)
    return f'\n\n<script type="application/ld+json">\n{json.dumps(schema, indent=2, ensure_ascii=False)}\n</script>\n'

def post_job_to_wordpress(job):
    if not WP_USER or not WP_PASSWORD or not WP_BASE:
        return None, None
    h = _wp_auth_headers()
    title       = sanitize_text(job.get("jobTitle", ""))
    description = sanitize_text(job.get("jobDescription", ""))
    if not title or not description:
        log_.warning("Skipping WP post — missing title or description")
        return None, None

    slug = re.sub(r"-{2,}", "-", re.sub(r"[^a-z0-9-]", "-", title.lower())).strip("-")[:80]
    try:
        posts = _wp_json(_wp_get(WP_JOBS_URL, params={"slug": slug, "status": "publish"}), f"duplicate check slug={slug}")
        if isinstance(posts, list) and posts:
            log_.info(f"⏭ Already on WP: {title}")
            return posts[0]["id"], posts[0].get("link")
    except Exception:
        pass

    logo_url    = sanitize_text(job.get("companyLogo", ""), is_url=True)
    location    = sanitize_text(job.get("jobLocation", ""))
    raw_type    = sanitize_text(job.get("jobType", "")) or "Full-time"
    job_type_s  = JOB_TYPE_MAPPING.get(raw_type.lower().strip(), "full-time")
    company     = sanitize_text(job.get("companyName", ""))
    application = sanitize_text(job.get("application", ""), is_url=True)
    deadline    = sanitize_text(job.get("deadline", ""))
    co_website  = sanitize_text(job.get("companyWebsite", ""), is_url=True)
    about       = sanitize_text(job.get("companyDetails", ""))
    date_posted = sanitize_text(job.get("datePosted", ""))
    source      = job.get("source", "")

    is_email_app = bool(re.match(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", application))
    is_url_app   = bool(re.match(r"^https?://[^\s]+$", application))
    if not (is_email_app or is_url_app):
        application = ""
    if job.get("_apply_method") != "source_job_url":
        application = blank_if_linkedin(application)
    co_website  = blank_if_linkedin(co_website)

    attachment_id = None
    if logo_url and logo_url.startswith("http"):
        try:
            img_r = SESSION.get(logo_url, timeout=15)
            if img_r.status_code == 200 and img_r.content:
                ct = img_r.headers.get("Content-Type", "image/jpeg").split(";")[0].strip()
                ext = "png" if "png" in ct else "gif" if "gif" in ct else "webp" if "webp" in ct else "jpg"
                fn = re.sub(r"-{2,}", "-", re.sub(r"[^a-z0-9]", "-", (company or "company").lower())).strip("-") + "-logo." + ext
                up_h = dict(_wp_auth_headers())
                up_h["Content-Disposition"] = f'attachment; filename="{fn}"'
                up_h["Content-Type"] = ct
                up_r = requests.post(WP_MEDIA_URL, headers=up_h, data=img_r.content,
                                     auth=(WP_USER, WP_PASSWORD), timeout=20, verify=False)
                if up_r.status_code in (200, 201):
                    up_data = _wp_json(up_r, "media upload")
                    if up_data:
                        attachment_id = up_data.get("id")
                else:
                    log_.warning(f"Logo upload HTTP {up_r.status_code} for {logo_url}")
            else:
                log_.warning(f"Logo download HTTP {img_r.status_code}: {logo_url}")
        except Exception as e:
            log_.warning(f"Logo upload failed: {e}")

    cat_term_id = get_or_create_term(f"{WP_API_BASE}/categories", location) if location else None
    tag_names = list(filter(None, [
        job_type_s.replace("-", " ").title(), company or None, source or None,
    ] + list(job.get("categories", []))[:3]))
    tag_ids = [t for t in (get_or_create_term(f"{WP_API_BASE}/tags", n) for n in dict.fromkeys(tag_names)) if t]

    expiry_comment = f"<!-- job-expiry: {deadline} -->" if deadline else ""
    # Real HTML (<h3>/<p>/<ul><li>) so WordPress shows proper paragraphs; JSON-LD keeps plain text
    content = text_to_html(description) + "\n\n" + expiry_comment + build_jsonld(job)

    payload = {
        "title": title, "content": content, "slug": slug, "status": "publish",
        "featured_media": attachment_id or 0,
        "meta": {
            "_job_location": location, "_job_type": job_type_s, "_application": application,
            "_job_expires": deadline, "_company_name": company, "_company_website": co_website,
            "_company_logo": str(attachment_id) if attachment_id else "",
            "_company_details": about, "_date_posted": date_posted,
        },
    }
    if cat_term_id:
        payload["categories"] = [cat_term_id]
    if tag_ids:
        payload["tags"] = tag_ids

    for attempt in range(3):
        try:
            r = requests.post(WP_JOBS_URL, json=payload, headers=h, auth=(WP_USER, WP_PASSWORD),
                              timeout=20, verify=False)
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
    "Source", "Job Title", "Job Type", "Job Location", "Date Posted", "Deadline",
    "Job Description", "Application", "Apply Method", "Company Name",
    "Company Logo", "Company Website", "Company Details", "Categories", "Job URL",
]

def _save_excel(jobs):
    if not _XLSX_AVAILABLE:
        log_.warning("openpyxl not installed — skipping Excel export")
        return
    if not jobs:
        return
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.append(EXCEL_HEADERS)
        for j in jobs:
            ws.append([
                j.get("source", ""), j.get("jobTitle", ""), j.get("jobType", ""),
                j.get("jobLocation", ""), j.get("datePosted", ""), j.get("deadline", ""),
                j.get("jobDescription", ""), j.get("application", ""), j.get("_apply_method", ""),
                j.get("companyName", ""), j.get("companyLogo", ""), j.get("companyWebsite", ""),
                j.get("companyDetails", ""), ", ".join(j.get("categories", [])), j.get("jobUrl", ""),
            ])
        wb.save(OUTPUT_FILE)
        log_.info(f"Saved {len(jobs)} rows → {OUTPUT_FILE}")
    except Exception as e:
        log_.error(f"Excel save failed: {e}")

# =============================================================================
#  RUN ONE SOURCE
# =============================================================================

def run_source(name, src, state, stats):
    label = src["label"]
    st = stats[name]
    log()
    log(C_HEADER("=" * 80))
    log(C_HEADER(f"  SOURCE: {label}  ({src['base']})"))
    log(C_HEADER("=" * 80))

    if not source_reachable(src["base"]):
        st["status"] = "SKIPPED (unreachable / blocked)"
        log(C_RED(f"  ⏭  {label} skipped — site unreachable or returned an error"))
        return

    try:
        urls = src["collect"]()
    except Exception as e:
        st["status"] = f"SKIPPED (collect error: {type(e).__name__})"
        log(C_RED(f"  ⏭  {label} skipped — URL collection failed: {e}"))
        return

    st["urls"] = len(urls)
    if not urls:
        st["status"] = "SKIPPED (no jobs found)"
        log(C_RED(f"  ⏭  {label} skipped — no job URLs found"))
        return

    consecutive_errors = 0
    for i, url in enumerate(urls, start=1):
        if MAX_JOBS and len(state["jobs"]) >= MAX_JOBS:
            log(f"\nMAX_JOBS limit ({MAX_JOBS}) reached.")
            st["status"] = "stopped (MAX_JOBS)"
            return
        if MAX_JOBS_PER_SOURCE and st["new"] >= MAX_JOBS_PER_SOURCE:
            log(f"\nMAX_JOBS_PER_SOURCE limit ({MAX_JOBS_PER_SOURCE}) reached for {label}.")
            st["status"] = "stopped (MAX_JOBS_PER_SOURCE)"
            return
        if consecutive_errors >= MAX_CONSEC_ERRORS:
            st["status"] = f"ABORTED ({consecutive_errors} consecutive errors)"
            log(C_RED(f"  ⏭  {label} aborted — {consecutive_errors} consecutive errors (blocked?)"))
            return

        if make_job_id(url) in state["ids"] or url in state["urls"]:
            st["skipped_known"] += 1
            log(C_DIM(f"[{label} {i}/{len(urls)}] already processed — {url}"))
            continue

        log(f"\n[{label} {i}/{len(urls)}] Scraping: {url}")
        try:
            raw_job = src["parse"](url)
        except Exception as e:
            st["errors"] += 1
            consecutive_errors += 1
            log(C_RED(f"  ERROR scraping {url}: {type(e).__name__}: {e}"))
            time.sleep(src["delay"])
            continue

        if not raw_job or not raw_job.get("title"):
            st["errors"] += 1
            consecutive_errors += 1
            log(C_RED("  ✗ nothing parsed (blocked / removed / empty page) — skipped"))
            time.sleep(src["delay"])
            continue
        consecutive_errors = 0

        try:
            job, status = process_job(raw_job, label, state["ids"], state["urls"], state["seen"])
        except Exception as e:
            st["errors"] += 1
            log(C_RED(f"  ERROR processing job: {type(e).__name__}: {e}"))
            time.sleep(src["delay"])
            continue

        if status == "flagged":
            st["flagged"] += 1
        elif status == "dup":
            st["dup"] += 1
        if job is None:
            time.sleep(src["delay"])
            continue

        st["new"] += 1
        state["jobs"].append(job)
        print_job_verbose(len(state["jobs"]), job)

        if WP_USER and WP_PASSWORD and state["wp_ok"]:
            print(C_BLUE("\n  📤 Posting to WordPress …"))
            try:
                wp_id, wp_url = post_job_to_wordpress(job)
            except Exception as e:
                wp_id, wp_url = None, None
                log(C_RED(f"  WP exception: {e}"))
            if wp_id:
                mark_posted(job["_jobId"], wp_id)
                st["posted"] += 1
                print(C_GREEN(f"  ✅ WP ID={wp_id}  🔗 {wp_url}"))
            else:
                mark_failed(job["_jobId"], "wp_post_failed_or_skipped")
                st["wp_failed"] += 1
                print(C_RED("  ❌ WordPress post failed / skipped"))
        else:
            mark_failed(job["_jobId"], "wp_post_failed_or_skipped")
            print(C_DIM("  WordPress posting disabled / unreachable"))

        if len(state["jobs"]) % 25 == 0:
            _save_excel(state["jobs"])
        time.sleep(src["delay"])

    st["status"] = st.get("status") or "done"

# =============================================================================
#  MAIN
# =============================================================================

def main():
    global WP_API_BASE, WP_JOBS_URL, WP_MEDIA_URL, MAX_JOBS, MAX_JOBS_PER_SOURCE

    ap = argparse.ArgumentParser(description="Multi-source Malawi jobs scraper")
    ap.add_argument("--sources", default=os.environ.get("SOURCES", DEFAULT_SOURCES),
                    help=f"comma separated: {','.join(SOURCES)}")
    ap.add_argument("--max-per-source", type=int, default=None)
    args, _unknown = ap.parse_known_args()  # ignores Colab/Jupyter's "-f kernel.json"
    if args.max_per_source is not None:
        MAX_JOBS_PER_SOURCE = args.max_per_source

    selected = [s.strip().lower() for s in args.sources.split(",") if s.strip()]
    start_time = datetime.now()

    print()
    print(C_HEADER("=" * 80))
    print(C_HEADER("  MALAWI MULTI-SOURCE JOB SCRAPER + WORDPRESS POSTING"))
    print(C_HEADER("=" * 80))
    print(f"  Sources         : {', '.join(selected)}")
    print(f"  Request delay   : {REQUEST_DELAY}s")
    print(f"  Max new jobs    : {'unlimited' if not MAX_JOBS else MAX_JOBS}  (per source: {'unlimited' if not MAX_JOBS_PER_SOURCE else MAX_JOBS_PER_SOURCE})")
    print(f"  Resolve apply   : {'✅ enabled' if RESOLVE_APPLY_URLS else '❌ disabled'}")
    print(f"  Paraphrase      : {'✅ enabled' if (ENABLE_PARAPHRASE and MISTRAL_API_KEY) else '❌ disabled'}")
    print(f"  WordPress post  : {'✅ enabled' if (WP_USER and WP_PASSWORD and WP_BASE) else '❌ disabled'}")
    print(f"  Excel export    : {'✅ enabled' if _XLSX_AVAILABLE else '❌ disabled (pip install openpyxl)'}")
    print(f"  LinkedIn        : location='{LI_LOCATION}' range={LI_TIME_RANGE} keywords={LI_KEYWORDS}")
    print(f"  HTML parser     : {_PARSER}")
    print(f"  Started         : {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(C_HEADER("=" * 80))

    wp_ok = False
    if WP_USER and WP_PASSWORD and WP_BASE:
        discovered = probe_wp_api()
        if discovered:
            wp_ok = True
            if discovered != WP_API_BASE:
                log(C_BLUE(f"  ℹ️  WP API base corrected to: {discovered}"))
                WP_API_BASE = discovered
                WP_JOBS_URL = f"{WP_API_BASE}/job-listings"
                WP_MEDIA_URL = f"{WP_API_BASE}/media"
        else:
            log(C_RED("  ⚠️  WP posting will be skipped — REST API unreachable"))

    _init_tracker()
    processed_ids, processed_urls = load_processed_ids()
    print(f"  Tracker loaded: {len(processed_ids)} previously processed job IDs\n")

    state = {"jobs": [], "ids": processed_ids, "urls": processed_urls, "seen": set(), "wp_ok": wp_ok}
    stats = {}

    for name in selected:
        src = SOURCES.get(name)
        if not src:
            log(C_RED(f"  Unknown source '{name}' — valid: {', '.join(SOURCES)}"))
            continue
        stats[name] = Counter()
        stats[name]["status"] = ""
        try:
            run_source(name, src, state, stats)
        except KeyboardInterrupt:
            raise
        except Exception as e:
            stats[name]["status"] = f"FAILED ({type(e).__name__}: {e})"
            log(C_RED(f"  ⏭  {src['label']} failed unexpectedly and was skipped: {e}"))
        _save_excel(state["jobs"])
        if MAX_JOBS and len(state["jobs"]) >= MAX_JOBS:
            break

    _save_excel(state["jobs"])

    duration = (datetime.now() - start_time).total_seconds() / 60.0
    print()
    print(C_HEADER("=" * 80))
    print(C_HEADER("  SCRAPE COMPLETE"))
    print(C_HEADER("=" * 80))
    for name, st in stats.items():
        label = SOURCES[name]["label"]
        print(f"  {C_LABEL(label):<28} urls={st['urls']:<5} new={C_GREEN(str(st['new']))} "
              f"posted={C_GREEN(str(st['posted']))} flagged={st['flagged']} dup={st['dup']} "
              f"known={st['skipped_known']} errors={C_RED(str(st['errors'])) if st['errors'] else 0} "
              f"wp_failed={st['wp_failed']}  → {st['status'] or 'done'}")
    total_posted = sum(s["posted"] for s in stats.values())
    total_flagged = sum(s["flagged"] for s in stats.values())
    print()
    print(f"  {C_LABEL('New jobs processed')}    : {C_GREEN(str(len(state['jobs'])))}")
    print(f"  {C_LABEL('Posted to WordPress')}   : {C_GREEN(str(total_posted))}")
    print(f"  {C_LABEL('Flagged (no apply)')}    : {C_RED(str(total_flagged)) if total_flagged else '0'}")
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

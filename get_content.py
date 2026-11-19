#!/usr/bin/env python3
"""
Fetch records retention schedules, detect JS-rendered pages, extract text.

Install:  pip install requests beautifulsoup4 pdfplumber

Output (created next to this script):
  snapshots/YYYY-MM-DD/<slug>.<html|pdf|txt>   raw files exactly as received
  extracted/YYYY-MM-DD/<slug>.txt              plain text for the LLM step
  manifest_YYYY-MM-DD.csv                      one row per URL: status, flags, hashes
"""

import csv
import hashlib
import re
import time
from datetime import date, datetime, timezone
from pathlib import Path
from urllib.parse import urldefrag, urlparse

import pdfplumber
import requests
from bs4 import BeautifulSoup

# ---------------------------------------------------------------- config

USER_AGENT = "RecordsRetentionResearch/0.1 (contact: you@example.edu)"  # <- edit
DELAY_SECONDS = 2.0          # minimum gap between requests to the same host
TIMEOUT = 60

# HTML pages with less visible text than this are suspicious (likely JS shell)
MIN_TEXT_CHARS = 800
# PDFs averaging fewer characters per page than this are probably scanned
MIN_PDF_CHARS_PER_PAGE = 50

SOURCES = [
    {"name": "UC Office of the President",
     "url": "https://recordsretention.ucop.edu/"},
    {"name": "Google Doc (add institution name)",
     "url": "https://docs.google.com/document/d/1vn5rRoBJI8ieUZYhsy04w5V5CzU2xO35wh8YshIkxX8/edit?tab=t.0#heading=h.s5zmxpjhy2if"},
    {"name": "Boston University",
     "url": "https://www.bu.edu/policies/record-retention-table"},
    {"name": "Penn State",
     "url": "https://policy.psu.edu/general-retention-schedule#E"},
    {"name": "Massachusetts Statewide Records Schedule",
     "url": "https://www.sec.state.ma.us/divisions/public-records/download/MA_Statewide_Records_Schedule.pdf"},
    {"name": "Suffolk University",
     "url": "https://www.suffolk.edu/-/media/suffolk/documents/academics/libraries/moakley-archive/records-management/su_retention-schedule_2023-v2.pdf?la=en&hash=DE60ECF1ECAF723848FDA7D98332BE41DEAAC94E"},
]

TODAY = date.today().isoformat()
BASE = Path(__file__).parent
SNAP_DIR = BASE / "snapshots" / TODAY
TEXT_DIR = BASE / "extracted" / TODAY

session = requests.Session()
session.headers.update({"User-Agent": USER_AGENT})

# ---------------------------------------------------------------- helpers


def slugify(name: str, url: str) -> str:
    base = re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")[:40]
    return f"{base}-{hashlib.sha1(url.encode()).hexdigest()[:8]}"


def normalize_url(url: str):
    """Return (url_to_fetch, note). Strips #fragments (servers never see them)
    and rewrites Google Docs edit links to a plain-text export link."""
    url, _ = urldefrag(url)
    m = re.match(r"https://docs\.google\.com/document/d/([\w-]+)", url)
    if m:
        export = f"https://docs.google.com/document/d/{m.group(1)}/export?format=txt"
        return export, "Google Doc rewritten to text export (works only if link-sharing is on)"
    return url, ""


_robots_cache = {}


def robots_allows(url: str) -> bool:
    """Check robots.txt once per host. Missing robots.txt means allowed."""
    parts = urlparse(url)
    host = f"{parts.scheme}://{parts.netloc}"
    if host not in _robots_cache:
        from urllib.robotparser import RobotFileParser
        rp = RobotFileParser()
        try:
            r = session.get(f"{host}/robots.txt", timeout=15)
            if r.status_code == 200:
                rp.parse(r.text.splitlines())
            else:
                rp.parse([])  # no usable robots.txt -> allow all
        except requests.RequestException:
            rp.parse([])
        _robots_cache[host] = rp
    return _robots_cache[host].can_fetch(USER_AGENT, url)


_last_hit = {}


def polite_wait(url: str):
    host = urlparse(url).netloc
    wait = DELAY_SECONDS - (time.time() - _last_hit.get(host, 0))
    if wait > 0:
        time.sleep(wait)
    _last_hit[host] = time.time()


def detect_kind(resp: requests.Response) -> str:
    ctype = resp.headers.get("Content-Type", "").lower()
    if "pdf" in ctype or resp.content[:5] == b"%PDF-":
        return "pdf"
    if "text/plain" in ctype:
        return "txt"
    return "html"


def clean_text(text: str) -> str:
    lines = [re.sub(r"[ \t\u00a0]+", " ", ln).strip() for ln in text.splitlines()]
    text = "\n".join(lines)
    return re.sub(r"\n{3,}", "\n\n", text).strip()

# ---------------------------------------------------------------- extraction


def extract_html(html: str):
    """Return (text, info). info holds the signals used for the JS check."""
    soup = BeautifulSoup(html, "html.parser")

    # Signals gathered BEFORE we strip scripts
    n_scripts = len(soup.find_all("script"))
    noscript_text = " ".join(t.get_text(" ") for t in soup.find_all("noscript")).lower()
    spa_root = soup.find(id=re.compile(r"^(root|app|__next|__nuxt)$"))

    for tag in soup(["script", "style", "noscript", "svg"]):
        tag.decompose()

    # Render tables as "cell | cell | cell" rows so structure survives
    tables = soup.find_all("table")
    for table in tables:
        rows = []
        for tr in table.find_all("tr"):
            cells = [c.get_text(" ", strip=True) for c in tr.find_all(["th", "td"])]
            if any(cells):
                rows.append(" | ".join(cells))
        table.replace_with(soup.new_string("\n" + "\n".join(rows) + "\n"))

    text = clean_text(soup.get_text("\n"))
    info = {"n_scripts": n_scripts, "n_tables": len(tables),
            "noscript_text": noscript_text, "spa_root": spa_root is not None}
    return text, info


def js_check(text: str, info: dict) -> list:
    """Return a list of reasons the page may need a real browser (empty = fine)."""
    reasons = []
    low = text.lower()
    if len(text) < MIN_TEXT_CHARS:
        reasons.append(f"little visible text ({len(text)} chars)")
    if "enable javascript" in info["noscript_text"] or "enable javascript" in low \
            or "requires javascript" in low:
        reasons.append("page asks for JavaScript")
    if info["spa_root"] and len(text) < 2000:
        reasons.append("single-page-app root element with thin content")
    if info["n_scripts"] >= 10 and len(text) < 1500:
        reasons.append(f"{info['n_scripts']} scripts but little text")
    return reasons


def extract_pdf(path: Path):
    """Return (text, n_pages, looks_scanned)."""
    parts = []
    with pdfplumber.open(path) as pdf:
        n_pages = len(pdf.pages)
        for i, page in enumerate(pdf.pages, start=1):
            parts.append(f"--- page {i} ---\n{page.extract_text() or ''}")
    text = clean_text("\n".join(parts))
    visible = len(re.sub(r"--- page \d+ ---", "", text).strip())
    looks_scanned = n_pages > 0 and (visible / n_pages) < MIN_PDF_CHARS_PER_PAGE
    return text, n_pages, looks_scanned

# ---------------------------------------------------------------- main fetch


def process(source: dict) -> dict:
    name, original_url = source["name"], source["url"]
    fetch_url, note = normalize_url(original_url)
    slug = slugify(name, original_url)
    row = {
        "name": name, "original_url": original_url, "fetched_url": fetch_url,
        "retrieved_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "status": "", "kind": "", "bytes": "", "sha256": "", "text_chars": "",
        "pages": "", "needs_browser": "", "needs_ocr": "", "notes": note, "error": "",
        "snapshot_path": "", "text_path": "",
    }

    def add_note(msg):
        row["notes"] = f"{row['notes']}; {msg}" if row["notes"] else msg

    try:
        if not robots_allows(fetch_url):
            row["error"] = "blocked by robots.txt (not fetched)"
            return row

        polite_wait(fetch_url)
        resp = session.get(fetch_url, timeout=TIMEOUT)
        row["status"] = resp.status_code
        resp.raise_for_status()

        if "accounts.google.com" in resp.url:
            row["error"] = "redirected to Google sign-in: document is not public"
            return row

        kind = detect_kind(resp)
        row.update(kind=kind, bytes=len(resp.content),
                   sha256=hashlib.sha256(resp.content).hexdigest())

        SNAP_DIR.mkdir(parents=True, exist_ok=True)
        TEXT_DIR.mkdir(parents=True, exist_ok=True)
        snap_path = SNAP_DIR / f"{slug}.{kind}"
        snap_path.write_bytes(resp.content)
        row["snapshot_path"] = str(snap_path)

        if kind == "pdf":
            text, n_pages, scanned = extract_pdf(snap_path)
            row.update(pages=n_pages, needs_ocr=scanned)
            if scanned:
                add_note("PDF has almost no text layer; likely scanned, needs OCR")
        elif kind == "txt":
            text = clean_text(resp.text)
        else:
            text, info = extract_html(resp.text)
            reasons = js_check(text, info)
            row["needs_browser"] = bool(reasons)
            if reasons:
                add_note("JS suspected: " + "; ".join(reasons))
            if info["n_tables"]:
                add_note(f"{info['n_tables']} HTML table(s) extracted")

        row["text_chars"] = len(text)
        text_path = TEXT_DIR / f"{slug}.txt"
        text_path.write_text(text, encoding="utf-8")
        row["text_path"] = str(text_path)

    except requests.RequestException as e:
        row["error"] = f"{type(e).__name__}: {e}"
    except Exception as e:  # keep the loop going if one parse fails
        row["error"] = f"parse error, {type(e).__name__}: {e}"
    return row


def main():
    rows = []
    for i, source in enumerate(SOURCES, start=1):
        print(f"[{i}/{len(SOURCES)}] {source['name']}")
        row = process(source)
        rows.append(row)
        if row["error"]:
            print(f"    ERROR: {row['error']}")
        else:
            flags = []
            if row["needs_browser"] is True:
                flags.append("NEEDS BROWSER")
            if row["needs_ocr"] is True:
                flags.append("NEEDS OCR")
            print(f"    {row['kind']}, {row['text_chars']} chars "
                  f"{'(' + ', '.join(flags) + ')' if flags else ''}")

    manifest = BASE / f"manifest_{TODAY}.csv"
    with manifest.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)
    print(f"\nManifest written to {manifest}")


if __name__ == "__main__":
    main()
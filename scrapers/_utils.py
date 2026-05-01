# scrapers/_utils.py
# Shared helpers for all BCN Transit Tracker scrapers
# Bay City News

import csv
import logging
import os
from datetime import datetime, timezone
from pathlib import Path

import requests

# ── Paths ─────────────────────────────────────────────────────────────────────
ROOT        = Path(__file__).resolve().parent.parent
RAW_DIR     = ROOT / "data" / "raw"
SCRAPE_LOG  = RAW_DIR / "_scrape_log.csv"

RAW_DIR.mkdir(parents=True, exist_ok=True)

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level   = logging.INFO,
    format  = "%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt = "%Y-%m-%d %H:%M:%S"
)
log = logging.getLogger("bcn_transit")


# ── Timestamp helpers ─────────────────────────────────────────────────────────
def now_utc() -> str:
    """ISO 8601 timestamp in UTC — used for scraped_at fields."""
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def today_str() -> str:
    """YYYY-MM-DD — used for raw file naming."""
    return datetime.now(timezone.utc).strftime("%Y-%m-%d")


# ── HTTP fetch with retry ─────────────────────────────────────────────────────
def fetch_url(url: str, retries: int = 3, timeout: int = 30) -> requests.Response:
    """
    GET a URL with simple retry logic. Raises on final failure.
    Usage:
        resp = fetch_url("https://example.com/data.csv")
        text = resp.text
    """
    headers = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
    "Referer": "https://www.transit.dot.gov/",
}
    for attempt in range(1, retries + 1):
        try:
            resp = requests.get(url, headers=headers, timeout=timeout)
            resp.raise_for_status()
            log.info(f"Fetched {url} [{resp.status_code}]")
            return resp
        except requests.RequestException as e:
            log.warning(f"Attempt {attempt}/{retries} failed for {url}: {e}")
            if attempt == retries:
                raise
    raise RuntimeError(f"All {retries} attempts failed for {url}")


# ── Raw file path ─────────────────────────────────────────────────────────────
def raw_path(agency_id: str) -> Path:
    """
    Returns the path for today's raw scrape file.
    Example: data/raw/bart_2026-04-30.csv
    """
    return RAW_DIR / f"{agency_id}_{today_str()}.csv"


# ── Scrape log ────────────────────────────────────────────────────────────────
def log_scrape(agency_id: str, status: str, rows: int, notes: str = "") -> None:
    """
    Appends one row to data/raw/_scrape_log.csv.
    Called at the end of every scraper run — success or failure.
    """
    file_exists = SCRAPE_LOG.exists()
    with open(SCRAPE_LOG, "a", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=[
            "scraped_at", "agency_id", "status", "rows_written", "notes"
        ])
        if not file_exists:
            writer.writeheader()
        writer.writerow({
            "scraped_at"   : now_utc(),
            "agency_id"    : agency_id,
            "status"       : status,
            "rows_written" : rows,
            "notes"        : notes,
        })
    log.info(f"Logged scrape: {agency_id} | {status} | {rows} rows")


# ── Normalized column schema ──────────────────────────────────────────────────
# Every scraper must return a DataFrame with exactly these columns.
# This is what build_master.py expects.
NORMALIZED_COLUMNS = [
    "agency_id",      # e.g. "bart"
    "agency_name",    # e.g. "Bay Area Rapid Transit"
    "ntd_id",         # NTD agency ID, e.g. "90003"
    "date",           # First day of the month: YYYY-MM-01
    "metric",         # e.g. "monthly_exits", "upt", "vrm"
    "value",          # Numeric
    "mode",           # e.g. "HR", "MB", "ALL"
    "unit",           # e.g. "exits", "trips", "miles"
    "source",         # e.g. "bart.gov", "NTD", "SFMTA_APC"
    "source_url",     # Direct URL where data was pulled from
    "scraped_at",     # ISO 8601 UTC timestamp of this scrape
    "is_provisional", # True if data may be revised (e.g. ahead of NTD release)
]

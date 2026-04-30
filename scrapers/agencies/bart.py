# scrapers/agencies/bart.py
# BART daily exits scraper
# Source: bart.gov ridership watch page (paid exits, updated daily M–F)
# URL: https://www.bart.gov/news/articles/2025/news20250109-1
#
# What this pulls:
#   - Daily paid exit counts for all completed months (aggregated to monthly)
#   - The current incomplete month is stored as provisional
#
# Output schema: see scrapers/_utils.py NORMALIZED_COLUMNS
# Output file:   data/raw/bart_YYYY-MM-DD.csv
#
# Bay City News | Andres Jimenez Larios

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

import io
import re
from datetime import date

import pandas as pd
from bs4 import BeautifulSoup

from scrapers._utils import (
    fetch_url, raw_path, log_scrape, now_utc, NORMALIZED_COLUMNS, log
)

# ── Constants ─────────────────────────────────────────────────────────────────
AGENCY_ID   = "bart"
AGENCY_NAME = "Bay Area Rapid Transit"
NTD_ID      = "90003"
SOURCE_URL  = "https://www.bart.gov/news/articles/2025/news20250109-1"

# Variance threshold vs NTD — BART exits run ~7.5% lower than NTD adjusted UPT
# This is expected and documented; not a flag-worthy discrepancy
EXPECTED_NTD_VARIANCE_NOTE = (
    "BART submits exits as UPT to NTD. NTD applies a ~7.5% upward adjustment. "
    "Variance up to 10% between bart.gov and NTD is expected and normal."
)


# ── 1. Fetch ──────────────────────────────────────────────────────────────────
def fetch() -> str:
    resp = fetch_url(SOURCE_URL)
    return resp.text


# ── 2. Parse ──────────────────────────────────────────────────────────────────
def parse(html: str) -> pd.DataFrame:
    """
    BART's ridership page contains one HTML table per month.
    Each table has two columns: Date (M/D/YY) and Ridership (# of paid exits).
    Tables are identified by their <th> header containing the month/year.

    Returns a dataframe with columns: date_str, exits
    """
    soup   = BeautifulSoup(html, "lxml")
    tables = soup.find_all("table")

    if not tables:
        raise ValueError("No tables found on BART ridership page — page structure may have changed")

    rows = []
    for table in tables:
        # Pull all <tr> rows with two <td> cells
        for tr in table.find_all("tr"):
            tds = tr.find_all("td")
            if len(tds) != 2:
                continue
            date_str = tds[0].get_text(strip=True)
            val_str  = tds[1].get_text(strip=True).replace(",", "").replace("\xa0", "")

            # Skip header rows and blank cells
            if not date_str or not val_str or date_str.lower() == "date":
                continue
            if not val_str.isdigit():
                continue

            rows.append({"date_str": date_str, "exits": int(val_str)})

    if not rows:
        raise ValueError("Parsed 0 rows from BART page — check page structure")

    df = pd.DataFrame(rows)
    log.info(f"Parsed {len(df)} daily exit rows from BART page")
    return df


# ── 3. Normalize — daily → monthly ───────────────────────────────────────────
def normalize(df: pd.DataFrame) -> pd.DataFrame:
    """
    Converts daily exit rows to monthly totals.
    Marks the current (incomplete) month as provisional.
    Returns a dataframe conforming to NORMALIZED_COLUMNS.
    """
    # Parse M/D/YY format (e.g. "3/25/26")
    df["date_parsed"] = pd.to_datetime(df["date_str"], format="%m/%d/%y", errors="coerce")
    df = df.dropna(subset=["date_parsed"])

    df["month_flo"] = df["date_parsed"].dt.to_period("M").dt.to_timestamp()
    df["days_in_month"] = df["date_parsed"].dt.days_in_month

    # Determine which months are complete vs partial
    today = pd.Timestamp(date.today())
    month_stats = (
        df.groupby("month_flo")
        .agg(
            days_reported = ("date_parsed", "count"),
            days_in_month = ("days_in_month", "first"),
            total_exits   = ("exits", "sum"),
            avg_daily     = ("exits", "mean"),
        )
        .reset_index()
    )
    month_stats["is_partial"]  = month_stats["days_reported"] < month_stats["days_in_month"]
    # Project partial months: avg daily * days in month
    month_stats["total_exits_proj"] = month_stats.apply(
        lambda r: round(r["avg_daily"] * r["days_in_month"]) if r["is_partial"] else r["total_exits"],
        axis=1
    )

    # Build normalized output — one row per month
    out_rows = []
    for _, row in month_stats.iterrows():
        out_rows.append({
            "agency_id"      : AGENCY_ID,
            "agency_name"    : AGENCY_NAME,
            "ntd_id"         : NTD_ID,
            "date"           : row["month_flo"].strftime("%Y-%m-%d"),
            "metric"         : "monthly_exits",
            "value"          : int(row["total_exits_proj"]),
            "mode"           : "HR",
            "unit"           : "exits",
            "source"         : "bart.gov",
            "source_url"     : SOURCE_URL,
            "scraped_at"     : now_utc(),
            "is_provisional" : bool(row["is_partial"]),
        })

    result = pd.DataFrame(out_rows)
    log.info(
        f"Normalized to {len(result)} monthly rows "
        f"({result['is_provisional'].sum()} provisional)"
    )
    return result[NORMALIZED_COLUMNS]


# ── 4. Run ────────────────────────────────────────────────────────────────────
def run() -> pd.DataFrame:
    log.info(f"Starting scrape: {AGENCY_NAME}")
    try:
        raw    = fetch()
        parsed = parse(raw)
        df     = normalize(parsed)

        out = raw_path(AGENCY_ID)
        df.to_csv(out, index=False)
        log_scrape(AGENCY_ID, "success", len(df), notes=EXPECTED_NTD_VARIANCE_NOTE)
        log.info(f"Wrote {len(df)} rows → {out}")
        return df

    except Exception as e:
        log_scrape(AGENCY_ID, "error", 0, notes=str(e))
        log.error(f"Scrape failed for {AGENCY_NAME}: {e}")
        raise


if __name__ == "__main__":
    df = run()
    print("\n── Latest 6 months ─────────────────────────────")
    print(df.tail(6).to_string(index=False))

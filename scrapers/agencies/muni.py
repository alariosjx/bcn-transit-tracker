# scrapers/agencies/muni.py
# Muni monthly boardings scraper
# Source: SFMTA Tableau CSV (SystemwideRidershipRecovery)
# URL: https://transtat.sfmta.com/t/public/views/SystemwideRidershipRecovery/MonthlySystemwideRecoveryAccessibleTable.csv
#
# What this pulls:
#   - Monthly total boardings from April 2020 to present (~1 month lag)
#   - 2019 baseline for recovery % calculation
#
# Output schema: see scrapers/_utils.py NORMALIZED_COLUMNS
# Output file:   data/raw/muni_YYYY-MM-DD.csv
#
# Bay City News | Andres Jimenez Larios

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

import io
from datetime import date

import pandas as pd

from scrapers._utils import (
    fetch_url, raw_path, log_scrape, now_utc, NORMALIZED_COLUMNS, log
)

AGENCY_ID   = "muni"
AGENCY_NAME = "San Francisco Municipal Railway"
NTD_ID      = "90015"
SOURCE_URL  = "https://transtat.sfmta.com/t/public/views/SystemwideRidershipRecovery/MonthlySystemwideRecoveryAccessibleTable.csv"




# ── Fetch ─────────────────────────────────────────────────────────────────────
def fetch() -> str:
    resp = fetch_url(SOURCE_URL, timeout=30)
    log.info(f"Received {len(resp.content):,} bytes from SFMTA")
    return resp.text


# ── Parse ─────────────────────────────────────────────────────────────────────
def parse(text: str) -> pd.DataFrame:
    """
    CSV has three columns:
      Measure Names | Month of MONTH | Measure Values

    Measure Names values:
      - "Monthly Total Boardings (accessible copy)"
      - "Baseline Monthly Total Boardings (accessible copy)"
      - "Monthly Recovery"

    We only want "Monthly Total Boardings".
    Month format: "April 2020"
    Values are formatted with commas: "2,555,000"
    """
    df = pd.read_csv(io.StringIO(text))
    log.info(f"Parsed {len(df)} rows, columns: {df.columns.tolist()}")

    # Rename columns for easier access
    df.columns = ["measure", "month_str", "value_str"]

    # Filter to total boardings only
    boardings = df[df["measure"] == "Monthly Total Boardings (accessible copy)"].copy()

    if boardings.empty:
        raise ValueError("No 'Monthly Total Boardings' rows found — check CSV structure")

    # Parse month strings like "April 2020" → datetime
    boardings["date"] = pd.to_datetime(boardings["month_str"], format="%B %Y", errors="coerce")
    boardings = boardings.dropna(subset=["date"])

    # Parse value strings like "2,555,000" → int
    boardings["value"] = boardings["value_str"].str.replace(",", "").str.strip()
    boardings["value"] = pd.to_numeric(boardings["value"], errors="coerce")
    boardings = boardings.dropna(subset=["value"])
    boardings["value"] = boardings["value"].astype(int)

    log.info(f"Parsed {len(boardings)} monthly boarding rows")
    log.info(f"Date range: {boardings['date'].min().strftime('%b %Y')} → {boardings['date'].max().strftime('%b %Y')}")
    return boardings[["date", "value"]].reset_index(drop=True)


# ── Normalize ─────────────────────────────────────────────────────────────────
def normalize(df: pd.DataFrame) -> pd.DataFrame:
    """
    Converts parsed boardings to normalized schema.
    All months from SFMTA CSV are complete — no provisional data.
    """
    today = pd.Timestamp(date.today())
    current_month = today.to_period("M").to_timestamp()

    out_rows = []
    for _, row in df.iterrows():
        is_provisional = row["date"] == current_month

        out_rows.append({
            "agency_id"      : AGENCY_ID,
            "agency_name"    : AGENCY_NAME,
            "ntd_id"         : NTD_ID,
            "date"           : row["date"].strftime("%Y-%m-%d"),
            "metric"         : "monthly_boardings",
            "value"          : int(row["value"]),
            "mode"           : "MB",
            "unit"           : "boardings",
            "source"         : "sfmta.com",
            "source_url"     : SOURCE_URL,
            "scraped_at"     : now_utc(),
            "is_provisional" : is_provisional,
        })

    result = pd.DataFrame(out_rows)
    log.info(f"Normalized {len(result)} monthly rows ({result['is_provisional'].sum()} provisional)")
    return result[NORMALIZED_COLUMNS]


# ── Run ───────────────────────────────────────────────────────────────────────
def run() -> pd.DataFrame:
    log.info(f"Starting scrape: {AGENCY_NAME}")
    try:
        raw    = fetch()
        parsed = parse(raw)
        df     = normalize(parsed)

        out = raw_path(AGENCY_ID)
        df.to_csv(out, index=False)
        log_scrape(AGENCY_ID, "success", len(df))
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
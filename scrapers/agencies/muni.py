# scrapers/agencies/muni.py
# Muni monthly boardings scraper
# Source: SFMTA Tableau CSV (SystemwideRidershipRecovery)
# URL: https://transtat.sfmta.com/t/public/views/SystemwideRidershipRecovery/MonthlySystemwideRecoveryAccessibleTable.csv
#
# What this pulls:
#   - Monthly total boardings from April 2020 to present (~1 month lag)
#   - SFMTA's own 2019 baseline for accurate recovery % calculation
#
# Output schema: NORMALIZED_COLUMNS + baseline_sfmta
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
    for attempt in range(1, 4):
        try:
            resp = fetch_url(SOURCE_URL, timeout=60)
            log.info(f"Received {len(resp.content):,} bytes from SFMTA")
            return resp.text
        except requests.exceptions.ReadTimeout:
            if attempt < 3:
                log.warning(f"Timeout on attempt {attempt}, retrying in {5 * attempt} seconds...")
                time.sleep(5 * attempt)  # Exponential backoff: 5s, 10s, 15s
            else:
                log.error("Failed to fetch data after 3 attempts")
                raise
    


# ── Parse ─────────────────────────────────────────────────────────────────────
def parse(text: str) -> pd.DataFrame:
    """
    CSV has three columns:
      Measure Names | Month of MONTH | Measure Values

    Measure Names values:
      - "Monthly Total Boardings (accessible copy)"
      - "Baseline Monthly Total Boardings (accessible copy)"
      - "Monthly Recovery"

    We pull both boardings and baseline so recovery % uses SFMTA's own 2019 figures
    rather than NTD UPT, which uses a different methodology and would inflate recovery %.
    Month format: "April 2020"
    Values are formatted with commas: "2,555,000"
    """
    df = pd.read_csv(io.StringIO(text))
    log.info(f"Parsed {len(df)} rows, columns: {df.columns.tolist()}")
    df.columns = ["measure", "month_str", "value_str"]

    def extract(measure_name: str) -> pd.DataFrame:
        rows = df[df["measure"] == measure_name].copy()
        rows["date"] = pd.to_datetime(rows["month_str"], format="%B %Y", errors="coerce")
        rows = rows.dropna(subset=["date"])
        rows["val"] = pd.to_numeric(
            rows["value_str"].str.replace(",", "").str.strip(), errors="coerce"
        )
        rows = rows.dropna(subset=["val"])
        rows["val"] = rows["val"].astype(int)
        return rows[["date", "val"]].reset_index(drop=True)

    boardings = extract("Monthly Total Boardings (accessible copy)")
    baseline  = extract("Baseline Monthly Total Boardings (accessible copy)")

    if boardings.empty:
        raise ValueError("No 'Monthly Total Boardings' rows found — check CSV structure")

    merged = boardings.merge(
        baseline.rename(columns={"val": "baseline_sfmta"}),
        on="date", how="left"
    ).rename(columns={"val": "value"})

    log.info(f"Parsed {len(merged)} monthly boarding rows")
    log.info(f"Date range: {merged['date'].min().strftime('%b %Y')} → {merged['date'].max().strftime('%b %Y')}")
    return merged


# ── Normalize ─────────────────────────────────────────────────────────────────
def normalize(df: pd.DataFrame) -> pd.DataFrame:
    """
    Converts parsed boardings to normalized schema.
    Includes baseline_sfmta column for accurate recovery % calculation in to_excel.py.
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
            "baseline_sfmta" : int(row["baseline_sfmta"]) if pd.notna(row.get("baseline_sfmta")) else None,
        })

    result = pd.DataFrame(out_rows)
    log.info(f"Normalized {len(result)} monthly rows ({result['is_provisional'].sum()} provisional)")
    # Return with baseline_sfmta as an extra column beyond NORMALIZED_COLUMNS
    return result


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
    print(df[["date", "value", "baseline_sfmta"]].tail(6).to_string(index=False))
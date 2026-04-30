# scrapers/_template.py
# ─────────────────────────────────────────────────────────────────────────────
# TEMPLATE — copy this to scrapers/agencies/[agency].py to add a new agency.
# Steps:
#   1. Fill in the AGENCY_ constants at the top
#   2. Implement fetch(), parse(), normalize() — each does one thing
#   3. Run:  python scrapers/agencies/[agency].py
#   4. Verify output in data/raw/[agency]_YYYY-MM-DD.csv
#   5. Add the agency to merge/config.py
#   6. Open a PR — have another reporter spot-check the numbers
# ─────────────────────────────────────────────────────────────────────────────
# Bay City News

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
from scrapers._utils import (
    fetch_url, raw_path, log_scrape, now_utc, NORMALIZED_COLUMNS, log
)

# ── Agency constants — fill these in ─────────────────────────────────────────
AGENCY_ID   = "template"                 # short slug, lowercase, no spaces
AGENCY_NAME = "Template Transit Agency"
NTD_ID      = "XXXXX"
SOURCE_URL  = "https://example.com/ridership"


# ── 1. Fetch — get raw data from the source ───────────────────────────────────
def fetch() -> str:
    """
    Pull raw content from the agency source.
    Return whatever you get: HTML string, CSV text, JSON string, etc.
    Don't parse here — just fetch.
    """
    resp = fetch_url(SOURCE_URL)
    return resp.text


# ── 2. Parse — extract the numbers ───────────────────────────────────────────
def parse(raw: str) -> pd.DataFrame:
    """
    Turn the raw content into a dataframe with at minimum:
        date (YYYY-MM-01), value (numeric)
    Don't normalize column names yet — do that in normalize().
    """
    # Example for a simple CSV:
    # import io
    # df = pd.read_csv(io.StringIO(raw))
    # return df

    raise NotImplementedError("Implement parse() for this agency")


# ── 3. Normalize — conform to the standard schema ────────────────────────────
def normalize(df: pd.DataFrame) -> pd.DataFrame:
    """
    Return a dataframe with exactly the columns in NORMALIZED_COLUMNS.
    See scrapers/_utils.py for the full list and descriptions.
    """
    df = df.copy()

    # Map your columns to the standard schema here
    # df = df.rename(columns={"their_col": "our_col"})

    df["agency_id"]      = AGENCY_ID
    df["agency_name"]    = AGENCY_NAME
    df["ntd_id"]         = NTD_ID
    df["metric"]         = "upt"          # change to match what you're pulling
    df["mode"]           = "ALL"          # change as appropriate
    df["unit"]           = "trips"        # change as appropriate
    df["source"]         = AGENCY_NAME
    df["source_url"]     = SOURCE_URL
    df["scraped_at"]     = now_utc()
    df["is_provisional"] = False

    # date must be YYYY-MM-01 (first of the month)
    df["date"] = pd.to_datetime(df["date"]).dt.to_period("M").dt.to_timestamp()

    return df[NORMALIZED_COLUMNS]


# ── 4. Run — called directly or by build_master.py ───────────────────────────
def run() -> pd.DataFrame:
    """
    Orchestrates fetch → parse → normalize, saves to data/raw/, logs result.
    Returns the normalized dataframe.
    """
    log.info(f"Starting scrape: {AGENCY_NAME}")
    try:
        raw    = fetch()
        parsed = parse(raw)
        df     = normalize(parsed)

        out = raw_path(AGENCY_ID)
        df.to_csv(out, index=False)
        log_scrape(AGENCY_ID, "success", len(df))
        log.info(f"Wrote {len(df)} rows to {out}")
        return df

    except Exception as e:
        log_scrape(AGENCY_ID, "error", 0, notes=str(e))
        log.error(f"Scrape failed for {AGENCY_NAME}: {e}")
        raise


if __name__ == "__main__":
    run()

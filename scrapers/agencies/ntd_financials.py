# scrapers/agencies/ntd_financials.py
# NTD Annual Financial Data scraper
# Source: data.transportation.gov — dataset npsm-38gk
# Pulls: Fares, Operating Expenses, Unlinked Passenger Trips (annual)
# Calculates: Fare Recovery Ratio, Cost Per Trip
#
# Annual data — run monthly to catch NTD updates (published ~Oct each year)
# Output: data/raw/ntd_financials_YYYY-MM-DD.csv
#
# Bay City News | Andres Jimenez Larios

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

import io
import pandas as pd

from scrapers._utils import fetch_url, raw_path, log_scrape, now_utc, log
from merge.config import AGENCIES

AGENCY_ID  = "ntd_financials"
SOURCE_URL = "https://data.transportation.gov/resource/npsm-38gk.csv"

FIELDS = [
    "Fares",
    "Operating Expenses",
    "Unlinked Passenger Trips",
]

# Friendly display names for the Excel tab
FIELD_LABELS = {
    "Fares"                   : "Fare Revenue",
    "Operating Expenses"      : "Operating Expenses",
    "Unlinked Passenger Trips": "Annual Ridership (UPT)",
}


# ── Fetch ─────────────────────────────────────────────────────────────────────
def fetch() -> pd.DataFrame:
    ntd_ids = [cfg["ntd_id"] for cfg in AGENCIES.values()]
    ntd_ids_quoted = ",".join(f"'{i}'" for i in ntd_ids)

    api_url = (
        f"{SOURCE_URL}"
        f"?$where=ntd_id in({ntd_ids_quoted})"
        "&$limit=5000"
    )

    log.info(f"Fetching NTD annual financials for {len(ntd_ids)} agencies...")
    resp = fetch_url(api_url, timeout=60)
    df = pd.read_csv(io.StringIO(resp.text))
    log.info(f"Fetched {len(df)} rows")
    return df


# ── Parse & normalize ─────────────────────────────────────────────────────────
def parse(df: pd.DataFrame) -> pd.DataFrame:
    """
    Filters to key financial fields, sums across modes per agency/year,
    pivots to wide format, and calculates derived metrics.
    """
    # Filter to fields we want
    fin = df[df["field"].isin(FIELDS)].copy()
    fin["value"] = pd.to_numeric(fin["value"], errors="coerce")
    fin["report_year"] = pd.to_numeric(fin["report_year"], errors="coerce").astype("Int64")

    # Map ntd_id to our agency_id
    ntd_to_agency = {cfg["ntd_id"]: aid for aid, cfg in AGENCIES.items()}
    ntd_to_name   = {cfg["ntd_id"]: cfg["agency_name"] for cfg in AGENCIES.values()}

    fin["agency_id"]   = fin["ntd_id"].astype(str).map(ntd_to_agency)
    fin["agency_name"] = fin["ntd_id"].astype(str).map(ntd_to_name)

    # Sum across modes per agency/year/field
    agg = (
        fin.groupby(["agency_id", "agency_name", "ntd_id", "report_year", "field"])["value"]
        .sum()
        .reset_index()
    )

    # Pivot to wide format
    wide = agg.pivot_table(
        index=["agency_id", "agency_name", "ntd_id", "report_year"],
        columns="field",
        values="value",
        aggfunc="sum"
    ).reset_index()
    wide.columns.name = None

    # Rename columns
    wide = wide.rename(columns=FIELD_LABELS)

    # Derived metrics
    wide["Fare Recovery Ratio"] = (
        wide["Fare Revenue"] / wide["Operating Expenses"]
    ).round(4)

    wide["Cost Per Trip"] = (
        wide["Operating Expenses"] / wide["Annual Ridership (UPT)"]
    ).round(2)

    wide = wide.sort_values(["agency_id", "report_year"])
    wide["scraped_at"] = now_utc()

    log.info(f"Parsed financials: {len(wide)} agency-year rows")
    log.info(f"Year range: {wide['report_year'].min()} – {wide['report_year'].max()}")
    return wide


# ── Run ───────────────────────────────────────────────────────────────────────
def run() -> pd.DataFrame:
    log.info("Starting NTD annual financials scrape")
    try:
        raw = fetch()
        df  = parse(raw)

        out = raw_path(AGENCY_ID)
        df.to_csv(out, index=False)
        log_scrape(AGENCY_ID, "success", len(df))
        log.info(f"Wrote {len(df)} rows → {out}")
        return df

    except Exception as e:
        log_scrape(AGENCY_ID, "error", 0, notes=str(e))
        log.error(f"NTD financials scrape failed: {e}")
        raise


if __name__ == "__main__":
    df = run()
    print("\n── BART Financials ──────────────────────────────────────────────")
    bart = df[df["agency_id"] == "bart"][[
        "report_year", "Fare Revenue", "Operating Expenses",
        "Fare Recovery Ratio", "Cost Per Trip", "Annual Ridership (UPT)"
    ]]
    print(bart.to_string(index=False))
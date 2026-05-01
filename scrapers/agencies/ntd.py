# scrapers/agencies/ntd.py
# NTD Monthly Module scraper
# Pulls from data.transportation.gov Socrata API — no Cloudflare, no Excel parsing
# Dataset: Complete Monthly Ridership (8bui-9xvu)
# Updated weekly by FTA. Full history back to 2002.
#
# Bay City News | Andres Jimenez Larios

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

import io
import pandas as pd

from scrapers._utils import (
    fetch_url, raw_path, log_scrape, now_utc, NORMALIZED_COLUMNS, log
)
from merge.config import AGENCIES

AGENCY_ID  = "ntd"
SOURCE_URL = "https://data.transportation.gov/resource/8bui-9xvu.csv"

# NTD IDs we want — pulled from config so adding an agency in config.py is enough
TARGET_NTD_IDS = [cfg["ntd_id"] for cfg in AGENCIES.values()]


# ── 1. Fetch ──────────────────────────────────────────────────────────────────
def fetch() -> tuple[bytes, str]:
    """
    Queries the Socrata API on data.transportation.gov.
    Filters to our target agencies and returns all rows as CSV bytes.
    No authentication needed — this is public open data.
    """
    ntd_ids_quoted = ",".join(f"'{i}'" for i in TARGET_NTD_IDS)

    api_url = (
        f"{SOURCE_URL}"
        f"?$where=ntd_id in({ntd_ids_quoted})"
        "&$limit=50000"
        "&$order=date DESC"
    )

    log.info(f"Querying NTD Socrata API for {len(TARGET_NTD_IDS)} agencies...")
    resp = fetch_url(api_url, timeout=60)
    log.info(f"Received {len(resp.content) / 1e3:.0f} KB")

    return resp.content, api_url


# ── 2. Parse ──────────────────────────────────────────────────────────────────
def parse(csv_bytes: bytes) -> pd.DataFrame:
    """
    Socrata returns clean long-format CSV — one row per agency/mode/month.
    Columns: ntd_id, agency, mode, tos, date, upt, vrm, vrh, voms, state, etc.
    """
    df = pd.read_csv(io.BytesIO(csv_bytes))
    log.info(f"Parsed {len(df)} rows | columns: {list(df.columns)}")
    return df


# ── 3. Normalize ──────────────────────────────────────────────────────────────
def normalize(df: pd.DataFrame, api_url: str) -> pd.DataFrame:
    """
    Maps Socrata column names to our standard schema.
    Filters to HR (heavy rail) mode for BART — that's BART's core subway service.
    """
    df = df.copy()

    # Parse date — Socrata format: 2026-02-01T00:00:00.000
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df = df.dropna(subset=["date"])
    df["date"] = df["date"].dt.to_period("M").dt.to_timestamp()

    # Map ntd_id to our agency_id slug
    ntd_to_agency_id = {cfg["ntd_id"]: aid for aid, cfg in AGENCIES.items()}
    df["agency_id"] = df["ntd_id"].astype(str).map(ntd_to_agency_id).fillna("unknown")

    # Agency name
    df["agency_name"] = df["agency_id"].map(
        {aid: cfg["agency_name"] for aid, cfg in AGENCIES.items()}
    )

    # UPT is the primary ridership metric
    df["metric"]         = "upt"
    df["value"]          = pd.to_numeric(df["upt"], errors="coerce").fillna(0).astype(int)
    df["mode"]           = df["mode"].fillna("ALL")
    df["unit"]           = "trips"
    df["source"]         = "NTD Socrata API"
    df["source_url"]     = api_url
    df["scraped_at"]     = now_utc()
    df["is_provisional"] = False

    return df[NORMALIZED_COLUMNS].copy()


# ── 4. Run ────────────────────────────────────────────────────────────────────
def run() -> pd.DataFrame:
    log.info("Starting NTD scrape via Socrata API")
    try:
        csv_bytes, api_url = fetch()
        parsed = parse(csv_bytes)
        df     = normalize(parsed, api_url)

        out = raw_path(AGENCY_ID)
        df.to_csv(out, index=False)
        log_scrape(AGENCY_ID, "success", len(df))
        log.info(f"Wrote {len(df)} rows → {out}")
        return df

    except Exception as e:
        log_scrape(AGENCY_ID, "error", 0, notes=str(e))
        log.error(f"NTD scrape failed: {e}")
        raise


if __name__ == "__main__":
    df = run()
    print("\n── BART HR rows (last 6 months) ─────────────────────────")
    bart = df[
        (df["ntd_id"] == "90003") & (df["mode"] == "HR")
    ].sort_values("date").tail(6)
    print(bart[["date", "agency_name", "mode", "value", "source"]].to_string(index=False))

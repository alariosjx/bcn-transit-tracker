# merge/build_master.py
# Merges agency-direct scrapes with NTD baseline
# Flags discrepancies, writes master_monthly.csv
# Bay City News

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import glob
import pandas as pd
from datetime import datetime, timezone

from merge.config import (
    AGENCIES, RAW_DIR, PROCESSED_DIR,
    VARIANCE_REVIEW, VARIANCE_INVESTIGATE
)
from scrapers._utils import log, now_utc

PROCESSED_DIR.mkdir(parents=True, exist_ok=True)


# ── Load latest raw scrape for an agency ─────────────────────────────────────
def load_latest_raw(agency_id: str) -> pd.DataFrame | None:
    pattern = str(RAW_DIR / f"{agency_id}_2*.csv")
    files   = sorted(glob.glob(pattern), reverse=True)
    if not files:
        log.warning(f"No raw file found for {agency_id}")
        return None
    latest = files[0]
    log.info(f"Loading raw: {latest}")
    return pd.read_csv(latest, parse_dates=["date"])

#-─ Load BART backfill data ─────────────────────────────────────────────────

def load_bart_backfill() -> pd.DataFrame | None:
    path = RAW_DIR / "bart_backfill.csv"
    if not path.exists():
        log.warning("No bart_backfill.csv found — run scrapers/agencies/bart_backfill.py")
        return None
    df = pd.read_csv(path, parse_dates=["date"])
    log.info(f"Loaded BART backfill: {len(df)} rows ({df['date'].min().strftime('%b %Y')} → {df['date'].max().strftime('%b %Y')})")
    return df


# ── Load NTD monthly data ─────────────────────────────────────────────────────
def load_ntd() -> pd.DataFrame | None:
    """
    Loads NTD data from the Socrata API scrape (data/raw/ntd_*.csv).
    Falls back gracefully if no file found.
    """
    pattern = str(RAW_DIR / "ntd_2*.csv")
    files   = sorted(glob.glob(pattern), reverse=True)
    if not files:
        log.warning("No NTD raw file found. Run scrapers/agencies/ntd.py first.")
        return None
    latest = files[0]
    log.info(f"Loading NTD: {latest}")
    df = pd.read_csv(latest, parse_dates=["date"])

    # Socrata returns all modes — we need to map to our schema
    # Rename columns to match what get_ntd_for_agency expects
    if "agency" in df.columns and "agency_name" not in df.columns:
        df = df.rename(columns={"agency": "agency_name"})
    if "ntd_id" not in df.columns:
        log.warning("ntd_id column missing from NTD file")
        return None

    df["ntd_id"] = df["ntd_id"].astype(str)
    df["upt"]    = pd.to_numeric(df.get("upt", df.get("value", 0)), errors="coerce").fillna(0)
    log.info(f"Loaded NTD: {len(df)} rows, modes: {df['mode'].unique().tolist()}")
    return df


# ── Aggregate NTD for one agency ─────────────────────────────────────────────
def get_ntd_for_agency(ntd: pd.DataFrame, ntd_id: str, modes: list) -> pd.DataFrame:
    """
    Filters NTD to one agency, sums across modes, returns monthly UPT series.
    """
    mask = ntd["ntd_id"].astype(str) == str(ntd_id)
    if modes != ["ALL"]:
        mask &= ntd["mode"].isin(modes)

    agency_ntd = (
        ntd[mask]
        .groupby("date")["upt"]
        .sum()
        .reset_index()
        .rename(columns={"upt": "ntd_value", "date": "date"})
    )
    agency_ntd["date"] = pd.to_datetime(agency_ntd["date"])
    return agency_ntd


# ── Flag variance ─────────────────────────────────────────────────────────────
def flag_variance(variance_pct: float | None) -> str:
    if variance_pct is None or pd.isna(variance_pct):
        return ""
    abs_var = abs(variance_pct)
    if abs_var >= VARIANCE_INVESTIGATE:
        return "INVESTIGATE"
    elif abs_var >= VARIANCE_REVIEW:
        return "REVIEW"
    return ""


# ── Build master for one agency ───────────────────────────────────────────────
def build_agency(agency_id: str, config: dict, ntd: pd.DataFrame | None) -> pd.DataFrame:
    """
    Merges agency-direct + NTD for a single agency.
    Returns a dataframe with one row per month.
    """
    rows = []
    agency_name  = config["agency_name"]
    ntd_id       = config["ntd_id"]
    ntd_modes    = config["ntd_modes"]
    use_direct   = config["primary_source"] == "agency_direct"

    # Load agency-direct data
    if use_direct:
        direct_df = load_latest_raw(agency_id)
        
        # For BART: merge backfill (2018-2024) with daily scraper (2025+)
        # Backfill uses bart.gov OD archives for source consistency
        if agency_id == "bart":
            backfill = load_bart_backfill()
            if backfill is not None and direct_df is not None:
                # Combine: backfill first, then daily scraper on top
                # sort_values + drop_duplicates keeps the last (daily scraper) for overlapping dates
                combined = pd.concat([backfill, direct_df], ignore_index=True)
                combined["date"] = combined["date"].dt.to_period("M").dt.to_timestamp()
                direct_df = combined.sort_values("date").drop_duplicates("date", keep="last")
            elif backfill is not None:
                direct_df = backfill
    else:
        direct_df = None
    # Load NTD data for this agency
    ntd_agency = get_ntd_for_agency(ntd, ntd_id, ntd_modes) if ntd is not None else pd.DataFrame()

    # Build a unified date index
    all_dates = set()
    if direct_df is not None and not direct_df.empty:
        all_dates.update(direct_df["date"].dt.to_period("M").dt.to_timestamp().unique())
    if not ntd_agency.empty:
        all_dates.update(ntd_agency["date"].dt.to_period("M").dt.to_timestamp().unique())

    for month_date in sorted(all_dates):

        # Agency-direct value
        direct_row = None
        if direct_df is not None and not direct_df.empty:
            mask = direct_df["date"].dt.to_period("M").dt.to_timestamp() == month_date
            matches = direct_df[mask]
            if not matches.empty:
                direct_row = matches.iloc[0]

        # NTD value
        ntd_val = None
        if not ntd_agency.empty:
            mask = ntd_agency["date"].dt.to_period("M").dt.to_timestamp() == month_date
            matches = ntd_agency[mask]
            if not matches.empty:
                ntd_val = float(matches.iloc[0]["ntd_value"])

        # Determine primary value and source
        if direct_row is not None:
            primary_val    = float(direct_row["value"])
            primary_source = str(direct_row["source"])
            is_provisional = bool(direct_row.get("is_provisional", False))
        elif ntd_val is not None:
            primary_val    = ntd_val
            primary_source = "NTD"
            is_provisional = False
        else:
            continue

        # Variance calculation
        if direct_row is not None and ntd_val is not None:
            variance_pct = (primary_val - ntd_val) / ntd_val
        else:
            variance_pct = None

        rows.append({
            "agency_id"    : agency_id,
            "agency_name"  : agency_name,
            "ntd_id"       : ntd_id,
            "date"         : month_date.strftime("%Y-%m-%d"),
            "metric"       : direct_row["metric"] if direct_row is not None else "upt",
            "value"        : int(primary_val),
            "source"       : primary_source,
            "ntd_value"    : int(ntd_val) if ntd_val is not None else None,
            "variance_pct" : round(variance_pct, 4) if variance_pct is not None else None,
            "flag"         : "" if config.get("variance_ignore") else flag_variance(variance_pct),
            "is_provisional": is_provisional,
            "last_updated" : now_utc(),
        })

    df = pd.DataFrame(rows)
    log.info(
        f"{agency_id}: {len(df)} months | "
        f"{(df['flag'] == 'INVESTIGATE').sum()} INVESTIGATE | "
        f"{(df['flag'] == 'REVIEW').sum()} REVIEW"
    )
    return df


# ── Main ──────────────────────────────────────────────────────────────────────
def run():
    log.info("── build_master.py starting ──")
    ntd = load_ntd()

    all_frames = []
    flags_frames = []

    for agency_id, config in AGENCIES.items():
        log.info(f"Processing: {agency_id}")
        df = build_agency(agency_id, config, ntd)
        all_frames.append(df)

        flagged = df[df["flag"].isin(["REVIEW", "INVESTIGATE"])]
        if not flagged.empty:
            flags_frames.append(flagged)

    master = pd.concat(all_frames, ignore_index=True) if all_frames else pd.DataFrame()

    # Write master
    master_path = PROCESSED_DIR / "master_monthly.csv"
    master.to_csv(master_path, index=False)
    log.info(f"Wrote master: {len(master)} rows → {master_path}")

    # Write variance flags
    if flags_frames:
        flags = pd.concat(flags_frames, ignore_index=True)
        flags_path = PROCESSED_DIR / "_variance_flags.csv"
        flags.to_csv(flags_path, index=False)
        log.warning(
            f"{len(flags)} rows flagged for review → {flags_path}\n"
            f"  INVESTIGATE: {(flags['flag'] == 'INVESTIGATE').sum()}\n"
            f"  REVIEW:      {(flags['flag'] == 'REVIEW').sum()}"
        )
    else:
        log.info("No variance flags — all discrepancies within threshold")

    log.info("── build_master.py complete ──")
    return master


if __name__ == "__main__":
    master = run()
    print("\n── Master (last 6 rows) ─────────────────────────────")
    print(master.tail(6).to_string(index=False))

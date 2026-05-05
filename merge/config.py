# merge/config.py
# Central registry of agencies and merge settings
# Add a new agency here after its scraper is verified
# Bay City News

from pathlib import Path

ROOT          = Path(__file__).resolve().parent.parent
RAW_DIR       = ROOT / "data" / "raw"
PROCESSED_DIR = ROOT / "data" / "processed"
PUBLIC_DIR    = ROOT / "data" / "public"

# ── Variance thresholds (agency-direct vs NTD) ────────────────────────────────
# When both sources exist for the same agency/month, we compare them.
# Anything above INVESTIGATE triggers a flag in the output CSV.
VARIANCE_REVIEW      = 0.05   # 5% — flag as REVIEW
VARIANCE_INVESTIGATE = 0.10   # 10% — flag as INVESTIGATE

# ── Agency registry ───────────────────────────────────────────────────────────
# Each entry maps an agency_id (used in filenames) to its NTD ID and metadata.
# Add a new agency here after its scraper is verified.
#
# primary_source options:
#   "agency_direct"  — scraper exists and is the preferred source
#   "ntd_only"       — no direct scraper yet; falls back to NTD entirely
#
# ntd_modes: which NTD mode codes to sum for this agency's total ridership
#   "ALL" = sum everything; or a list like ["MB", "LR"]

AGENCIES = {
    "bart": {
        "agency_name"    : "Bay Area Rapid Transit",
        "ntd_id"         : "90003",
        "ntd_modes"      : ["HR"],
        "primary_source" : "agency_direct",
        "scraper"        : "scrapers.agencies.bart",
        "notes"          : (
            "BART exits run ~7.5% lower than NTD adjusted UPT. "
            "Variance up to 10% is expected and documented."
        ),
    },
    # ── Future agencies — uncomment and build scraper when ready ──────────────
    "muni": {
        "agency_name"    : "San Francisco Municipal Railway",
        "ntd_id"         : "90015",
        "ntd_modes"      : ["MB", "LR"],
        "primary_source" : "agency_direct",
        "scraper"        : "scrapers.agencies.muni",
        "variance_ignore": True,  # SFMTA total boardings ~35-40% higher than NTD — expected
        "notes"          : (
            "SFMTA total boardings run ~35-40% higher than NTD adjusted UPT. "
            "Different methodology — SFMTA counts all boardings including transfers. "
            "Variance up to 45% is expected and documented."
        ),
    },
    "smart": {
        "agency_name"    : "Sonoma-Marin Area Rail Transit",
        "ntd_id"         : "90299",
        "ntd_modes"      : ["CR"],
        "primary_source" : "agency_direct",
        "scraper"        : "scrapers.agencies.smart",
        "variance_ignore": True,
        "notes"          : (
            "SMART Excel matches NTD CR exactly from Oct 2022 onward (APC system). "
            "Aug 2017 shows variance — NTD startup artifact, agency number is correct. "
            "variance_ignore=True since numbers are verified identical in normal operation."
        ),
    },
    # "ac_transit": {
    #     "agency_name"    : "AC Transit",
    #     "ntd_id"         : "90014",
    #     "ntd_modes"      : ["ALL"],
    #     "primary_source" : "ntd_only",
    #     "scraper"        : None,
    #     "notes"          : "",
    # },
    # "caltrain": {
    #     "agency_name"    : "Caltrain",
    #     "ntd_id"         : "90134",
    #     "ntd_modes"      : ["CR"],
    #     "primary_source" : "ntd_only",
    #     "scraper"        : None,
    #     "notes"          : "",
    # },
    # "vta": {
    #     "agency_name"    : "Santa Clara VTA",
    #     "ntd_id"         : "90013",
    #     "ntd_modes"      : ["ALL"],
    #     "primary_source" : "ntd_only",
    #     "scraper"        : None,
    #     "notes"          : "",
    # },
}

# ── NTD metrics to include in agency profiles ─────────────────────────────────
# These are pulled from NTD and surfaced in the Agency Profiles Excel tab.
NTD_PROFILE_METRICS = [
    "upt",    # Unlinked Passenger Trips
    "vrm",    # Vehicle Revenue Miles
    "vrh",    # Vehicle Revenue Hours
    "voms",   # Vehicles Operated in Maximum Service
]

# ── NTD financial metrics (from NTD annual data — separate from monthly) ──────
# These go in the Agency Profiles sheet for context
NTD_FINANCIAL_METRICS = [
    "total_operating_expenses",
    "total_operating_revenues",
    "fare_revenues",
    "federal_funds",
    "state_funds",
    "local_funds",
]

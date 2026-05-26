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
        "variance_ignore": True,
        "notes"          : (
            "bart.gov exits and NTD UPT are different metrics — variance is expected and documented. "
            "Source priority: bart.gov OD archives (Feb 2018–Dec 2024), daily scraper (Jan 2025+), NTD (Jan 2002–Jan 2018)."
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
    "actransit": {
        "agency_name"    : "AC Transit",
        "ntd_id"         : "90014",
        "ntd_modes"      : ["MB", "RB", "CB"],
        "primary_source" : "ntd",
        "scraper"        : None,
        "variance_ignore": False,
        "notes"          : (
            "AC Transit NTD modes: MB (motor bus) + RB (rapid bus) + CB (commuter bus) summed. "
            "CB is transbay bus service between East Bay and SF — tracked separately for story context."
        ),
    },
    "caltrain": {
        "agency_name"    : "Caltrain",
        "ntd_id"         : "90134",
        "ntd_modes"      : ["CR"],
        "primary_source" : "ntd",
        "scraper"        : None,
        "variance_ignore": False,
        "notes"          : (
            "Caltrain NTD mode CR only. MB rows (~417K/year) are shuttles — excluded. "
            "Tableau public dashboard available but not scrapable; NTD is authoritative."
        ),
    },
    "sfbayferry": {
        "agency_name"    : "SF Bay Ferry",
        "ntd_id"         : "90225",
        "ntd_modes"      : ["FB"],
        "primary_source" : "ntd",
        "scraper"        : None,
        "variance_ignore": False,
        "notes"          : "SF Bay Ferry (WETA). Ferry boat only.",
    },
    "samtrans": {
        "agency_name"    : "SamTrans",
        "ntd_id"         : "90009",
        "ntd_modes"      : ["MB"],
        "primary_source" : "ntd",
        "scraper"        : None,
        "variance_ignore": False,
        "notes"          : "San Mateo County Transit District. MB only — DR excluded.",
    },
    "vta": {
        "agency_name"    : "VTA",
        "ntd_id"         : "90013",
        "ntd_modes"      : ["MB", "LR"],
        "primary_source" : "ntd",
        "scraper"        : None,
        "variance_ignore": False,
        "notes"          : "Santa Clara VTA. MB + LR summed. DR excluded.",
    },
    "ggferry": {
        "agency_name"    : "Golden Gate Ferry",
        "ntd_id"         : "90016",
        "ntd_modes"      : ["FB"],
        "primary_source" : "ntd",
        "scraper"        : None,
        "variance_ignore": False,
        "notes"          : "Golden Gate Ferry only. Split from Golden Gate Bus.",
    },
    "ggbus": {
        "agency_name"    : "Golden Gate Bus",
        "ntd_id"         : "90016",
        "ntd_modes"      : ["MB"],
        "primary_source" : "ntd",
        "scraper"        : None,
        "variance_ignore": False,
        "notes"          : "Golden Gate Bus only. Split from Golden Gate Ferry.",
    },
    "marin": {
        "agency_name"    : "Marin Transit",
        "ntd_id"         : "90234",
        "ntd_modes"      : ["MB"],
        "primary_source" : "ntd",
        "scraper"        : None,
        "variance_ignore": False,
        "notes"          : "Marin County Transit District. MB only.",
    },
    "napa": {
        "agency_name"    : "Napa Valley Transportation Authority",
        "ntd_id"         : "90088",
        "ntd_modes"      : ["MB", "CB"],
        "primary_source" : "ntd",
        "scraper"        : None,
        "variance_ignore": False,
        "notes"          : "NVTA. MB + CB summed.",
    },
    "vallejo": {
        "agency_name"    : "Vallejo Transit",
        "ntd_id"         : "90028",
        "ntd_modes"      : ["MB", "CB"],
        "primary_source" : "ntd",
        "scraper"        : None,
        "variance_ignore": False,
        "notes"          : "City of Vallejo. MB + CB summed.",
    },
    "countyconnection": {
        "agency_name"    : "County Connection",
        "ntd_id"         : "90078",
        "ntd_modes"      : ["MB"],
        "primary_source" : "ntd",
        "scraper"        : None,
        "variance_ignore": False,
        "notes"          : "Central Contra Costa Transit Authority.",
    },
    "westcat": {
        "agency_name"    : "WestCAT",
        "ntd_id"         : "90159",
        "ntd_modes"      : ["MB", "CB"],
        "primary_source" : "ntd",
        "scraper"        : None,
        "variance_ignore": False,
        "notes"          : "Western Contra Costa Transit Authority.",
    },
    "tridelta": {
        "agency_name"    : "Tri Delta Transit",
        "ntd_id"         : "90162",
        "ntd_modes"      : ["MB", "DR"],
        "primary_source" : "ntd",
        "scraper"        : None,
        "variance_ignore": False,
        "notes"          : "Eastern Contra Costa Transit Authority.",
    },
    "wheels": {
        "agency_name"    : "Wheels (LAVTA)",
        "ntd_id"         : "90144",
        "ntd_modes"      : ["MB"],
        "primary_source" : "ntd",
        "scraper"        : None,
        "variance_ignore": False,
        "notes"          : "Livermore/Amador Valley Transit Authority.",
    },
    "unioncity": {
        "agency_name"    : "Union City Transit",
        "ntd_id"         : "90161",
        "ntd_modes"      : ["MB"],
        "primary_source" : "ntd",
        "scraper"        : None,
        "variance_ignore": False,
        "notes"          : "City of Union City.",
    },
    "alamedaferry": {
        "agency_name"    : "Alameda Ferry",
        "ntd_id"         : "90150",
        "ntd_modes"      : ["FB"],
        "primary_source" : "ntd",
        "scraper"        : None,
        "variance_ignore": False,
        "notes"          : "City of Alameda Ferry Services.",
    },
    "santarosa": {
        "agency_name"    : "Santa Rosa CityBus",
        "ntd_id"         : "90017",
        "ntd_modes"      : ["MB"],
        "primary_source" : "ntd",
        "scraper"        : None,
        "variance_ignore": False,
        "notes"          : "City of Santa Rosa.",
    },
    "fairfield": {
        "agency_name"    : "Fairfield-Suisun Transit",
        "ntd_id"         : "90092",
        "ntd_modes"      : ["MB"],
        "primary_source" : "ntd",
        "scraper"        : None,
        "variance_ignore": False,
        "notes"          : "City of Fairfield.",
    },
    "vacaville": {
        "agency_name"    : "Vacaville City Coach",
        "ntd_id"         : "90155",
        "ntd_modes"      : ["MB"],
        "primary_source" : "ntd",
        "scraper"        : None,
        "variance_ignore": False,
        "notes"          : "City of Vacaville.",
    },
    "petaluma": {
        "agency_name"    : "Petaluma Transit",
        "ntd_id"         : "90213",
        "ntd_modes"      : ["MB"],
        "primary_source" : "ntd",
        "scraper"        : None,
        "variance_ignore": False,
        "notes"          : "City of Petaluma.",
    },
    "mst": {
        "agency_name"    : "Monterey-Salinas Transit",
        "ntd_id"         : "90062",
        "ntd_modes"      : ["MB", "CB"],
        "primary_source" : "ntd",
        "scraper"        : None,
        "variance_ignore": False,
        "notes"          : "Monterey-Salinas Transit. Extended coverage county.",
    },
    "santacruz": {
        "agency_name"    : "Santa Cruz Metro",
        "ntd_id"         : "90006",
        "ntd_modes"      : ["MB"],
        "primary_source" : "ntd",
        "scraper"        : None,
        "variance_ignore": False,
        "notes"          : "Santa Cruz Metropolitan Transit District. Extended coverage county.",
    },
    "sjrtd": {
        "agency_name"    : "San Joaquin RTD",
        "ntd_id"         : "90012",
        "ntd_modes"      : ["MB", "CB"],
        "primary_source" : "ntd",
        "scraper"        : None,
        "variance_ignore": False,
        "notes"          : "San Joaquin Regional Transit District. Extended coverage county.",
    },
    "ace": {
        "agency_name"    : "Altamont Corridor Express",
        "ntd_id"         : "90182",
        "ntd_modes"      : ["CR"],
        "primary_source" : "ntd",
        "scraper"        : None,
        "variance_ignore": False,
        "notes"          : "ACE commuter rail Stockton to San Jose.",
    },
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

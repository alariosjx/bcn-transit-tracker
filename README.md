# BCN Transit Tracker

**Bay City News** | Maintained by [Andres Jimenez Larios](https://github.com/alariosjx) and the BCN data team.

A reproducible, automated system for tracking Bay Area transit ridership. Pulls data directly from agency sources when available, falls back to the National Transit Database (NTD), flags discrepancies between the two, and exports analyst-ready CSV and Excel files. Charts are produced in BCN house style using R/ggplot.

Designed to run on a reporter's laptop **or** automatically via GitHub Actions.

---

## What this produces

| File | Description |
|---|---|
| `data/public/bcn_transit_master.xlsx` | Excel with multiple tabs: master data, monthly pivot, YoY, agency profiles, methodology |
| `data/public/bart_monthly.csv` | BART monthly exits, clean and ready for Datawrapper/Flourish |
| `data/public/agency_profiles.csv` | NTD financial + ridership profile per agency |
| `viz/` | BCN-styled PNG charts |

---

## Agencies covered

| Agency | Source | Frequency | NTD ID |
|---|---|---|---|
| BART | bart.gov daily exits | Daily | 90003 |
| *More agencies added one at a time as each is verified* | | | |

---

## Setup (one time)

```bash
# Clone the repo
git clone https://github.com/baycitnews/bcn-transit-tracker.git
cd bcn-transit-tracker

# Create virtual environment
python -m venv venv
source venv/bin/activate      # Mac/Linux
# venv\Scripts\activate       # Windows

# Install dependencies
pip install -r requirements.txt
```

For the R/analysis layer:
```r
# In RStudio, open analysis/bcn_transit_tracker.Rproj
renv::restore()
```

---

## Running locally

```bash
# Activate your venv first
source venv/bin/activate

# Scrape BART (writes to data/raw/)
python scrapers/agencies/bart.py

# Build master dataset (merges agency-direct + NTD, flags discrepancies)
python merge/build_master.py

# Export to Excel + public CSVs
python export/to_excel.py
```

Or run everything in one command:
```bash
python run_all.py
```

---

## GitHub Actions (automated)

Three workflows run automatically:

| Workflow | Schedule | What it does |
|---|---|---|
| `scrape-daily.yml` | Weekdays, 6 AM PT | Scrapes BART daily exits, commits raw data |
| `scrape-monthly.yml` | 5th of each month, 7 AM PT | Scrapes monthly sources + NTD check, rebuilds master |
| `build-deliverables.yml` | On push to `data/processed/` | Rebuilds Excel and public CSVs |

All workflows also have a **Run workflow** button in the Actions tab for manual runs.

No API keys are required for any current source. If a future source requires
credentials, add them to **Settings → Secrets and variables → Actions** —
never commit keys to the repo.

---

## Adding a new agency

1. Copy `scrapers/agencies/_template.py` to `scrapers/agencies/[agency].py`
2. Fill in the four required functions: `fetch()`, `parse()`, `normalize()`, `run()`
3. Add the agency to `merge/config.py`
4. Run `python scrapers/agencies/[agency].py` and verify the output in `data/raw/`
5. Run `python merge/build_master.py` and check the variance flags
6. Open a PR — another reporter should spot-check the numbers before merging

---

## Discrepancy flags

When both an agency-direct source and NTD report the same month, the master
dataset includes both values and a variance column:

| Flag | Meaning |
|---|---|
| *(blank)* | Variance < 5% — within expected range |
| `REVIEW` | Variance 5–10% — worth a look |
| `INVESTIGATE` | Variance > 10% — do not publish without understanding the difference |

Thresholds are in `merge/config.py`.

---

## Repo structure

```
bcn-transit-tracker/
├── .github/workflows/          # GitHub Actions YAML files
├── scrapers/
│   ├── _utils.py               # Shared HTTP, logging, timestamp helpers
│   ├── _template.py            # Copy this to add a new agency
│   └── agencies/
│       └── bart.py             # BART daily exits scraper
├── merge/
│   ├── config.py               # Agency list, thresholds, NTD IDs
│   └── build_master.py         # Merges all sources, flags discrepancies
├── export/
│   └── to_excel.py             # Builds Excel + public CSVs
├── analysis/                   # R scripts (BCN charts, gt tables)
├── R/                          # bcn_style.R and other shared R helpers
├── data/
│   ├── raw/                    # Timestamped scrape output (gitignored, large files)
│   ├── processed/              # Cleaned + merged data
│   └── public/                 # Reporter-facing deliverables (committed to repo)
├── viz/                        # Output charts (committed to repo)
├── docs/                       # Methodology, architecture notes
├── run_all.py                  # Convenience script — runs full pipeline
├── requirements.txt
├── .gitignore
└── README.md
```

---

## Credits

Built and maintained by the [Bay City News](https://www.baycitynews.com) data team.
Lead: Andres Jimenez Larios ([@alariosjx](https://github.com/alariosjx))

Data sources: [bart.gov](https://www.bart.gov), [SFMTA](https://www.sfmta.com),
[National Transit Database](https://www.transit.dot.gov/ntd)

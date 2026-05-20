#!/usr/bin/env python3
# scrapers/agencies/bart_backfill.py
# ONETIME backfill of BART monthly ridership from official zip archives
# Covers 2010-2024 using the Total Trips OD sheet in each monthly xlsx
#
# Run once: python scrapers/agencies/bart_backfill.py
# Output: data/raw/bart_backfill.csv
#
# Bay City News | Andres Jimenez Larios

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

import io
import re
import zipfile
from datetime import date

import pandas as pd
import requests

from scrapers._utils import log, now_utc, NORMALIZED_COLUMNS

AGENCY_ID   = "bart"
AGENCY_NAME = "Bay Area Rapid Transit"
NTD_ID      = "90003"
SOURCE      = "bart.gov"
BASE        = "https://www.bart.gov/sites/default/files"

HEADERS = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)"}

# ── URL map per year ──────────────────────────────────────────────────────────
# Format: year -> (url, sheet_name)
# Sheet name varies between eras
YEAR_URLS = {
    2018: (f"{BASE}/docs/ridership_2018.zip",            "Total Trips OD"),
    2019: (f"{BASE}/docs/ridership_2019.zip",            "Total Trips OD"),
    2020: (f"{BASE}/docs/ridership_2020.zip",            "Total Trips OD"),
    2021: (f"{BASE}/docs/ridership_2021.zip",            "Total Trips OD"),
    2022: (f"{BASE}/docs/Ridership_2022.zip",            "Total Trips OD"),
    2023: (f"{BASE}/2024-02/ridership_2023.zip",         "Total Trips OD"),
    2024: (f"{BASE}/2025-02/ridership_2024.zip",         "Total Trips OD"),
}

# Month name -> number for older filename formats
MONTH_NAMES = {
    "january": 1, "february": 2, "march": 3, "april": 4,
    "may": 5, "june": 6, "july": 7, "august": 8,
    "september": 9, "october": 10, "november": 11, "december": 12
}


# ── Parse month from filename ─────────────────────────────────────────────────
def parse_month_from_filename(filename: str) -> tuple[int, int] | None:
    """
    Extracts (year, month) from BART zip filenames across all naming conventions:
      - Ridership_201901.xlsx  -> (2019, 1)
      - Ridership_April2010.xlsx -> (2010, 4)
      - Ridership_April2013.xlsx -> (2013, 4)
    Returns None if parsing fails.
    """
    # Strip directory prefix and extension
    name = Path(filename).stem.lower()

    # Pattern 1: YYYYMM e.g. ridership_201901
    m = re.search(r'(\d{4})(\d{2})$', name)
    if m:
        return int(m.group(1)), int(m.group(2))

    # Pattern 2: MonthYYYY e.g. ridership_april2010
    m = re.search(r'([a-z]+)(\d{4})$', name)
    if m:
        month_name = m.group(1)
        year = int(m.group(2))
        month = MONTH_NAMES.get(month_name)
        if month:
            return year, month

    return None


# ── Extract grand total from sheet ────────────────────────────────────────────
def extract_grand_total(sheet_bytes: bytes, sheet_name: str) -> int | None:
    try:
        xl = pd.ExcelFile(io.BytesIO(sheet_bytes), engine="openpyxl")
        for name in [sheet_name, "Total Trips OD", "Total Trips"]:
            if name not in xl.sheet_names:
                continue
            df = xl.parse(name, header=None)

            # Method 1: grand total cell (last row, last col)
            grand_total = pd.to_numeric(df.iloc[-1].iloc[-1], errors="coerce")
            if pd.notna(grand_total) and grand_total > 100000:
                return int(round(grand_total))

            # Method 2: second to last column
            grand_total = pd.to_numeric(df.iloc[-1].iloc[-2], errors="coerce")
            if pd.notna(grand_total) and grand_total > 100000:
                return int(round(grand_total))

            # Method 3: matrix sum (fallback for files with blank totals)
            matrix = df.iloc[1:-2, 1:-1]
            matrix_sum = pd.to_numeric(matrix.stack(), errors="coerce").sum()
            if matrix_sum > 100000:
                return int(round(matrix_sum))

        return None
    except Exception as e:
        log.warning(f"Failed to extract grand total: {e}")
        return None


# ── Process one zip ───────────────────────────────────────────────────────────
def process_zip(year: int, url: str, sheet_name: str) -> list[dict]:
    log.info(f"Downloading {year} zip: {url}")
    r = requests.get(url, headers=HEADERS, timeout=60)
    if r.status_code != 200:
        log.warning(f"Failed to download {year}: HTTP {r.status_code}")
        return []

    log.info(f"  {year}: {len(r.content):,} bytes")

    z = zipfile.ZipFile(io.BytesIO(r.content))
    xlsx_files = [f for f in z.namelist() if f.lower().endswith(".xlsx")]

    rows = []
    for filename in sorted(xlsx_files):
        parsed = parse_month_from_filename(filename)
        if not parsed:
            log.warning(f"  Could not parse date from: {filename}")
            continue

        file_year, month = parsed
        cal_date = date(file_year, month, 1)

        # Skip future dates
        if cal_date > date.today():
            continue

        total = extract_grand_total(z.read(filename), sheet_name)
        if total is None or total <= 0:
            log.warning(f"  No valid total for {filename}")
            continue

        log.info(f"  {cal_date.strftime('%b %Y')}: {total:,}")
        rows.append({
            "agency_id"      : AGENCY_ID,
            "agency_name"    : AGENCY_NAME,
            "ntd_id"         : NTD_ID,
            "date"           : cal_date.strftime("%Y-%m-%d"),
            "metric"         : "monthly_exits",
            "value"          : total,
            "mode"           : "HR",
            "unit"           : "exits",
            "source"         : SOURCE,
            "source_url"     : url,
            "scraped_at"     : now_utc(),
            "is_provisional" : False,
        })

    return rows


# ── Run ───────────────────────────────────────────────────────────────────────
def run():
    log.info("── BART Historical Backfill (2010–2024) ──────────────────────")
    all_rows = []

    for year, (url, sheet_name) in sorted(YEAR_URLS.items()):
        rows = process_zip(year, url, sheet_name)
        all_rows.extend(rows)
        log.info(f"  {year}: {len(rows)} months processed")

    if not all_rows:
        log.error("No rows extracted — check URLs and sheet names")
        return

    df = pd.DataFrame(all_rows)[NORMALIZED_COLUMNS]
    df = df.sort_values("date").drop_duplicates("date", keep="last")

    out = Path(__file__).resolve().parent.parent.parent / "data" / "raw" / "bart_backfill.csv"
    df.to_csv(out, index=False)

    log.info(f"\n── Summary ──────────────────────────────────────────────────")
    log.info(f"Total months: {len(df)}")
    log.info(f"Date range:   {df['date'].min()} → {df['date'].max()}")
    log.info(f"Wrote: {out}")

    print("\n── Sample (last 6 months) ───────────────────────────────────")
    print(df.tail(6)[["date", "value"]].to_string(index=False))


if __name__ == "__main__":
    run()
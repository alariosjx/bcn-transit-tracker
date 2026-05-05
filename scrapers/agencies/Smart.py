# scrapers/agencies/smart.py
# SMART monthly ridership scraper
# Source: Sonoma-Marin Area Rail Transit ridership Excel
# URL: https://www.sonomamarintrain.org/RidershipReports
#
# What this pulls:
#   - Total Monthly Ridership from FY18 (Aug 2017) to present
#   - Converts fiscal year layout (Jul–Jun) to calendar dates
#   - NTD ID: 90299, mode CR — numbers match agency exactly
#
# Output schema: NORMALIZED_COLUMNS
# Output file:   data/raw/smart_YYYY-MM-DD.csv
#
# Bay City News | Andres Jimenez Larios

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

import io
import re
import requests
from datetime import date

import pandas as pd
from bs4 import BeautifulSoup

from scrapers._utils import (
    fetch_url, raw_path, log_scrape, now_utc, NORMALIZED_COLUMNS, log
)

AGENCY_ID   = "smart"
AGENCY_NAME = "Sonoma-Marin Area Rail Transit"
NTD_ID      = "90299"
BASE_URL    = "https://www.sonomamarintrain.org"
INDEX_URL   = f"{BASE_URL}/RidershipReports"

# Fiscal year month order — FY runs Jul through Jun
FY_MONTH_ORDER = ["Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar","Apr","May","Jun"]

# Map month abbreviation → month number
MONTH_MAP = {
    "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4,
    "May": 5, "Jun": 6, "Jul": 7, "Aug": 8,
    "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12,
}


# ── Find Excel URL ────────────────────────────────────────────────────────────
def find_excel_url() -> str:
    """
    Scrapes the SMART ridership page to find the current Excel download URL.
    The filename changes each month (e.g. SMART Ridership Web Posting_3.26.xlsx).
    """
    resp = fetch_url(INDEX_URL)
    soup = BeautifulSoup(resp.content, "html.parser")

    for a in soup.find_all("a", href=True):
        href = a["href"]
        if ".xlsx" in href and "Ridership" in href:
            url = BASE_URL + href if href.startswith("/") else href
            log.info(f"Found Excel URL: {url}")
            return url

    raise ValueError("Could not find Excel download link on SMART ridership page")


# ── Fetch ─────────────────────────────────────────────────────────────────────
def fetch(excel_url: str) -> bytes:
    resp = fetch_url(excel_url)
    log.info(f"Downloaded Excel: {len(resp.content):,} bytes")
    return resp.content


# ── Parse ─────────────────────────────────────────────────────────────────────
def parse(excel_bytes: bytes) -> pd.DataFrame:
    """
    Parses the SMART ridership Excel.

    Layout (Sheet1):
      Row 3:  header — "Month", FY18, FY19, ..., FY26
      Rows 4-15: monthly ridership by fiscal year (Jul through Jun)
      Row 16: TOTAL row — skip
      Rows 18+: Average Weekday Ridership — skip

    Fiscal year conversion:
      FY18 = FY ending June 2018 → Jul 2017 through Jun 2018
      FY26 = FY ending June 2026 → Jul 2025 through Jun 2026

    So for FY=N:
      Jul–Dec → calendar year = N - 1
      Jan–Jun → calendar year = N
    """
    xl = pd.ExcelFile(io.BytesIO(excel_bytes))
    df = xl.parse("Sheet1", header=None)

    # Find header row — contains "Month" and "FY" columns
    header_row = None
    for i, row in df.iterrows():
        if str(row.iloc[1]).strip() == "Month":
            header_row = i
            break

    if header_row is None:
        raise ValueError("Could not find header row in SMART Excel")

    # Extract FY columns and their indices
    header = df.iloc[header_row]
    fy_cols = {}
    for col_idx, val in header.items():
        val_str = str(val).strip()
        if re.match(r"FY\d{2}", val_str):
            fy_num = int(val_str[2:])  # e.g. "FY26" → 26
            fy_year = 2000 + fy_num    # e.g. 26 → 2026
            fy_cols[col_idx] = fy_year

    log.info(f"Found FY columns: {list(fy_cols.values())}")

    # Extract monthly rows (rows after header, before TOTAL)
    rows = []
    for i in range(header_row + 1, len(df)):
        month_val = str(df.iloc[i, 1]).strip()
        if month_val not in FY_MONTH_ORDER:
            break  # hit TOTAL or empty row

        month_num = MONTH_MAP[month_val]

        for col_idx, fy_year in fy_cols.items():
            val = df.iloc[i, col_idx]

            # Skip missing or dash values
            if pd.isna(val) or str(val).strip() in ("-", "nan", ""):
                continue

            try:
                ridership = int(round(float(val)))
            except (ValueError, TypeError):
                continue

            if ridership <= 0:
                continue

            # Convert FY + month → calendar year
            # Jul–Dec of FY N = calendar year N-1
            # Jan–Jun of FY N = calendar year N
            if month_num >= 7:
                cal_year = fy_year - 1
            else:
                cal_year = fy_year

            # Skip future months
            cal_date = date(cal_year, month_num, 1)
            if cal_date > date.today():
                continue

            rows.append({
                "date"     : cal_date,
                "value"    : ridership,
                "fy"       : fy_year,
                "month_str": month_val,
            })

    df_out = pd.DataFrame(rows)

    # Deduplicate — prefer most recent FY for overlapping months
    df_out = df_out.sort_values(["date", "fy"]).drop_duplicates("date", keep="last")
    df_out = df_out.sort_values("date").reset_index(drop=True)

    log.info(f"Parsed {len(df_out)} monthly rows")
    log.info(f"Date range: {df_out['date'].min().strftime('%b %Y')} → {df_out['date'].max().strftime('%b %Y')}")
    return df_out


# ── Normalize ─────────────────────────────────────────────────────────────────
def normalize(df: pd.DataFrame, source_url: str) -> pd.DataFrame:
    today = pd.Timestamp(date.today())
    current_month = today.to_period("M").to_timestamp()

    out_rows = []
    for _, row in df.iterrows():
        row_date = pd.Timestamp(row["date"])
        is_provisional = row_date == current_month

        out_rows.append({
            "agency_id"      : AGENCY_ID,
            "agency_name"    : AGENCY_NAME,
            "ntd_id"         : NTD_ID,
            "date"           : row["date"].strftime("%Y-%m-%d"),
            "metric"         : "monthly_boardings",
            "value"          : int(row["value"]),
            "mode"           : "CR",
            "unit"           : "boardings",
            "source"         : "sonomamarintrain.org",
            "source_url"     : source_url,
            "scraped_at"     : now_utc(),
            "is_provisional" : is_provisional,
        })

    result = pd.DataFrame(out_rows)
    log.info(f"Normalized {len(result)} rows ({result['is_provisional'].sum()} provisional)")
    return result[NORMALIZED_COLUMNS]


# ── Run ───────────────────────────────────────────────────────────────────────
def run() -> pd.DataFrame:
    log.info(f"Starting scrape: {AGENCY_NAME}")
    try:
        excel_url  = find_excel_url()
        raw_bytes  = fetch(excel_url)
        parsed     = parse(raw_bytes)
        df         = normalize(parsed, excel_url)

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
    print(df[["date", "value", "source"]].tail(6).to_string(index=False))
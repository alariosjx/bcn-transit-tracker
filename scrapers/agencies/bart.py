# scrapers/agencies/bart.py
# BART monthly exits scraper
# Sources:
#   Completed months: bart.gov monthly XLS (official final total)
#   Current month:    bart.gov daily ridership watch page (provisional)
#
# URL pattern: https://www.bart.gov/sites/default/files/YYYY-MM/Ridership_YYYYMM.xlsx
# Daily page:  https://www.bart.gov/news/articles/2025/news20250109-1
#
# Bay City News | Andres Jimenez Larios

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

import io
from datetime import date
from dateutil.relativedelta import relativedelta

import pandas as pd
from bs4 import BeautifulSoup

from scrapers._utils import (
    fetch_url, raw_path, log_scrape, now_utc, NORMALIZED_COLUMNS, log, RAW_DIR
)

AGENCY_ID   = "bart"
AGENCY_NAME = "Bay Area Rapid Transit"
NTD_ID      = "90003"
DAILY_URL   = "https://www.bart.gov/news/articles/2025/news20250109-1"
XLS_BASE    = "https://www.bart.gov/sites/default/files/{pub_month}/Ridership_{data_month}.xlsx"

EXPECTED_NTD_VARIANCE_NOTE = (
    "BART submits exits as UPT to NTD. NTD applies a ~7.5% upward adjustment. "
    "Variance up to 10% between bart.gov and NTD is expected and normal."
)


# ── Finalized-month cache ─────────────────────────────────────────────────────
def load_finalized_cache() -> dict:
    """
    Reads the most recent raw CSV for this agency and returns a dict of
    already-finalized months so build_monthly can skip re-fetching the XLS.
    Key: month Timestamp. Value: (value, source_url).
    """
    files = sorted(RAW_DIR.glob(f"{AGENCY_ID}_2*.csv"), reverse=True)
    if not files:
        return {}
    try:
        df = pd.read_csv(files[0], parse_dates=["date"])
        cache = {}
        for _, r in df[~df["is_provisional"].astype(bool)].iterrows():
            month_ts = pd.to_datetime(r["date"]).to_period("M").to_timestamp()
            cache[month_ts] = (int(r["value"]), str(r.get("source_url", "")))
        log.info(f"XLS cache: {len(cache)} finalized months from {files[0].name}")
        return cache
    except Exception as e:
        log.warning(f"Could not load finalized cache: {e}")
        return {}


# ── Monthly XLS fetch ─────────────────────────────────────────────────────────
def fetch_monthly_xls(data_month: pd.Timestamp) -> tuple[int | None, str]:
    """
    Downloads the official monthly XLS and returns the Grand Total exits.
    BART publishes files one month after the data month.
    Falls back to same-month if next-month URL 404s.
    """
    data_str   = data_month.strftime("%Y%m")
    pub_month1 = (data_month + relativedelta(months=1)).strftime("%Y-%m")
    pub_month2 = data_month.strftime("%Y-%m")

    url = None
    resp = None
    for pub in [pub_month1, pub_month2]:
        candidate = XLS_BASE.format(pub_month=pub, data_month=data_str)
        try:
            resp = fetch_url(candidate, timeout=30)
            url = candidate
            break
        except Exception:
            log.warning(f"XLS not found at {candidate}")

    if resp is None:
        log.warning(f"Monthly XLS not available for {data_month.strftime('%B %Y')}")
        return None, ""

    log.info(f"Fetched XLS: {url}")

    # Parse — use Total Trips sheet
    xl = pd.ExcelFile(io.BytesIO(resp.content))
    sheet = "Total Trips" if "Total Trips" in xl.sheet_names else xl.sheet_names[0]
    df = xl.parse(sheet, header=None)

    # Find Grand Total row — last column is the system-wide monthly total
    total = None
    for row_idx in range(len(df) - 1, max(len(df) - 10, -1), -1):
        first_cell = str(df.iloc[row_idx, 0]).strip()
        if "grand total" in first_cell.lower():
            row_vals     = df.iloc[row_idx, 1:]
            numeric_vals = pd.to_numeric(row_vals, errors="coerce").dropna()
            if not numeric_vals.empty:
                total = int(numeric_vals.iloc[-1])
                break

    if total:
        log.info(f"{data_month.strftime('%B %Y')} XLS Grand Total: {total:,}")
    else:
        log.warning(f"Could not find Grand Total in XLS for {data_month.strftime('%B %Y')}")

    return total, url


# ── Daily page fetch and parse ────────────────────────────────────────────────
def fetch_daily_page() -> pd.DataFrame:
    """
    Fetches the BART daily ridership watch page.
    Returns df with columns: date_parsed (datetime), exits (int)
    """
    resp = fetch_url(DAILY_URL)
    soup = BeautifulSoup(resp.text, "lxml")

    rows = []
    for table in soup.find_all("table"):
        for tr in table.find_all("tr"):
            tds = tr.find_all("td")
            if len(tds) != 2:
                continue
            date_str = tds[0].get_text(strip=True)
            val_str  = tds[1].get_text(strip=True).replace(",", "").replace("\xa0", "")
            if not date_str or not val_str or date_str.lower() == "date":
                continue
            if not val_str.isdigit():
                continue
            rows.append({"date_str": date_str, "exits": int(val_str)})

    df = pd.DataFrame(rows)
    df["date_parsed"] = pd.to_datetime(df["date_str"], format="%m/%d/%y", errors="coerce")
    df = df.dropna(subset=["date_parsed"])
    log.info(f"Daily page: {len(df)} rows parsed")
    return df


# ── Build monthly rows ────────────────────────────────────────────────────────
def build_monthly(daily_df: pd.DataFrame, finalized_cache: dict | None = None) -> pd.DataFrame:
    """
    For each month in the daily data:
    - Completed months: use cached finalized value (if available) or fetch official XLS Grand Total
    - Current month: sum daily rows (provisional)
    """
    today         = pd.Timestamp(date.today())
    current_month = today.to_period("M").to_timestamp()

    daily_df = daily_df.copy()
    daily_df["month"] = daily_df["date_parsed"].dt.to_period("M").dt.to_timestamp()

    out_rows = []
    for month in sorted(daily_df["month"].unique()):
        month_daily   = daily_df[daily_df["month"] == month]
        days_reported = len(month_daily)
        days_in_month = pd.Timestamp(month).days_in_month
        is_current    = month == current_month

        if is_current or days_reported < days_in_month:
            # Provisional — sum what we have
            daily_sum = int(month_daily["exits"].sum())
            if days_reported < days_in_month and not is_current:
                # Project partial month
                total = int(round(daily_sum / days_reported * days_in_month))
            else:
                total = daily_sum
            is_provisional = True
            source_url = DAILY_URL
            log.info(f"{month.strftime('%b %Y')}: provisional = {total:,} ({days_reported}/{days_in_month} days)")

        else:
            # Completed month — use cache to avoid re-fetching XLS we already have
            if finalized_cache and month in finalized_cache:
                total, source_url = finalized_cache[month]
                log.info(f"{month.strftime('%b %Y')}: cached = {total:,}")
            else:
                xls_total, xls_url = fetch_monthly_xls(month)
                if xls_total:
                    total      = xls_total
                    source_url = xls_url
                else:
                    total      = int(month_daily["exits"].sum())
                    source_url = DAILY_URL
                    log.warning(f"{month.strftime('%b %Y')}: using daily sum = {total:,}")
            is_provisional = False

        out_rows.append({
            "agency_id"      : AGENCY_ID,
            "agency_name"    : AGENCY_NAME,
            "ntd_id"         : NTD_ID,
            "date"           : month.strftime("%Y-%m-%d"),
            "metric"         : "monthly_exits",
            "value"          : total,
            "mode"           : "HR",
            "unit"           : "exits",
            "source"         : "bart.gov",
            "source_url"     : source_url,
            "scraped_at"     : now_utc(),
            "is_provisional" : is_provisional,
        })

    return pd.DataFrame(out_rows)[NORMALIZED_COLUMNS]


# ── Run ───────────────────────────────────────────────────────────────────────
def run() -> pd.DataFrame:
    log.info(f"Starting scrape: {AGENCY_NAME}")
    try:
        finalized_cache = load_finalized_cache()
        daily_df        = fetch_daily_page()
        df              = build_monthly(daily_df, finalized_cache)

        out = raw_path(AGENCY_ID)
        df.to_csv(out, index=False)
        log_scrape(AGENCY_ID, "success", len(df), notes=EXPECTED_NTD_VARIANCE_NOTE)
        log.info(f"Wrote {len(df)} rows → {out}")
        return df

    except Exception as e:
        log_scrape(AGENCY_ID, "error", 0, notes=str(e))
        log.error(f"Scrape failed: {e}")
        raise


if __name__ == "__main__":
    df = run()
    print("\n── Latest 6 months ─────────────────────────────")
    print(df.tail(6).to_string(index=False))
# export/to_excel.py
# Builds data/public/bcn_transit_master.xlsx
# Tabs: Master | Monthly Pivot | YoY | Agency Profiles | ReadMe
# Also writes per-agency clean CSVs to data/public/
# Bay City News

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timezone

from merge.config import AGENCIES, PROCESSED_DIR, PUBLIC_DIR
from scrapers._utils import log, now_utc

PUBLIC_DIR.mkdir(parents=True, exist_ok=True)
EXCEL_OUT = PUBLIC_DIR / "bcn_transit_master.xlsx"

# ── BCN palette ───────────────────────────────────────────────────────────────
BCN_RED    = "C0392B"
BCN_DARK   = "1A1A1A"
BCN_GRAY   = "F5F5F5"
BCN_BORDER = "DDDDDD"
FLAG_REVIEW      = "FFF3CD"   # amber
FLAG_INVESTIGATE = "F8D7DA"   # red


# ── Load data ─────────────────────────────────────────────────────────────────
def load_master() -> pd.DataFrame:
    path = PROCESSED_DIR / "master_monthly.csv"
    if not path.exists():
        raise FileNotFoundError(
            "master_monthly.csv not found. Run merge/build_master.py first."
        )
    df = pd.read_csv(path, parse_dates=["date"])
    log.info(f"Loaded master: {len(df)} rows")
    return df


def load_ntd() -> pd.DataFrame | None:
    path = PROCESSED_DIR / "monthly_ntd.csv"
    if not path.exists():
        log.warning("monthly_ntd.csv not found — Agency Profiles tab will be limited")
        return None
    return pd.read_csv(path, parse_dates=["date"])


# ── Style helpers ─────────────────────────────────────────────────────────────
def header_style(ws, row: int, cols: int) -> None:
    """Apply BCN header style to the first row."""
    for col in range(1, cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.font      = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        cell.fill      = PatternFill("solid", fgColor=BCN_RED)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def freeze_and_autofit(ws, freeze_cell: str = "A2") -> None:
    ws.freeze_panes = freeze_cell
    for col_cells in ws.columns:
        max_len = max(
            (len(str(c.value)) for c in col_cells if c.value is not None),
            default=8
        )
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(max_len + 2, 40)


def write_df_to_sheet(ws, df: pd.DataFrame, start_row: int = 1) -> None:
    """Write a dataframe to a worksheet, with header styling."""
    for col_idx, col_name in enumerate(df.columns, 1):
        ws.cell(row=start_row, column=col_idx, value=col_name)
    header_style(ws, start_row, len(df.columns))

    for row_idx, row in enumerate(df.itertuples(index=False), start_row + 1):
        for col_idx, val in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.value     = None if pd.isna(val) else val
            cell.font      = Font(name="Arial", size=10)
            cell.alignment = Alignment(vertical="center")
            # Alternate row shading
            if row_idx % 2 == 0:
                cell.fill = PatternFill("solid", fgColor=BCN_GRAY)


def highlight_flags(ws, flag_col_idx: int, start_row: int, n_rows: int, n_cols: int) -> None:
    """Color-code rows based on flag value."""
    for row_idx in range(start_row + 1, start_row + n_rows + 1):
        flag_val = ws.cell(row=row_idx, column=flag_col_idx).value
        if flag_val in ("REVIEW", "INVESTIGATE"):
            color = FLAG_INVESTIGATE if flag_val == "INVESTIGATE" else FLAG_REVIEW
            for col in range(1, n_cols + 1):
                ws.cell(row=row_idx, column=col).fill = PatternFill("solid", fgColor=color)


# ── Tab 1: Master ─────────────────────────────────────────────────────────────
def build_master_tab(wb, master: pd.DataFrame) -> None:
    ws = wb.create_sheet("Master")
    ws.sheet_properties.tabColor = BCN_RED

    df = master.copy()
    df["date"] = pd.to_datetime(df["date"]).dt.strftime("%Y-%m-%d")
    df["variance_pct"] = df["variance_pct"].apply(
        lambda x: f"{x:.1%}" if pd.notna(x) else ""
    )

    write_df_to_sheet(ws, df)

    # Highlight flagged rows
    flag_col = list(df.columns).index("flag") + 1
    highlight_flags(ws, flag_col, 1, len(df), len(df.columns))

    freeze_and_autofit(ws, "A2")
    log.info("Built tab: Master")


# ── Tab 2: Monthly Pivot ──────────────────────────────────────────────────────
def build_monthly_tab(wb, master: pd.DataFrame) -> None:
    ws = wb.create_sheet("Monthly")

    df = master[master["metric"].isin(["monthly_exits", "upt"])].copy()
    df["date"] = pd.to_datetime(df["date"])
    df["month_label"] = df["date"].dt.strftime("%b %Y")

    pivot = df.pivot_table(
        index   = ["agency_name"],
        columns = "month_label",
        values  = "value",
        aggfunc = "sum"
    )

    # Sort columns chronologically
    pivot = pivot.reindex(
        sorted(pivot.columns, key=lambda x: pd.to_datetime(x, format="%b %Y")),
        axis=1
    )
    pivot = pivot.reset_index()

    write_df_to_sheet(ws, pivot)
    freeze_and_autofit(ws, "B2")
    log.info("Built tab: Monthly Pivot")


# ── Tab 3: Year-over-Year ─────────────────────────────────────────────────────
def build_yoy_tab(wb, master: pd.DataFrame) -> None:
    ws = wb.create_sheet("YoY")

    df = master[master["metric"].isin(["monthly_exits", "upt"])].copy()
    df["date"]     = pd.to_datetime(df["date"])
    df["year"]     = df["date"].dt.year
    df["month"]    = df["date"].dt.month

    # Join to same month prior year
    df_yoy = df.merge(
        df[["agency_id", "year", "month", "value"]].rename(
            columns={"value": "value_prior", "year": "year_prior"}
        ).assign(year=lambda x: x["year_prior"] + 1),
        on=["agency_id", "year", "month"],
        how="left"
    )
    df_yoy["yoy_change"]  = df_yoy["value"] - df_yoy["value_prior"]
    df_yoy["yoy_pct"]     = (df_yoy["yoy_change"] / df_yoy["value_prior"]).round(4)
    df_yoy["yoy_pct_fmt"] = df_yoy["yoy_pct"].apply(
        lambda x: f"{x:.1%}" if pd.notna(x) else ""
    )

    out = df_yoy[[
        "agency_name", "date", "value", "value_prior",
        "yoy_change", "yoy_pct_fmt", "source", "flag"
    ]].copy()
    out["date"] = out["date"].dt.strftime("%Y-%m-%d")
    out = out.sort_values(["agency_name", "date"])

    write_df_to_sheet(ws, out)
    freeze_and_autofit(ws, "A2")
    log.info("Built tab: YoY")


# ── Tab 4: Agency Profiles ────────────────────────────────────────────────────
def build_profiles_tab(wb, master: pd.DataFrame, ntd: pd.DataFrame | None) -> None:
    """
    One row per agency — latest available snapshot of key metrics.
    Ridership from master; operating data from NTD if available.
    This is the 'at a glance' reference sheet for reporters.
    """
    ws = wb.create_sheet("Agency Profiles")
    ws.sheet_properties.tabColor = "2C3E50"

    profile_rows = []

    for agency_id, config in AGENCIES.items():
        agency_master = master[master["agency_id"] == agency_id].copy()
        if agency_master.empty:
            continue

        agency_master["date"] = pd.to_datetime(agency_master["date"])
        latest = agency_master.sort_values("date").iloc[-1]
        latest_date  = latest["date"]
        latest_value = latest["value"]
        latest_source = latest["source"]

        # 2019 baseline for recovery %
        baseline_2019 = agency_master[
            (agency_master["date"].dt.year == 2019) &
            (agency_master["date"].dt.month == latest_date.month)
        ]["value"]
        recovery_pct = (
            round(latest_value / baseline_2019.iloc[0] * 100, 1)
            if not baseline_2019.empty else None
        )

        # 2025 annual (for context)
        annual_2025 = agency_master[
            agency_master["date"].dt.year == 2025
        ]["value"].sum()

        # NTD financial data (most recent year available)
        ntd_metrics = {}
        if ntd is not None:
            ntd_agency = ntd[ntd["ntd_id"].astype(str) == config["ntd_id"]]
            if not ntd_agency.empty:
                latest_ntd = ntd_agency.sort_values("date").iloc[-1]
                ntd_metrics = {
                    "NTD Latest Month"  : latest_ntd["date"].strftime("%b %Y"),
                    "NTD UPT"           : int(latest_ntd.get("upt", 0) or 0),
                    "NTD VRM"           : int(latest_ntd.get("vrm", 0) or 0),
                    "NTD VRH"           : int(latest_ntd.get("vrh", 0) or 0),
                }

        row = {
            "Agency"              : config["agency_name"],
            "NTD ID"              : config["ntd_id"],
            "Latest Month"        : latest_date.strftime("%B %Y"),
            "Latest Ridership"    : int(latest_value),
            "Source"              : latest_source,
            "Recovery vs 2019 (%)": recovery_pct,
            "2025 Annual Total"   : int(annual_2025) if annual_2025 > 0 else None,
            **ntd_metrics,
            "Notes"               : config.get("notes", ""),
        }
        profile_rows.append(row)

    profiles = pd.DataFrame(profile_rows)
    write_df_to_sheet(ws, profiles)
    freeze_and_autofit(ws, "A2")
    log.info("Built tab: Agency Profiles")


# ── Tab 5: ReadMe ─────────────────────────────────────────────────────────────
def build_readme_tab(wb) -> None:
    ws = wb.create_sheet("ReadMe")

    lines = [
        ["BCN Transit Tracker — Data Dictionary"],
        ["Bay City News | baycitynews.com"],
        [""],
        [f"Last updated: {datetime.now(timezone.utc).strftime('%B %d, %Y at %H:%M UTC')}"],
        [""],
        ["TABS"],
        ["Master",         "Full long-form table. One row per agency per month. All sources, flags, and timestamps."],
        ["Monthly Pivot",  "Wide format — one row per agency, one column per month. Good for quick scanning."],
        ["YoY",            "Year-over-year comparison. Same month, prior year vs current."],
        ["Agency Profiles","Latest snapshot per agency. Ridership + NTD financial data. Good for reporter reference."],
        ["ReadMe",         "This tab."],
        [""],
        ["FLAG COLUMN (Master tab)"],
        ["(blank)",        "Variance between agency-direct and NTD is under 5%. Normal."],
        ["REVIEW",         "Variance 5–10%. Worth investigating before publishing."],
        ["INVESTIGATE",    "Variance over 10%. Do not publish without understanding the difference."],
        [""],
        ["SOURCES"],
        ["bart.gov",       "BART daily paid exits. Updated Monday–Friday. More timely than NTD."],
        ["NTD",            "National Transit Database monthly module. ~2 month lag. Adjusted figures."],
        ["SFMTA APC",      "SF Muni automatic passenger counters. Published monthly on sfmta.com."],
        [""],
        ["NOTES"],
        ["- BART exits run ~7.5% lower than NTD adjusted UPT. This is expected and documented."],
        ["- NTD applies statistical adjustments that can differ from agency-reported figures."],
        ["- Provisional rows (is_provisional=True) are partial months or pre-NTD-release figures."],
        ["- Do not use provisional figures in published stories without noting they may be revised."],
        [""],
        ["CONTACT"],
        ["Data team: Andres Jimenez Larios | ajlarios@baycitynews.com"],
        ["GitHub: https://github.com/baycitnews/bcn-transit-tracker"],
    ]

    for row_idx, line in enumerate(lines, 1):
        for col_idx, val in enumerate(line, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = Font(name="Arial", size=10)
            if row_idx == 1:
                cell.font = Font(name="Arial", size=14, bold=True, color=BCN_RED)
            elif col_idx == 1 and row_idx > 5 and val and val.isupper():
                cell.font = Font(name="Arial", size=10, bold=True)

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 80
    log.info("Built tab: ReadMe")


# ── Per-agency public CSVs ────────────────────────────────────────────────────
def write_public_csvs(master: pd.DataFrame) -> None:
    """
    Writes one clean CSV per agency to data/public/ — ready for Datawrapper/Flourish.
    Only includes confirmed (non-provisional) rows.
    """
    for agency_id in master["agency_id"].unique():
        df = master[
            (master["agency_id"] == agency_id) &
            (~master["is_provisional"])
        ].copy()
        df["date"] = pd.to_datetime(df["date"]).dt.strftime("%Y-%m-%d")
        out = PUBLIC_DIR / f"{agency_id}_monthly.csv"
        df.to_csv(out, index=False)
        log.info(f"Wrote public CSV: {out} ({len(df)} rows)")


# ── Main ──────────────────────────────────────────────────────────────────────
def run():
    log.info("── to_excel.py starting ──")

    master = load_master()
    ntd    = load_ntd()

    # Create workbook — remove default sheet
    from openpyxl import Workbook
    wb = Workbook()
    wb.remove(wb.active)

    build_master_tab(wb, master)
    build_monthly_tab(wb, master)
    build_yoy_tab(wb, master)
    build_profiles_tab(wb, master, ntd)
    build_readme_tab(wb)

    wb.save(EXCEL_OUT)
    log.info(f"Saved Excel: {EXCEL_OUT}")

    write_public_csvs(master)
    log.info("── to_excel.py complete ──")


if __name__ == "__main__":
    run()

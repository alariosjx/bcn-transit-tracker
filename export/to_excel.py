# export/to_excel.py
# Builds data/public/bcn_transit_master.xlsx — reporter-facing deliverable
# Tabs: Summary | Master | Monthly Pivot | YoY | ReadMe
# Also writes Datawrapper-ready CSVs to data/public/datawrapper/
#
# Bay City News | Andres Jimenez Larios

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime, timezone

from merge.config import PROCESSED_DIR, PUBLIC_DIR
from scrapers._utils import log, now_utc

PUBLIC_DIR.mkdir(parents=True, exist_ok=True)
DW_DIR = PUBLIC_DIR / "datawrapper"
DW_DIR.mkdir(parents=True, exist_ok=True)

EXCEL_OUT = PUBLIC_DIR / "bcn_transit_master.xlsx"

BCN_RED   = "C0392B"
BCN_GRAY  = "F5F5F5"
FLAG_HIGH = "D4EDDA"
FLAG_LOW  = "F8D7DA"
FLAG_WARN = "FFF3CD"


# ── Style helpers ─────────────────────────────────────────────────────────────
def header_row(ws, row: int, ncols: int) -> None:
    for col in range(1, ncols + 1):
        cell = ws.cell(row=row, column=col)
        cell.font      = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        cell.fill      = PatternFill("solid", fgColor=BCN_RED)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def freeze_autofit(ws, freeze: str = "A2") -> None:
    ws.freeze_panes = freeze
    for col_cells in ws.columns:
        max_len = max((len(str(c.value)) for c in col_cells if c.value), default=8)
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(max_len + 2, 45)


def write_df(ws, df: pd.DataFrame, start_row: int = 1) -> None:
    for ci, col in enumerate(df.columns, 1):
        ws.cell(row=start_row, column=ci, value=col)
    header_row(ws, start_row, len(df.columns))
    for ri, row in enumerate(df.itertuples(index=False), start_row + 1):
        for ci, val in enumerate(row, 1):
            cell = ws.cell(row=ri, column=ci)
            cell.value     = None if (isinstance(val, float) and pd.isna(val)) else val
            cell.font      = Font(name="Arial", size=10)
            cell.alignment = Alignment(vertical="center")
            if ri % 2 == 0:
                cell.fill = PatternFill("solid", fgColor=BCN_GRAY)


# ── Load data ─────────────────────────────────────────────────────────────────
def load_master() -> pd.DataFrame:
    path = PROCESSED_DIR / "master_monthly.csv"
    if not path.exists():
        raise FileNotFoundError("master_monthly.csv not found. Run merge/build_master.py first.")
    df = pd.read_csv(path, parse_dates=["date"])
    log.info(f"Loaded master: {len(df)} rows")
    return df


# ── Tab 1: Summary ────────────────────────────────────────────────────────────
def tab_summary(wb, master: pd.DataFrame) -> None:
    """
    Plain-language summary tab. Reporter-first.
    Flags: YoY change, recovery vs 2019, month-over-month swings.
    """
    ws = wb.create_sheet("Summary")
    ws.sheet_properties.tabColor = "2C3E50"

    df = master.copy()
    df["year"]  = df["date"].dt.year
    df["month"] = df["date"].dt.month

    rows = []

    for agency_id in df["agency_id"].unique():
        adf = df[df["agency_id"] == agency_id].sort_values("date")
        if len(adf) < 2:
            continue

        agency_name = adf["agency_name"].iloc[0]
        latest      = adf.iloc[-1]
        prev        = adf.iloc[-2]

        same_ly    = adf[(adf["year"] == latest["year"] - 1) & (adf["month"] == latest["month"])]
        base_2019  = adf[(adf["year"] == 2019) & (adf["month"] == latest["month"])]

        # YoY
        yoy_pct, yoy_text = None, "No prior year data"
        if not same_ly.empty and same_ly["value"].iloc[0] > 0:
            prior    = same_ly["value"].iloc[0]
            yoy_pct  = (latest["value"] - prior) / prior
            arrow    = "▲" if yoy_pct >= 0 else "▼"
            yoy_text = f"{arrow} {abs(yoy_pct):.1%} vs {latest['date'].strftime('%b')} {latest['year']-1} ({int(prior):,} → {int(latest['value']):,})"

        # MoM
        mom_pct, mom_text = None, "No prior month"
        if prev["value"] > 0:
            mom_pct  = (latest["value"] - prev["value"]) / prev["value"]
            arrow    = "▲" if mom_pct >= 0 else "▼"
            mom_text = f"{arrow} {abs(mom_pct):.1%} vs {prev['date'].strftime('%b %Y')} ({int(prev['value']):,} → {int(latest['value']):,})"

        # Recovery vs 2019
        rec_pct, rec_text = None, "No 2019 baseline"
        if not base_2019.empty and base_2019["value"].iloc[0] > 0:
            base    = base_2019["value"].iloc[0]
            rec_pct = latest["value"] / base
            rec_text = f"{rec_pct:.1%} of {latest['date'].strftime('%b')} 2019 ({int(base):,} baseline)"

        # Flags
        flags = []
        if yoy_pct is not None:
            if yoy_pct >= 0.10:
                flags.append(f"Strong growth year-over-year (+{yoy_pct:.1%})")
            elif yoy_pct <= -0.10:
                flags.append(f"Notable decline year-over-year ({yoy_pct:.1%})")
        if mom_pct is not None and abs(mom_pct) >= 0.05:
            flags.append(f"Large month-over-month swing ({mom_pct:+.1%})")
        if str(latest.get("flag", "")) in ("REVIEW", "INVESTIGATE"):
            flags.append(f"⚠ Data flag: {latest['flag']} — check agency vs NTD figures")
        if latest.get("is_provisional"):
            flags.append("Provisional — partial month, may be revised")

        rows.append({
            "Agency"           : agency_name,
            "Latest Month"     : latest["date"].strftime("%B %Y"),
            "Ridership"        : f"{int(latest['value']):,}",
            "Source"           : latest["source"],
            "YoY Change"       : yoy_text,
            "Month-over-Month" : mom_text,
            "Recovery vs 2019" : rec_text,
            "Noteworthy"       : " | ".join(flags) if flags else "Within normal range",
        })

    summary_df = pd.DataFrame(rows)
    write_df(ws, summary_df)

    # Color rows by flag
    for ri in range(2, len(summary_df) + 2):
        note = ws.cell(row=ri, column=len(summary_df.columns)).value or ""
        if "Strong growth" in note:
            color = FLAG_HIGH
        elif "Notable decline" in note or "INVESTIGATE" in note:
            color = FLAG_LOW
        elif "Provisional" in note or "REVIEW" in note or "Large month" in note:
            color = FLAG_WARN
        else:
            continue
        for ci in range(1, len(summary_df.columns) + 1):
            ws.cell(row=ri, column=ci).fill = PatternFill("solid", fgColor=color)

    freeze_autofit(ws, "A2")
    log.info("Built tab: Summary")


# ── Tab 2: Master ─────────────────────────────────────────────────────────────
def tab_master(wb, master: pd.DataFrame) -> None:
    ws = wb.create_sheet("Master")
    ws.sheet_properties.tabColor = BCN_RED

    df = master.copy()
    df["date"]         = df["date"].dt.strftime("%Y-%m-%d")
    df["variance_pct"] = df["variance_pct"].apply(lambda x: f"{x:.1%}" if pd.notna(x) else "")

    write_df(ws, df)
    freeze_autofit(ws, "A2")
    log.info("Built tab: Master")


# ── Tab 3: Monthly Pivot ──────────────────────────────────────────────────────
def tab_monthly_pivot(wb, master: pd.DataFrame) -> None:
    ws = wb.create_sheet("Monthly Pivot")

    df = master.copy()
    df["month_label"] = df["date"].dt.strftime("%b %Y")

    pivot = df.pivot_table(
        index="agency_name", columns="month_label", values="value", aggfunc="sum"
    )
    pivot = pivot.reindex(
        sorted(pivot.columns, key=lambda x: pd.to_datetime(x, format="%b %Y"), reverse=True), axis=1
    ).reset_index()

    write_df(ws, pivot)
    freeze_autofit(ws, "B2")
    log.info("Built tab: Monthly Pivot")


# ── Tab 4: YoY ────────────────────────────────────────────────────────────────
def tab_yoy(wb, master: pd.DataFrame) -> None:
    ws = wb.create_sheet("Year-over-Year")

    df = master.copy()
    df["year"]  = df["date"].dt.year
    df["month"] = df["date"].dt.month

    yoy = df.merge(
        df[["agency_id", "year", "month", "value"]].rename(
            columns={"value": "value_prior", "year": "prior_year"}
        ).assign(year=lambda x: x["prior_year"] + 1),
        on=["agency_id", "year", "month"], how="left"
    )
    yoy["yoy_change"]  = yoy["value"] - yoy["value_prior"]
    yoy["yoy_pct"]     = (yoy["yoy_change"] / yoy["value_prior"]).round(4)
    yoy["yoy_pct_fmt"] = yoy["yoy_pct"].apply(lambda x: f"{x:+.1%}" if pd.notna(x) else "")

    out = yoy[["agency_name", "date", "value", "value_prior", "yoy_change", "yoy_pct_fmt", "source", "flag"]].copy()
    out["date"] = out["date"].dt.strftime("%Y-%m-%d")
    write_df(ws, out.sort_values(["agency_name", "date"], ascending=[True, False]))

    for ri, row in enumerate(yoy.itertuples(index=False), 2):
        if pd.notna(row.yoy_pct):
            if row.yoy_pct >= 0.05:
                color = FLAG_HIGH
            elif row.yoy_pct <= -0.05:
                color = FLAG_LOW
            else:
                continue
            for ci in range(1, len(out.columns) + 1):
                ws.cell(row=ri, column=ci).fill = PatternFill("solid", fgColor=color)

    freeze_autofit(ws, "A2")
    log.info("Built tab: Year-over-Year")


# ── Tab 5: ReadMe ─────────────────────────────────────────────────────────────
def tab_readme(wb) -> None:
    ws = wb.create_sheet("ReadMe")
    ts = datetime.now(timezone.utc).strftime("%B %d, %Y at %H:%M UTC")

    lines = [
        ["BCN Transit Tracker — Bay City News"],
        [f"Last updated: {ts}"],
        [""],
        ["START HERE: Open the Summary tab. It tells you what's notable in plain language."],
        [""],
        ["TABS"],
        ["Summary",         "Plain-language flags. Green = strong growth. Red = decline. Start here."],
        ["Master",          "Full dataset — every month, every agency. All sources and QA flags."],
        ["Monthly Pivot",   "Wide format — one row per agency, columns are months. Good for quick scans."],
        ["Year-over-Year",  "% change vs same month last year. Green = up >5%, Red = down >5%."],
        ["ReadMe",          "This tab."],
        [""],
        ["DATA FLAGS"],
        ["(blank)",         "Variance between agency and NTD is under 5% — normal."],
        ["REVIEW",          "5–10% variance. Worth a look before publishing."],
        ["INVESTIGATE",     "Over 10% variance. Do not publish without understanding the difference."],
        ["is_provisional",  "Partial month estimate. Numbers will change when month completes."],
        [""],
        ["SOURCES"],
        ["bart.gov",        "BART daily paid exits. Scraped automatically each weekday at 6 AM PT."],
        ["NTD Monthly",     "National Transit Database. ~2 month lag. Adjusted and finalized."],
        [""],
        ["CONTACT"],
        ["Andres Jimenez Larios | Bay City News"],
        ["GitHub: https://github.com/alariosjx/bcn-transit-tracker"],
    ]

    for ri, line in enumerate(lines, 1):
        for ci, val in enumerate(line, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.font = Font(name="Arial", size=10)
            if ri == 1:
                cell.font = Font(name="Arial", size=14, bold=True, color=BCN_RED)
            elif ri == 2:
                cell.font = Font(name="Arial", size=10, color="888888")
            elif ri == 4:
                cell.font = Font(name="Arial", size=10, bold=True)
            elif ci == 1 and val and val.isupper():
                cell.font = Font(name="Arial", size=10, bold=True)

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 80
    log.info("Built tab: ReadMe")


# ── Datawrapper CSVs ──────────────────────────────────────────────────────────
def write_datawrapper_csvs(master: pd.DataFrame) -> None:
    df = master[~master["is_provisional"]].copy()

    # 1. Time series — monthly ridership (long format)
    ts = df[["date", "agency_name", "value", "source"]].copy()
    ts["date"] = ts["date"].dt.strftime("%Y-%m-%d")
    ts.to_csv(DW_DIR / "bart_monthly_timeseries.csv", index=False)

    # 2. YoY wide — good for grouped bar charts
    df["year"]        = df["date"].dt.year
    df["month"]       = df["date"].dt.month
    df["month_label"] = df["date"].dt.strftime("%b")

    yoy_wide = df.pivot_table(
        index=["agency_name", "month", "month_label"],
        columns="year", values="value", aggfunc="sum"
    ).reset_index()
    yoy_wide.columns = [str(c) for c in yoy_wide.columns]
    yoy_wide.sort_values(["agency_name", "month"]).to_csv(
        DW_DIR / "bart_yoy_comparison.csv", index=False
    )

    # 3. Recovery tracker — % of 2019
    baseline = df[df["year"] == 2019][["agency_name", "month", "value"]].rename(
        columns={"value": "baseline_2019"}
    )
    rec = df[df["year"] >= 2020].merge(baseline, on=["agency_name", "month"], how="left")
    rec["recovery_pct"] = (rec["value"] / rec["baseline_2019"] * 100).round(1)
    rec["date"]         = rec["date"].dt.strftime("%Y-%m-%d")
    rec[["agency_name", "date", "value", "baseline_2019", "recovery_pct"]].to_csv(
        DW_DIR / "bart_recovery_tracker.csv", index=False
    )

    log.info(f"Wrote 3 Datawrapper CSVs → {DW_DIR}")


# ── Main ──────────────────────────────────────────────────────────────────────
def run():
    log.info("── to_excel.py starting ──")
    master = load_master()

    wb = Workbook()
    wb.remove(wb.active)

    tab_summary(wb, master)
    tab_master(wb, master)
    tab_monthly_pivot(wb, master)
    tab_yoy(wb, master)
    tab_readme(wb)

    wb.save(EXCEL_OUT)
    log.info(f"Saved Excel: {EXCEL_OUT}")

    write_datawrapper_csvs(master)

    # Per-agency public CSVs
    for agency_id in master["agency_id"].unique():
        adf = master[(master["agency_id"] == agency_id) & (~master["is_provisional"])].copy()
        adf["date"] = adf["date"].dt.strftime("%Y-%m-%d")
        adf.to_csv(PUBLIC_DIR / f"{agency_id}_monthly.csv", index=False)
        log.info(f"Wrote: {agency_id}_monthly.csv")

    log.info("── to_excel.py complete ──")


if __name__ == "__main__":
    run()

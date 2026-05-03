#!/usr/bin/env python3
# tests/test_datawrapper_csvs.py
# Validates Datawrapper CSV outputs against the current master dataset.
# No fake data injection — tests run against real production files.
#
# Usage: python tests/test_datawrapper_csvs.py
# Bay City News | Andres Jimenez Larios

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
from merge.config import PROCESSED_DIR, PUBLIC_DIR
from export.to_excel import write_datawrapper_csvs

DW_DIR      = PUBLIC_DIR / "datawrapper"
MASTER_PATH = PROCESSED_DIR / "master_monthly.csv"

PASS = "✅ PASS"
FAIL = "❌ FAIL"

def check(label: str, condition: bool) -> bool:
    print(f"  {PASS if condition else FAIL}  {label}")
    return condition


def run_tests():
    print("\n── BCN Transit Tracker — Datawrapper CSV Tests ─────────────────────")
    print("Running export on current master...\n")

    master = pd.read_csv(MASTER_PATH, parse_dates=["date"])

    try:
        write_datawrapper_csvs(master)
    except Exception as e:
        print(f"  {FAIL}  Export crashed: {e}")
        return

    results = []

    # ── BART Timeseries ───────────────────────────────────────────────────────
    print("── BART Timeseries CSV ──────────────────────────────────────────────")
    ts = pd.read_csv(DW_DIR / "bart_monthly_timeseries.csv", parse_dates=["date"])
    results.append(check("BART only — no other agencies", ts["agency_name"].nunique() == 1))
    results.append(check("Starts from 2002", ts["date"].min().year == 2002))
    results.append(check("Most recent month is present", ts["date"].max().year >= 2025))
    results.append(check("No missing values in value column", ts["value"].notna().all()))

    # ── BART YoY ─────────────────────────────────────────────────────────────
    print("\n── BART YoY Comparison CSV ──────────────────────────────────────────")
    yoy = pd.read_csv(DW_DIR / "bart_yoy_comparison.csv")
    results.append(check("Has 12 rows (Jan–Dec)", len(yoy) == 12))
    results.append(check("Has month_label column", "month_label" in yoy.columns))
    results.append(check("Has prior year column", any(str(y) in yoy.columns for y in range(2020, 2030))))
    results.append(check("Has current year column", any(str(y) in yoy.columns for y in range(2024, 2030))))

    # ── BART Recovery ─────────────────────────────────────────────────────────
    print("\n── BART Recovery Tracker CSV ────────────────────────────────────────")
    rec_bart = pd.read_csv(DW_DIR / "bart_recovery_tracker.csv", parse_dates=["date"])
    results.append(check("BART only", rec_bart["agency_name"].nunique() == 1))
    results.append(check("Starts from 2020", rec_bart["date"].min().year == 2020))
    results.append(check("No negative recovery values", (rec_bart["recovery_pct"] >= 0).all()))
    results.append(check("Has baseline_2019 column", "baseline_2019" in rec_bart.columns))

    # ── Muni Timeseries ───────────────────────────────────────────────────────
    print("\n── Muni Timeseries CSV ──────────────────────────────────────────────")
    ts_muni = pd.read_csv(DW_DIR / "muni_monthly_timeseries.csv", parse_dates=["date"])
    results.append(check("Muni only — no other agencies", ts_muni["agency_name"].nunique() == 1))
    results.append(check("Starts from 2002 (NTD historical)", ts_muni["date"].min().year <= 2002))
    results.append(check("Most recent month is present", ts_muni["date"].max().year >= 2025))
    results.append(check("No missing values in value column", ts_muni["value"].notna().all()))

    # ── Muni YoY ─────────────────────────────────────────────────────────────
    print("\n── Muni YoY Comparison CSV ──────────────────────────────────────────")
    yoy_muni = pd.read_csv(DW_DIR / "muni_yoy_comparison.csv")
    results.append(check("Has 12 rows (Jan–Dec)", len(yoy_muni) == 12))
    results.append(check("Has month_label column", "month_label" in yoy_muni.columns))

    # ── Muni Recovery ─────────────────────────────────────────────────────────
    print("\n── Muni Recovery Tracker CSV ────────────────────────────────────────")
    rec_muni = pd.read_csv(DW_DIR / "muni_recovery_tracker.csv", parse_dates=["date"])
    results.append(check("Muni only", rec_muni["agency_name"].nunique() == 1))
    results.append(check("Starts from 2020", rec_muni["date"].min().year == 2020))
    results.append(check("No negative recovery values", (rec_muni["recovery_pct"] >= 0).all()))
    mar_2026 = rec_muni[rec_muni["date"] == "2026-03-01"]
    results.append(check("Mar 2026 recovery is ~84%",
        not mar_2026.empty and abs(mar_2026["recovery_pct"].iloc[0] - 84.4) < 2
    ))

    # ── Summary ───────────────────────────────────────────────────────────────
    passed = sum(results)
    total  = len(results)
    print(f"\n── Results: {passed}/{total} passed ─────────────────────────────────────")
    if passed == total:
        print("  All tests passed.\n")
    else:
        print(f"  {total - passed} test(s) failed. Review output above.\n")


if __name__ == "__main__":
    run_tests()
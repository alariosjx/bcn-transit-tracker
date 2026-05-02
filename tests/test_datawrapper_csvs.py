#!/usr/bin/env python3
# tests/test_datawrapper_csvs.py
# Injects a fake May 2026 row into the master dataset, runs the
# Datawrapper CSV export, validates outputs, then cleans up.
#
# Usage: python tests/test_datawrapper_csvs.py
# Bay City News | Andres Jimenez Larios

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import shutil
import pandas as pd
from merge.config import PROCESSED_DIR, PUBLIC_DIR
from export.to_excel import write_datawrapper_csvs
from scrapers._utils import log

DW_DIR      = PUBLIC_DIR / "datawrapper"
MASTER_PATH = PROCESSED_DIR / "master_monthly.csv"
BACKUP_PATH = PROCESSED_DIR / "master_monthly_backup.csv"

FAKE_MAY = {
    "agency_id"      : "bart",
    "agency_name"    : "Bay Area Rapid Transit",
    "ntd_id"         : "90003",
    "date"           : "2026-05-01",
    "metric"         : "monthly_exits",
    "value"          : 5_100_000,   # fake number
    "mode"           : "HR",
    "unit"           : "exits",
    "source"         : "bart.gov",
    "source_url"     : "TEST",
    "scraped_at"     : "2026-05-31T00:00:00Z",
    "is_provisional" : False,
    "ntd_value"      : None,
    "variance_pct"   : None,
    "flag"           : "",
    "last_updated"   : "2026-05-31T00:00:00Z",
}

PASS = "✅ PASS"
FAIL = "❌ FAIL"

def check(label: str, condition: bool) -> bool:
    print(f"  {PASS if condition else FAIL}  {label}")
    return condition


def run_tests():
    print("\n── BCN Transit Tracker — Datawrapper CSV Tests ─────────────────────")
    print("Injecting fake May 2026 row (5,100,000 exits)...\n")

    # ── Setup: backup master and inject fake row ──────────────────────────────
    shutil.copy(MASTER_PATH, BACKUP_PATH)

    master = pd.read_csv(MASTER_PATH, parse_dates=["date"])
    original_len = len(master)

    fake_row = pd.DataFrame([FAKE_MAY])
    fake_row["date"] = pd.to_datetime(fake_row["date"])
    master_with_fake = pd.concat([master, fake_row], ignore_index=True)
    master_with_fake.to_csv(MASTER_PATH, index=False)

    # ── Run the export ────────────────────────────────────────────────────────
    try:
        master_reloaded = pd.read_csv(MASTER_PATH, parse_dates=["date"])
        write_datawrapper_csvs(master_reloaded)
    except Exception as e:
        print(f"  {FAIL}  Export crashed: {e}")
        _restore(MASTER_PATH, BACKUP_PATH)
        return

    # ── Run checks ───────────────────────────────────────────────────────────
    results = []
    print("── Timeseries CSV ───────────────────────────────────────────────────")
    ts = pd.read_csv(DW_DIR / "bart_monthly_timeseries.csv", parse_dates=["date"])
    results.append(check(f"Has {original_len + 1} rows (original + fake May)", len(ts) == original_len + 1))
    results.append(check("Starts from 2002", ts["date"].min().year == 2002))
    results.append(check("Contains fake May 2026", (ts["date"] == "2026-05-01").any()))
    results.append(check("No missing values in value column", ts["value"].notna().all()))

    print("\n── YoY Comparison CSV ───────────────────────────────────────────────")
    yoy = pd.read_csv(DW_DIR / "bart_yoy_comparison.csv")
    results.append(check("Has 12 rows (Jan–Dec)", len(yoy) == 12))
    results.append(check("Has month_label column", "month_label" in yoy.columns))
    results.append(check("Has 2025 column", "2025" in yoy.columns))
    results.append(check("Has 2026 column", "2026" in yoy.columns))
    may_row = yoy[yoy["month_label"] == "May"]
    results.append(check("May 2026 shows 5,100,000", not may_row.empty and may_row["2026"].iloc[0] == 5_100_000))
results.append(check("Jun–Dec 2026 are blank", yoy[yoy["month_label"].isin(["Jun","Jul","Aug","Sep","Oct","Nov","Dec"])]["2026"].isna().all()))
    print("\n── Recovery Tracker CSV ─────────────────────────────────────────────")
    rec = pd.read_csv(DW_DIR / "bart_recovery_tracker.csv", parse_dates=["date"])
    results.append(check("Starts from 2020", rec["date"].min().year == 2020))
    results.append(check("Contains fake May 2026", (rec["date"] == "2026-05-01").any()))
    may_rec = rec[rec["date"] == "2026-05-01"]
    results.append(check("May 2026 recovery_pct is calculated", not may_rec.empty and pd.notna(may_rec["recovery_pct"].iloc[0])))
    results.append(check("No negative recovery values", (rec["recovery_pct"] >= 0).all()))

    # ── Summary ───────────────────────────────────────────────────────────────
    passed = sum(results)
    total  = len(results)
    print(f"\n── Results: {passed}/{total} passed ─────────────────────────────────────")
    if passed == total:
        print("  All tests passed. Pipeline handles new months correctly.\n")
    else:
        print(f"  {total - passed} test(s) failed. Review output above.\n")

    # ── Cleanup: restore original master ──────────────────────────────────────
    _restore(MASTER_PATH, BACKUP_PATH)
    print("  Master CSV restored to original state.\n")


def _restore(master_path: Path, backup_path: Path):
    shutil.copy(backup_path, master_path)
    backup_path.unlink()


if __name__ == "__main__":
    run_tests()
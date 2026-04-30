# run_all.py
# Runs the full BCN Transit Tracker pipeline:
#   1. Scrape all configured agencies
#   2. Merge agency-direct + NTD, flag discrepancies
#   3. Export Excel + public CSVs
#
# Usage:
#   python run_all.py              # full pipeline
#   python run_all.py --scrape     # scrape only
#   python run_all.py --merge      # merge only
#   python run_all.py --export     # export only
#
# Bay City News

import argparse
import importlib
import sys

from merge.config import AGENCIES
from scrapers._utils import log


def scrape_all():
    log.info("── Scraping all agencies ──────────────────────────")
    for agency_id, config in AGENCIES.items():
        if config["primary_source"] != "agency_direct" or not config["scraper"]:
            log.info(f"Skipping {agency_id} — NTD only, no direct scraper")
            continue
        try:
            module = importlib.import_module(config["scraper"])
            module.run()
        except Exception as e:
            log.error(f"Scrape failed for {agency_id}: {e}")
            # Don't stop the pipeline on one scraper failure
            continue


def merge():
    log.info("── Building master dataset ────────────────────────")
    from merge.build_master import run
    run()


def export():
    log.info("── Exporting Excel + CSVs ─────────────────────────")
    from export.to_excel import run
    run()


def main():
    parser = argparse.ArgumentParser(description="BCN Transit Tracker pipeline")
    parser.add_argument("--scrape", action="store_true")
    parser.add_argument("--merge",  action="store_true")
    parser.add_argument("--export", action="store_true")
    args = parser.parse_args()

    # If no flags, run everything
    run_all = not any([args.scrape, args.merge, args.export])

    if run_all or args.scrape:
        scrape_all()
    if run_all or args.merge:
        merge()
    if run_all or args.export:
        export()

    log.info("── Pipeline complete ──────────────────────────────")


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
from pathlib import Path
import shutil

from config import INPUT_DIR, OUTPUT_DIR, OUTPUT_FILE
from report_builder import build_report
from transaction_extractor import (
    extract_all_transactions,
    summarize_transactions,
    summarize_transactions_by_portfolio,
)
from utils import iter_input_files
from valuation_extractor import extract_all_valuations


def is_valuation_file(path: Path) -> bool:
    name = path.name.lower()
    return "valuation" in name or "估值表" in path.name


def seed_sample_data_if_requested():
    """Copy synthetic sample files to input/ when --sample is used by CLI helper."""
    sample_dir = Path(__file__).resolve().parent / "sample_data"
    INPUT_DIR.mkdir(exist_ok=True)
    if not sample_dir.exists():
        raise FileNotFoundError("sample_data directory is missing")
    for file_path in sample_dir.glob("*.xlsx"):
        shutil.copy2(file_path, INPUT_DIR / file_path.name)


def run_pipeline():
    INPUT_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)
    all_files = list(iter_input_files(INPUT_DIR))
    valuation_files = [path for path in all_files if is_valuation_file(path)]
    transaction_files = [path for path in all_files if path not in valuation_files]

    print(f"Found files: {len(all_files)}")
    print(f"Valuation files: {len(valuation_files)}")
    print(f"Transaction files: {len(transaction_files)}")

    valuation_summary, underlying_rows = extract_all_valuations(valuation_files)
    txn_details = extract_all_transactions(transaction_files)
    txn_summary = summarize_transactions(txn_details)
    txn_summary_by_portfolio = summarize_transactions_by_portfolio(txn_details)

    build_report(
        OUTPUT_FILE,
        valuation_summary,
        underlying_rows,
        txn_details,
        txn_summary,
        txn_summary_by_portfolio,
    )
    print(f"Done: {OUTPUT_FILE}")
    return OUTPUT_FILE


def main():
    import argparse
    parser = argparse.ArgumentParser(description="Generate a monthly fund-of-funds reporting workbook.")
    parser.add_argument("--sample", action="store_true", help="Copy synthetic sample files into input/ before running")
    args = parser.parse_args()
    if args.sample:
        seed_sample_data_if_requested()
    run_pipeline()


if __name__ == "__main__":
    main()

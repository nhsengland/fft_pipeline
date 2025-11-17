"""
FFT Pipeline CLI Entry Point
--------------------------

This module provides a command-line interface for the FFT pipeline.
"""

import argparse
import logging
import sys
from pathlib import Path
from typing import Optional

from .pipeline import run_full_pipeline

def parse_arguments() -> argparse.Namespace:
    """Parse command line arguments for the FFT pipeline."""
    parser = argparse.ArgumentParser(
        description="Run the full FFT pipeline process",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )

    parser.add_argument(
        "--raw-data-dir",
        type=str,
        default=str(Path("data") / "inputs" / "raw_data" / "inpatient"),
        help="Directory containing raw data files (default: data/inputs/raw_data/inpatient)",
    )

    parser.add_argument(
        "--force",
        action="store_true",
        help="Force update all files even if they've been processed before",
    )

    parser.add_argument(
        "--log-level",
        choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
        default="INFO",
        help="Set the logging level",
    )

    return parser.parse_args()

def main() -> int:
    """Main entry point for the FFT pipeline."""
    args = parse_arguments()

    try:
        # Run the full pipeline
        success = run_full_pipeline(
            raw_data_dir=args.raw_data_dir,
            force_update=args.force
        )

        if success:
            print("✅ FFT Pipeline completed successfully!")
            return 0
        else:
            print("❌ FFT Pipeline failed. Check logs for details.")
            return 1
    except Exception as e:
        print(f"❌ An error occurred: {e}")
        return 1

if __name__ == "__main__":
    sys.exit(main())

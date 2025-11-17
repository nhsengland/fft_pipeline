"""
Main FFT Pipeline
---------------

This module orchestrates the entire FFT pipeline process:
1. Discovers raw data files
2. Updates the Monthly Rolling Totals file chronologically
3. Runs the inpatient FFT process to generate output files
"""

import os
import sys
import logging
import importlib.util
import subprocess
from pathlib import Path
from datetime import datetime
from typing import Optional

from .discovery import discover_fft_files, get_file_periods
from .rolling_totals import update_monthly_rolling_totals, check_monthly_rolling_totals

# Configure logging
def setup_logging() -> None:
    """Configure logging for the pipeline."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_dir = Path("logfiles") / "full_pipeline"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_filename = str(log_dir / f"full_pipeline_{timestamp}.log")

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[
            logging.FileHandler(log_filename),
            logging.StreamHandler()
        ]
    )

def run_full_pipeline(raw_data_dir: str = None, force_update: bool = False) -> bool:
    """
    Run the full FFT pipeline process.

    Args:
        raw_data_dir: Directory containing raw data files (optional)
        force_update: Whether to force update all files even if they've been processed before

    Returns:
        True if successful, False otherwise
    """
    setup_logging()
    logging.info("Starting full FFT pipeline process")

    # Step 1: Discover FFT files
    logging.info("Discovering FFT data files...")
    try:
        discovered_files = discover_fft_files(raw_data_dir)
        if not discovered_files:
            logging.error("No FFT data files found")
            return False

        file_paths = [file_path for file_path, _ in discovered_files]
        periods = get_file_periods(discovered_files)
        logging.info(f"Discovered {len(file_paths)} files: {periods}")

        # Step 2: Determine which files need processing
        files_to_process = file_paths
        if not force_update:
            # Check existing periods in Monthly Rolling Totals
            df = check_monthly_rolling_totals()
            if df is not None:
                existing_periods = df["FFT Period"].tolist()
                logging.info(f"Existing periods in Monthly Rolling Totals: {existing_periods}")

                # Filter out periods that already exist
                files_to_process = []
                for file_path, period in discovered_files:
                    if period not in existing_periods:
                        files_to_process.append(file_path)
                        logging.info(f"File for period {period} will be processed: {file_path}")
                    else:
                        logging.info(f"File for period {period} already processed, skipping: {file_path}")

                if not files_to_process:
                    logging.info("All discovered files have already been processed")
                    # Run the inpatient FFT process anyway to generate output
                    return run_inpatient_fft()

        # Step 3: Update Monthly Rolling Totals
        logging.info(f"Updating Monthly Rolling Totals with {len(files_to_process)} files...")
        success = update_monthly_rolling_totals(files_to_process)
        if not success:
            logging.error("Failed to update Monthly Rolling Totals")
            return False

        # Step 4: Run the inpatient FFT process
        return run_inpatient_fft()

    except Exception as e:
        logging.error(f"Error in pipeline: {e}")
        return False

def run_inpatient_fft() -> bool:
    """
    Run the inpatient FFT process to generate the output file.

    Returns:
        True if successful, False otherwise
    """
    logging.info("Running inpatient FFT process...")

    # Create necessary directories
    log_dirs = [
        Path("logfiles") / "inpatient_fft",
        Path("logfiles") / "full_pipeline",
        Path("data") / "outputs",
        Path("data") / "rolling_totals"
    ]

    for log_dir in log_dirs:
        log_dir.mkdir(parents=True, exist_ok=True)
        logging.info(f"Ensured log directory exists: {log_dir}")

    try:
        result = subprocess.run([sys.executable, "inpatient_fft.py"],
                               capture_output=True, text=True, check=False)

        if result.returncode == 0:
            logging.info("Inpatient FFT process completed successfully")
            logging.info(result.stdout)
            return True
        else:
            logging.error(f"Inpatient FFT process failed with code {result.returncode}")
            logging.error(f"Error output: {result.stderr}")
            return False
    except Exception as e:
        logging.error(f"Error running inpatient FFT process: {e}")
        return False

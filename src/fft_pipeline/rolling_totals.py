"""
Rolling Totals Module
-------------------

This module handles updating the Monthly Rolling Totals file and ensuring chronological ordering.
"""

import os
import shutil
import logging
import pandas as pd
from pathlib import Path
from datetime import datetime
from typing import Optional

def update_monthly_rolling_totals(file_paths: list[str], backup: bool = True) -> bool:
    """
    Update the Monthly Rolling Totals file with data from the specified files.
    Processes files in chronological order and ensures proper ordering of entries.

    Args:
        file_paths: List of FFT file paths to process
        backup: Whether to create a backup of the Monthly Rolling Totals file

    Returns:
        True if successful, False otherwise
    """
    # If no files provided, nothing to do
    if not file_paths:
        logging.warning("No files provided for updating Monthly Rolling Totals")
        return False

    # Path to the Monthly Rolling Totals file
    rolling_totals_path = Path("data") / "rolling_totals" / "Monthly Rolling Totals.xlsx"

    # Create a backup if requested
    if backup:
        create_backup(rolling_totals_path)

    # Process each file in the provided list
    success = True
    for file_path in file_paths:
        try:
            logging.info(f"Processing {file_path}")
            result = process_file(file_path)
            if not result:
                logging.error(f"Failed to process {file_path}")
                success = False
        except Exception as e:
            logging.error(f"Error processing {file_path}: {e}")
            success = False

    # Ensure the entries are in chronological order
    if success:
        success = reorder_rolling_totals()

    return success

def create_backup(file_path: Path) -> Path:
    """
    Create a backup of the specified file.

    Args:
        file_path: Path to the file to back up

    Returns:
        Path to the backup file
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = file_path.with_name(f"{file_path.stem}.backup.{timestamp}{file_path.suffix}")

    logging.info(f"Creating backup at {backup_path}")
    shutil.copy2(file_path, backup_path)

    return backup_path

def process_file(file_path: str) -> bool:
    """
    Process a specific FFT file to update the Monthly Rolling Totals.
    Internally calls process_specific_month.py with the specified file.

    Args:
        file_path: Path to the FFT file to process

    Returns:
        True if successful, False otherwise
    """
    try:
        # Import the process_month module
        import sys
        import importlib.util

        spec = importlib.util.spec_from_file_location("process_month",
                                                     "src/scripts/process_month.py")
        process_month_module = importlib.util.module_from_spec(spec)
        sys.modules["process_month"] = process_month_module
        spec.loader.exec_module(process_month_module)

        # Call the process_specific_month function
        result = process_month_module.process_specific_month(file_path)
        return result
    except Exception as e:
        logging.error(f"Error in process_file: {e}")
        return False

def reorder_rolling_totals() -> bool:
    """
    Reorder the entries in the Monthly Rolling Totals file to ensure chronological order.
    Calls the fix_rolling_totals_order.py script.

    Returns:
        True if successful, False otherwise
    """
    try:
        # Import the fix_rolling_totals module
        import sys
        import importlib.util

        spec = importlib.util.spec_from_file_location("fix_rolling_totals",
                                                      "src/scripts/fix_rolling_totals.py")
        fix_rolling_totals_module = importlib.util.module_from_spec(spec)
        sys.modules["fix_rolling_totals"] = fix_rolling_totals_module
        spec.loader.exec_module(fix_rolling_totals_module)

        # Call the fix_rolling_totals_order function
        result = fix_rolling_totals_module.fix_rolling_totals_order()
        return result
    except Exception as e:
        logging.error(f"Error in reorder_rolling_totals: {e}")
        return False

def check_monthly_rolling_totals() -> Optional[pd.DataFrame]:
    """
    Check the current state of the Monthly Rolling Totals file.

    Returns:
        DataFrame containing the Monthly Rolling Totals or None if file doesn't exist
    """
    rolling_totals_path = Path("data") / "rolling_totals" / "Monthly Rolling Totals.xlsx"

    if not rolling_totals_path.exists():
        logging.error(f"Monthly Rolling Totals file does not exist at {rolling_totals_path}")
        return None

    try:
        df = pd.read_excel(rolling_totals_path, sheet_name='IP')
        return df
    except Exception as e:
        logging.error(f"Error reading Monthly Rolling Totals file: {e}")
        return None

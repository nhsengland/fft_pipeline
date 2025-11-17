"""
File Discovery Module
--------------------

This module handles the discovery of FFT raw data files in the inputs directory.
"""

import os
import re
import logging
from pathlib import Path
from datetime import datetime
from typing import Optional

# Define the pattern to match FFT inpatient files
FFT_INPATIENT_PATTERN = r"FFT_Inpatients_V\d+\s+([\w\-]+)\.xlsx"

def discover_fft_files(input_dir: str = None) -> list[tuple[str, str]]:
    """
    Discover all FFT inpatient files in the input directory and extract their period information.

    Args:
        input_dir: Path to the directory containing raw data files. If None, uses default.

    Returns:
        A list of tuples (file_path, period) sorted chronologically by period.
    """
    if input_dir is None:
        # Use default path
        input_dir = Path("data") / "inputs" / "raw_data" / "inpatient"
    else:
        input_dir = Path(input_dir)

    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory {input_dir} does not exist")

    # Find all Excel files matching the pattern
    files = []
    for file in input_dir.glob("*.xlsx"):
        match = re.match(FFT_INPATIENT_PATTERN, file.name)
        if match:
            period = match.group(1)  # Extract the period (e.g., "Jun-25")
            files.append((str(file), period))

    if not files:
        logging.warning(f"No FFT inpatient files found in {input_dir}")
        return []

    # Sort by period chronologically
    return sort_files_by_period(files)

def sort_files_by_period(files: list[tuple[str, str]]) -> list[tuple[str, str]]:
    """
    Sort a list of (file_path, period) tuples by period chronologically.

    Args:
        files: List of tuples (file_path, period) where period is in format "Mon-YY"

    Returns:
        Sorted list of tuples
    """
    # Create a mapping of month abbreviations to numeric values
    month_order = {
        'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4,
        'May': 5, 'Jun': 6, 'Jul': 7, 'Aug': 8,
        'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12
    }

    def get_sort_key(period: str) -> tuple[int, int]:
        """Extract year and month from period for sorting"""
        if not isinstance(period, str):
            return (0, 0)  # Default for non-string periods

        parts = period.split('-')
        if len(parts) != 2:
            return (0, 0)  # Default for invalid format

        month = parts[0][:3]  # Extract first three characters (month abbreviation)
        year = int(parts[1])  # Extract year (last two digits)

        return (year, month_order.get(month, 0))

    # Sort the files by the extracted year and month
    return sorted(files, key=lambda x: get_sort_key(x[1]))

def get_file_periods(files: list[tuple[str, str]]) -> list[str]:
    """
    Extract the period information from a list of file tuples.

    Args:
        files: List of tuples (file_path, period)

    Returns:
        List of period strings (e.g. ["Jun-25", "Jul-25", "Aug-25"])
    """
    return [period for _, period in files]

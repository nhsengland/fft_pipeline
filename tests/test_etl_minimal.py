"""
Minimal test suite for the ETL functions.

Most function testing is now handled by doctests directly in the function docstrings.
Run doctests with: uv run python -m doctest src/etl_functions.py -v

This file contains only tests for functions without doctests or tests requiring
complex fixtures that are difficult to implement as doctests.
"""

from src.etl_functions import *
import pytest
import pandas as pd
import os
import glob
from unittest.mock import patch, MagicMock
from pathlib import Path
import tempfile


def test_list_excel_files_with_tempdir():
    """
    Test list_excel_files function using a temporary directory with test files.

    This test creates a controlled environment with files with known modification times.
    """
    # Create a temporary directory
    with tempfile.TemporaryDirectory() as temp_dir:
        # Create test files with known modification times
        file1 = Path(temp_dir) / "IPFFT-Jan24.xlsx"
        file2 = Path(temp_dir) / "IPFFT-Feb24.xlsx"
        file3 = Path(temp_dir) / "IPFFT-Mar24.xlsx"

        # Touch the files to create them
        file1.touch()
        file2.touch()
        file3.touch()

        # Set modification times in ascending order
        os.utime(file1, (1000, 1000))
        os.utime(file2, (2000, 2000))
        os.utime(file3, (3000, 3000))

        # Set parameters
        file_pattern = "IPFFT-*.xlsx"
        prefix_suffix_separator = "-"
        date_format = "%b%y"

        # Act: Call the function with our temp directory
        result = list_excel_files(temp_dir, file_pattern, prefix_suffix_separator, date_format)

        # Assert: Verify the returned file list
        assert len(result) == 3
        # Files should be sorted by modification time (newest first)
        assert "Mar24" in str(result[0])
        assert "Feb24" in str(result[1])
        assert "Jan24" in str(result[2])


def test_list_excel_files_raises_error_when_no_files_found():
    """
    Test that an error is raised when no files are found

    Using a temporary directory to ensure it's empty
    """
    # Create a temporary directory that will definitely be empty
    with tempfile.TemporaryDirectory() as temp_dir:
        # Set parameters
        file_pattern = "NonExistent-*.xlsx"
        prefix_suffix_separator = "-"
        date_format = "%b%y"

        # Act & Assert: Verify that ValueError is raised
        with pytest.raises(ValueError, match="No matching Excel files found"):
            list_excel_files(temp_dir, file_pattern, prefix_suffix_separator, date_format)


@pytest.mark.skip(reason="Need to create a macro-enabled test file")
def test_open_macro_excel_file():
    """
    Test that open_macro_excel_file can load a workbook.

    This test would ideally use a real macro-enabled Excel file,
    but we'll mock the function for now.
    """
    # This test requires a real macro-enabled Excel file
    # Since creating one is complex and requires Office software,
    # we'll skip this test for now
    pass


def test_write_dataframes_to_sheets():
    """
    Test that write_dataframes_to_sheets correctly writes to specified sheets
    """
    # Arrange: Create a mock workbook
    wb = MagicMock()
    ws = MagicMock()
    wb.__getitem__.return_value = ws
    wb.sheetnames = ["Sheet1", "Sheet2"]

    # Create test dataframes
    df1 = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
    df2 = pd.DataFrame({"C": [5, 6], "D": [7, 8]})

    # Create the dfs_info list of tuples as expected by the function
    # Each tuple is (dataframe, sheet_name, start_row, start_col)
    dfs_info = [
        (df1, "Sheet1", 1, 1),
        (df2, "Sheet2", 1, 1)
    ]

    # Act
    write_dataframes_to_sheets(wb, dfs_info)

    # Assert: Check the workbook was accessed correctly
    wb.__getitem__.assert_any_call("Sheet1")
    wb.__getitem__.assert_any_call("Sheet2")
    assert wb.__getitem__.call_count == 2
"""
Minimal test suite for the ETL functions.

Most function testing is now handled by doctests directly in the function docstrings.
Run doctests with: uv run python -m doctest src/etl_functions.py -v

This file contains only tests for functions without doctests or tests requiring
complex fixtures that are difficult to implement as doctests.
"""

import glob
import os
import tempfile
import time
from pathlib import Path
from unittest.mock import MagicMock, patch

import pandas as pd
import pytest

from src.etl_functions import *


def test_list_excel_files_with_tempdir():
    """Test list_excel_files function with a temporary directory and actual files."""
    # Create a temp directory for our test files
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)

        # Create test files with dates in the filename
        file_pattern = 'FFTestIP-*.xlsx'
        prefix_suffix_separator = '-'
        date_format = '%b%y'

        # Create files in reverse chronological order
        file1 = temp_path / 'FFTestIP-Jan24.xlsx'
        file2 = temp_path / 'FFTestIP-Feb24.xlsx'
        file3 = temp_path / 'FFTestIP-Mar24.xlsx'

        # Create the files
        file1.touch()
        file2.touch()
        file3.touch()

        # Set modification times explicitly (oldest to newest)
        os.utime(file1, (1000, 1000))
        os.utime(file2, (2000, 2000))
        os.utime(file3, (3000, 3000))

        # Act: Call the function with our temp directory
        result = list_excel_files(
            temp_dir,
            file_pattern,
            prefix_suffix_separator,
            date_format
        )

        # Assert: Verify the returned file list
        EXPECTED_FILES = 3
        assert len(result) == EXPECTED_FILES
        # Files should be sorted by modification time (newest first)
        assert "Mar24" in str(result[0])
        assert "Feb24" in str(result[1])
        assert "Jan24" in str(result[2])


def test_validate_column_length_empty_df():
    """Test validate_column_length with an empty DataFrame."""
    # Arrange
    df = pd.DataFrame(columns=['Org Code'])

    # Act and Assert: Function should return the empty DataFrame unchanged
    result = validate_column_length(df, 'Org Code', 3)

    # Verify the result is identical to the input
    pd.testing.assert_frame_equal(result, df)


def test_validate_numeric_columns_empty_df():
    """Test validate_numeric_columns with an empty DataFrame."""
    # Arrange
    df = pd.DataFrame(columns=['Value'])

    # Act and Assert: Function should return the empty DataFrame unchanged
    result = validate_numeric_columns(df, 'Value', 'int')

    # Verify the result is identical to the input
    pd.testing.assert_frame_equal(result, df)


@patch('src.etl_functions.load_workbook')
def test_open_macro_excel_file(mock_load_workbook):
    """Test open_macro_excel_file function."""
    # Arrange
    file_path = 'test.xlsm'
    mock_wb = MagicMock()
    mock_load_workbook.return_value = mock_wb

    # Act
    wb = open_macro_excel_file(file_path)

    # Assert
    mock_load_workbook.assert_called_once_with(str(Path(file_path)), keep_vba=True)
    assert wb == mock_wb


def test_write_dataframes_to_sheets():
    """Test write_dataframes_to_sheets function."""
    # Arrange: Create a mock workbook and worksheets
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
    EXPECTED_CALLS = 2
    assert wb.__getitem__.call_count == EXPECTED_CALLS

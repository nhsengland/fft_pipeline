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


# Mock file system for testing using monkeypath fixture
@pytest.fixture
def mock_filesystem(monkeypatch):
    # Arrange: Mock glob.glob method to simulate finding files
    mock_files = [
        "/mocked/path/IPFFT-Jan24.xlsx",
        "/mocked/path/IPFFT-Feb24.xlsx",
        "/mocked/path/IPFFT-Mar24.xlsx",
    ]

    # Mock the result of glob.glob
    monkeypatch.setattr(glob, "glob", lambda pattern: mock_files)

    # Mock os.path.join for consistent paths
    monkeypatch.setattr(os.path, "join", lambda *args: "/".join(args))

    return mock_files


def test_list_excel_files_finds_matching_files(mock_filesystem):
    """
    Test list_excel_files function returns all files of matching name format
    """
    # Arrange
    folder_path = "/mocked/path"
    file_pattern = "IPFFT-*.xlsx"
    prefix_suffix_separator = "-"
    date_format = "%b%y"

    # Act
    result = list_excel_files(folder_path, file_pattern, prefix_suffix_separator, date_format)

    # Assert: Verify the returned file list matches expected
    assert len(result) == 3
    assert result[0] == "/mocked/path/IPFFT-Mar24.xlsx"  # Most recent first
    assert result[-1] == "/mocked/path/IPFFT-Jan24.xlsx"  # Oldest last


def test_list_excel_files_raises_error_when_no_files_found(monkeypatch):
    """
    Test that an error is raised when no files are found
    """
    # Arrange: Mock glob.glob to return empty list
    monkeypatch.setattr(glob, "glob", lambda pattern: [])

    folder_path = "/mocked/path"
    file_pattern = "IPFFT-*.xlsx"
    prefix_suffix_separator = "-"
    date_format = "%b%y"

    # Act & Assert: Verify that ValueError is raised
    with pytest.raises(ValueError, match="No matching Excel files found"):
        list_excel_files(folder_path, file_pattern, prefix_suffix_separator, date_format)


def test_open_macro_excel_file(monkeypatch):
    """
    Test that the open_macro_excel_file function opens a macro-enabled Excel file correctly
    """
    # Arrange: Mock openpyxl.load_workbook
    mock_workbook = MagicMock()
    monkeypatch.setattr("openpyxl.load_workbook", lambda *args, **kwargs: mock_workbook)

    # Mock os.path.exists to return True
    monkeypatch.setattr(os.path, "exists", lambda x: True)

    file_path = "test.xlsm"

    # Act
    wb = open_macro_excel_file(file_path)

    # Assert
    assert wb == mock_workbook


def test_write_dataframes_to_sheets():
    """
    Test that write_dataframes_to_sheets correctly writes to specified sheets
    """
    # Arrange: Create a mock workbook and worksheet
    wb = MagicMock()
    ws = MagicMock()
    wb.__getitem__.return_value = ws

    # Create test dataframes
    df1 = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
    df2 = pd.DataFrame({"C": [5, 6], "D": [7, 8]})

    # Act
    write_dataframes_to_sheets(wb, {"Sheet1": df1, "Sheet2": df2})

    # Assert: Check the workbook was accessed correctly
    wb.__getitem__.assert_any_call("Sheet1")
    wb.__getitem__.assert_any_call("Sheet2")
    assert wb.__getitem__.call_count == 2

# Add more tests for functions that don't have doctests or require complex setup
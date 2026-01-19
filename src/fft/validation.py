"""Workbook comparison utilities for validation."""

import re
from datetime import datetime
from pathlib import Path
from typing import TypedDict

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter

from .config import VALIDATION_TOLERANCE

# Type for Excel cell values from openpyxl
CellValue = str | int | float | bool | datetime | None


class CellDifference(TypedDict):
    """Represents a difference between two cells."""

    sheet: str
    cell: str
    expected: CellValue
    actual: CellValue


class SheetResult(TypedDict):
    """Comparison result for a single sheet."""

    name: str
    identical: bool
    differences: list[CellDifference]
    missing_in_actual: bool
    missing_in_expected: bool


def compare_workbooks(
    expected_path: Path | str,
    actual_path: Path | str,
    sheets_to_compare: list[str] | None = None,
    data_only: bool = True,
) -> list[SheetResult]:
    """Compare two Excel workbooks and return detailed differences.

    Compares cell values (or formulas if data_only=False) between two workbooks.
    Useful for validating automatically generated output against known good results.

    Args:
        expected_path: Path to the ground truth workbook
        actual_path: Path to the workbook to validate
        sheets_to_compare: List of sheet names to compare (None = all common sheets)
        data_only: If True, compare computed values; if False, compare formulas

    Returns:
        List of SheetResult dicts, one per sheet compared

    Raises:
        FileNotFoundError: If either workbook doesn't exist

    >>> from pathlib import Path
    >>> from openpyxl import Workbook
    >>> import tempfile

    # Create two identical workbooks
    >>> with tempfile.TemporaryDirectory() as tmpdir:
    ...     expected = Path(tmpdir) / "expected.xlsx"
    ...     actual = Path(tmpdir) / "actual.xlsx"
    ...     wb1 = Workbook()
    ...     wb1.active.title = "Data"
    ...     wb1["Data"]["A1"] = "Hello"
    ...     wb1["Data"]["B1"] = 100
    ...     wb1.save(expected)
    ...     wb2 = Workbook()
    ...     wb2.active.title = "Data"
    ...     wb2["Data"]["A1"] = "Hello"
    ...     wb2["Data"]["B1"] = 100
    ...     wb2.save(actual)
    ...     results = compare_workbooks(expected, actual)
    ...     results[0]["identical"]
    True

    # Create workbooks with differences
    >>> with tempfile.TemporaryDirectory() as tmpdir:
    ...     expected = Path(tmpdir) / "expected.xlsx"
    ...     actual = Path(tmpdir) / "actual.xlsx"
    ...     wb1 = Workbook()
    ...     wb1.active.title = "Data"
    ...     wb1["Data"]["A1"] = "Hello"
    ...     wb1["Data"]["B1"] = 100
    ...     wb1.save(expected)
    ...     wb2 = Workbook()
    ...     wb2.active.title = "Data"
    ...     wb2["Data"]["A1"] = "Hello"
    ...     wb2["Data"]["B1"] = 999
    ...     wb2.save(actual)
    ...     results = compare_workbooks(expected, actual)
    ...     results[0]["identical"]
    False
    >>> with tempfile.TemporaryDirectory() as tmpdir:
    ...     expected = Path(tmpdir) / "expected.xlsx"
    ...     actual = Path(tmpdir) / "actual.xlsx"
    ...     wb1 = Workbook()
    ...     wb1.active.title = "Data"
    ...     wb1["Data"]["A1"] = "Hello"
    ...     wb1["Data"]["B1"] = 100
    ...     wb1.save(expected)
    ...     wb2 = Workbook()
    ...     wb2.active.title = "Data"
    ...     wb2["Data"]["A1"] = "Hello"
    ...     wb2["Data"]["B1"] = 999
    ...     wb2.save(actual)
    ...     results = compare_workbooks(expected, actual)
    ...     results[0]["differences"][0]["cell"]
    'B1'

    # Edge case: Missing file
    >>> compare_workbooks("nonexistent.xlsx", "also_missing.xlsx")
    Traceback (most recent call last):
        ...
    FileNotFoundError: Expected workbook not found: nonexistent.xlsx

    # Edge case: Sheet missing in actual
    >>> with tempfile.TemporaryDirectory() as tmpdir:
    ...     expected = Path(tmpdir) / "expected.xlsx"
    ...     actual = Path(tmpdir) / "actual.xlsx"
    ...     wb1 = Workbook()
    ...     _ = wb1.create_sheet("Extra")
    ...     wb1.save(expected)
    ...     wb2 = Workbook()
    ...     wb2.save(actual)
    ...     results = compare_workbooks(expected, actual, sheets_to_compare=["Extra"])
    ...     results[0]["missing_in_actual"]
    True

    """
    expected_path = Path(expected_path)
    actual_path = Path(actual_path)

    if not expected_path.exists():
        raise FileNotFoundError(f"Expected workbook not found: {expected_path}")
    if not actual_path.exists():
        raise FileNotFoundError(f"Actual workbook not found: {actual_path}")

    wb_expected = load_workbook(expected_path, data_only=data_only)
    wb_actual = load_workbook(actual_path, data_only=data_only)

    if sheets_to_compare is None:
        sheets_to_compare = list(set(wb_expected.sheetnames) | set(wb_actual.sheetnames))

    return [
        _compare_sheet(sheet_name, wb_expected, wb_actual)
        for sheet_name in sheets_to_compare
    ]


def _compare_sheet(sheet_name: str, wb_expected, wb_actual) -> SheetResult:
    """Compare a single sheet between two workbooks."""
    if sheet_name not in wb_expected.sheetnames:
        return {
            "name": sheet_name,
            "identical": False,
            "differences": [],
            "missing_in_expected": True,
            "missing_in_actual": False,
        }

    if sheet_name not in wb_actual.sheetnames:
        return {
            "name": sheet_name,
            "identical": False,
            "differences": [],
            "missing_in_expected": False,
            "missing_in_actual": True,
        }

    ws_expected = wb_expected[sheet_name]
    ws_actual = wb_actual[sheet_name]

    differences = _find_cell_differences(sheet_name, ws_expected, ws_actual)

    return {
        "name": sheet_name,
        "identical": len(differences) == 0,
        "differences": differences,
        "missing_in_expected": False,
        "missing_in_actual": False,
    }


def _find_cell_differences(
    sheet_name: str, ws_expected, ws_actual
) -> list[CellDifference]:
    """Find all cell differences between two worksheets."""
    max_row = max(ws_expected.max_row or 1, ws_actual.max_row or 1)
    max_col = max(ws_expected.max_column or 1, ws_actual.max_column or 1)

    differences = []
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            val_expected = ws_expected.cell(row=row, column=col).value
            val_actual = ws_actual.cell(row=row, column=col).value

            if not _values_are_equivalent(val_expected, val_actual):
                differences.append(
                    {
                        "sheet": sheet_name,
                        "cell": ws_expected.cell(row=row, column=col).coordinate,
                        "expected": val_expected,
                        "actual": val_actual,
                    }
                )

    return differences


def compare_data_by_key(  # noqa: PLR0913 # Justified: validation function needs all params
    expected_path: Path | str,
    actual_path: Path | str,
    sheet_name: str,
    *,
    key_column: str | list[str] = "B",  # Single column or composite key columns
    start_row: int = 15,
    data_only: bool = True,
) -> SheetResult:
    """Compare sheet data by matching records via key column(s) rather than row position.

    Args:
        expected_path: Path to ground truth workbook
        actual_path: Path to generated workbook
        sheet_name: Name of sheet to compare
        key_column: Column letter(s) for unique identification
                   - Single column: "B" for Trust_Code
                   - Composite key: ["B", "D", "F"] for Trust_Code + Site_Code + Ward_Name
        start_row: First row of data
        data_only: Compare values vs formulas

    Returns:
        SheetResult with differences between matching records

    """
    expected_path = Path(expected_path)
    actual_path = Path(actual_path)

    if not expected_path.exists():
        raise FileNotFoundError(f"Expected workbook not found: {expected_path}")
    if not actual_path.exists():
        raise FileNotFoundError(f"Actual workbook not found: {actual_path}")

    wb_expected = load_workbook(expected_path, data_only=data_only)
    wb_actual = load_workbook(actual_path, data_only=data_only)

    if sheet_name not in wb_expected.sheetnames:
        return {
            "name": sheet_name,
            "identical": False,
            "differences": [],
            "missing_in_expected": True,
            "missing_in_actual": False,
        }

    if sheet_name not in wb_actual.sheetnames:
        return {
            "name": sheet_name,
            "identical": False,
            "differences": [],
            "missing_in_expected": False,
            "missing_in_actual": True,
        }

    ws_expected = wb_expected[sheet_name]
    ws_actual = wb_actual[sheet_name]

    # Build dictionaries mapping key -> row data
    expected_records = _extract_records_by_key(ws_expected, key_column, start_row)
    actual_records = _extract_records_by_key(ws_actual, key_column, start_row)

    differences = _compare_records_by_key(sheet_name, expected_records, actual_records)

    return {
        "name": sheet_name,
        "identical": len(differences) == 0,
        "differences": differences,
        "missing_in_expected": False,
        "missing_in_actual": False,
    }


def _extract_records_by_key(
    worksheet, key_column: str | list[str], start_row: int
) -> dict:
    """Extract records from worksheet, keyed by identifier column(s).

    Args:
        worksheet: Excel worksheet
        key_column: Single column letter or list of column letters for composite key
        start_row: First row of data

    Returns:
        Dict mapping key(s) to row data

    """
    records = {}

    # Convert column letters to indices
    if isinstance(key_column, str):
        key_col_indices = [column_index_from_string(key_column.upper())]
    else:
        key_col_indices = [column_index_from_string(col.upper()) for col in key_column]

    max_row = worksheet.max_row or start_row
    max_col = worksheet.max_column or 10

    for row_num in range(start_row, max_row + 1):
        # Extract key value(s)
        key_parts = []
        for col_idx in key_col_indices:
            key_value = worksheet.cell(row=row_num, column=col_idx).value
            if key_value and str(key_value).strip():
                key_parts.append(str(key_value).strip())
            else:
                key_parts.append("")

        # Only process rows where all key parts are non-empty
        if all(part for part in key_parts):
            # Create composite key by joining with separator
            composite_key = "|".join(key_parts)

            row_data = {}
            for col_num in range(1, max_col + 1):
                cell_value = worksheet.cell(row=row_num, column=col_num).value
                # Store by column index, not absolute coordinate
                row_data[col_num] = cell_value
            records[composite_key] = row_data

    return records


def _compare_records_by_key(
    sheet_name: str, expected_records: dict, actual_records: dict
) -> list[CellDifference]:
    """Compare matching records between expected and actual data."""
    differences = []

    # Compare records that exist in both datasets
    for key, expected_row in expected_records.items():
        if key not in actual_records:
            continue  # Skip missing records for now

        actual_row = actual_records[key]

        # Compare all columns in the row
        for col_num in expected_row:
            expected_val = expected_row[col_num]
            actual_val = actual_row.get(col_num)

            if not _values_are_equivalent(expected_val, actual_val):
                # Convert column number back to letter for display
                col_letter = get_column_letter(col_num)
                differences.append(
                    {
                        "sheet": sheet_name,
                        "cell": f"{col_letter}({key})",  # Show column letter with key
                        "expected": expected_val,
                        "actual": actual_val,
                    }
                )

    return differences


def compare_data_range(
    expected_path: Path | str,
    actual_path: Path | str,
    sheet_name: str,
    start_row: int = 15,
    data_only: bool = True,
) -> SheetResult:
    """Compare only the data range of a sheet, ignoring template/control areas.

    Args:
        expected_path: Path to ground truth workbook
        actual_path: Path to generated workbook
        sheet_name: Name of sheet to compare
        start_row: First row of data (skip template controls)
        data_only: Compare values vs formulas

    Returns:
        SheetResult with differences in data area only

    """
    expected_path = Path(expected_path)
    actual_path = Path(actual_path)

    if not expected_path.exists():
        raise FileNotFoundError(f"Expected workbook not found: {expected_path}")
    if not actual_path.exists():
        raise FileNotFoundError(f"Actual workbook not found: {actual_path}")

    wb_expected = load_workbook(expected_path, data_only=data_only)
    wb_actual = load_workbook(actual_path, data_only=data_only)

    if sheet_name not in wb_expected.sheetnames:
        return {
            "name": sheet_name,
            "identical": False,
            "differences": [],
            "missing_in_expected": True,
            "missing_in_actual": False,
        }

    if sheet_name not in wb_actual.sheetnames:
        return {
            "name": sheet_name,
            "identical": False,
            "differences": [],
            "missing_in_expected": False,
            "missing_in_actual": True,
        }

    ws_expected = wb_expected[sheet_name]
    ws_actual = wb_actual[sheet_name]

    # Compare only from start_row onwards
    differences = _find_cell_differences_range(
        sheet_name, ws_expected, ws_actual, start_row
    )

    return {
        "name": sheet_name,
        "identical": len(differences) == 0,
        "differences": differences,
        "missing_in_expected": False,
        "missing_in_actual": False,
    }


def _find_cell_differences_range(
    sheet_name: str, ws_expected, ws_actual, start_row: int
) -> list[CellDifference]:
    """Find cell differences starting from specified row."""
    max_row = max(ws_expected.max_row or 1, ws_actual.max_row or 1)
    max_col = max(ws_expected.max_column or 1, ws_actual.max_column or 1)

    differences = []
    for row in range(start_row, max_row + 1):
        for col in range(1, max_col + 1):
            val_expected = ws_expected.cell(row=row, column=col).value
            val_actual = ws_actual.cell(row=row, column=col).value

            if not _values_are_equivalent(val_expected, val_actual):
                differences.append(
                    {
                        "sheet": sheet_name,
                        "cell": ws_expected.cell(row=row, column=col).coordinate,
                        "expected": val_expected,
                        "actual": val_actual,
                    }
                )

    return differences


def _values_are_equivalent(val_expected, val_actual) -> bool:
    """Check if two values are equivalent with reasonable tolerance.

    >>> _values_are_equivalent(0.025381903642773207, 0.02538190364277321)
    True
    >>> _values_are_equivalent(0.942, 0.942)
    True
    >>> _values_are_equivalent("Hello", "Hello")
    True
    >>> _values_are_equivalent("Hello", "World")
    False
    >>> _values_are_equivalent(100, 100)
    True
    >>> _values_are_equivalent(100.1, 100.2)
    False

    """
    # Exact match
    if val_expected == val_actual:
        return True

    # Handle template compatibility: None/NULL/NA/"0"/"-" are equivalent for missing values
    missing_vals = {None, "NULL", "NA", "", "0", "-", "nan"}
    if val_expected in missing_vals and val_actual in missing_vals:
        return True

    # For numeric comparisons
    try:
        num_expected = float(val_expected) if val_expected is not None else None
        num_actual = float(val_actual) if val_actual is not None else None

        if num_expected is None or num_actual is None:
            return False

        # Use absolute tolerance to eliminate floating point noise
        return abs(num_expected - num_actual) <= VALIDATION_TOLERANCE

    except (ValueError, TypeError):
        # Not numeric, compare as strings
        return str(val_expected) == str(val_actual)


def print_comparison_report(
    results: list[SheetResult], max_diffs_per_sheet: int = 10
) -> None:
    """Print a human-readable comparison report.

    Args:
        results: List of SheetResult from compare_workbooks()
        max_diffs_per_sheet: Maximum differences to show per sheet

    >>> from pathlib import Path
    >>> from openpyxl import Workbook
    >>> import tempfile
    >>> with tempfile.TemporaryDirectory() as tmpdir:
    ...     expected = Path(tmpdir) / "expected.xlsx"
    ...     actual = Path(tmpdir) / "actual.xlsx"
    ...     wb1 = Workbook()
    ...     wb1.active.title = "Data"
    ...     wb1["Data"]["A1"] = "Test"
    ...     wb1.save(expected)
    ...     wb2 = Workbook()
    ...     wb2.active.title = "Data"
    ...     wb2["Data"]["A1"] = "Test"
    ...     wb2.save(actual)
    ...     results = compare_workbooks(expected, actual)
    ...     print_comparison_report(results)
    Sheet: Data ✓ (identical)
    <BLANKLINE>
    Summary: 1/1 sheets identical

    """
    total_sheets = len(results)
    identical_sheets = sum(1 for r in results if r["identical"])

    for result in results:
        if result["missing_in_expected"]:
            print(f"Sheet: {result['name']} ⚠ (missing in expected workbook)")
        elif result["missing_in_actual"]:
            print(f"Sheet: {result['name']} ⚠ (missing in actual workbook)")
        elif result["identical"]:
            print(f"Sheet: {result['name']} ✓ (identical)")
        else:
            print(f"Sheet: {result['name']} ✗ ({len(result['differences'])} differences)")
            for diff in result["differences"][:max_diffs_per_sheet]:
                print(
                    f"  - Cell {diff['cell']}: expected {diff['expected']!r}, "
                    f"got {diff['actual']!r}"
                )
            if len(result["differences"]) > max_diffs_per_sheet:
                print(
                    f"  ... and {len(result['differences']) - max_diffs_per_sheet} more"
                )

    print()
    print(f"Summary: {identical_sheets}/{total_sheets} sheets identical")


def find_matching_ground_truth(output_path: Path, ground_truth_dir: Path) -> Path | None:
    """Find the best matching ground truth file for a given pipeline output.

    Matches files based on month/period and service type patterns rather than
    exact filenames, allowing for different naming conventions.

    Args:
        output_path: Path to the pipeline output file
        ground_truth_dir: Directory containing ground truth files

    Returns:
        Path to the best matching ground truth file, or None if no match found

    >>> from pathlib import Path
    >>> import tempfile

    # Test with matching month and service type
    >>> with tempfile.TemporaryDirectory() as tmpdir:
    ...     truth_dir = Path(tmpdir) / "ground_truth"
    ...     truth_dir.mkdir()
    ...     # Create ground truth file with different naming pattern
    ...     (truth_dir / "021225_133523_FFT_IP_MacroWebfile_Jul-25.xlsm").touch()
    ...     # Pipeline output with different pattern but same month/type
    ...     output = Path(tmpdir) / "FFT-inpatient-data-Jul-25.xlsm"
    ...     match = find_matching_ground_truth(output, truth_dir)
    ...     match.name if match else None
    '021225_133523_FFT_IP_MacroWebfile_Jul-25.xlsm'

    # Test with no matching files
    >>> with tempfile.TemporaryDirectory() as tmpdir:
    ...     truth_dir = Path(tmpdir) / "ground_truth"
    ...     truth_dir.mkdir()
    ...     output = Path(tmpdir) / "FFT-inpatient-data-Aug-25.xlsm"
    ...     match = find_matching_ground_truth(output, truth_dir)
    ...     match is None
    True

    # Test with multiple files, should pick best match
    >>> with tempfile.TemporaryDirectory() as tmpdir:
    ...     truth_dir = Path(tmpdir) / "ground_truth"
    ...     truth_dir.mkdir()
    ...     # Create files with different months
    ...     (truth_dir / "FFT_IP_Jul-25.xlsm").touch()
    ...     (truth_dir / "FFT_IP_Aug-25.xlsm").touch()
    ...     (truth_dir / "FFT_AE_Jul-25.xlsm").touch()  # Different service type
    ...     # Should match IP Jul-25, not AE Jul-25 or IP Aug-25
    ...     output = Path(tmpdir) / "FFT-inpatient-data-Jul-25.xlsm"
    ...     match = find_matching_ground_truth(output, truth_dir)
    ...     match.name if match else None
    'FFT_IP_Jul-25.xlsm'

    """
    if not ground_truth_dir.exists():
        return None

    output_month = _extract_month_pattern(output_path.name)
    output_service = _extract_service_type(output_path.name)

    if not output_month or not output_service:
        return None

    best_match = None
    best_score = 0

    # Check all files in ground truth directory
    for truth_file in ground_truth_dir.glob("*.xl*"):
        truth_month = _extract_month_pattern(truth_file.name)
        truth_service = _extract_service_type(truth_file.name)

        if not truth_month or not truth_service:
            continue

        # Score the match
        score = 0

        # Month must match exactly
        if truth_month == output_month:
            score += 100
        else:
            continue  # No match without matching month

        # Service type should match
        if truth_service == output_service:
            score += 50

        # Prefer newer files (by filename timestamp if present)
        if truth_file.stat().st_mtime > (best_match.stat().st_mtime if best_match else 0):
            score += 1

        if score > best_score:
            best_match = truth_file
            best_score = score

    return best_match


def _extract_month_pattern(filename: str) -> str | None:
    """Extract month pattern (e.g., 'Jul-25', 'Aug-25') from filename.

    >>> _extract_month_pattern("FFT-inpatient-data-Jul-25.xlsm")
    'Jul-25'
    >>> _extract_month_pattern("021225_133523_FFT_IP_MacroWebfile_Aug-25.xlsm")
    'Aug-25'
    >>> _extract_month_pattern("no-month-pattern.xlsx")

    """
    # Look for pattern like Jul-25, Aug-25, etc.
    pattern = r"([A-Z][a-z]{2}-\d{2})"
    match = re.search(pattern, filename)
    return match.group(1) if match else None


def _extract_service_type(filename: str) -> str | None:
    """Extract normalized service type from filename.

    Maps various service type representations to standard codes:
    - inpatient/IP -> inpatient
    - ae/AE -> ae
    - ambulance/AMB -> ambulance

    >>> _extract_service_type("FFT-inpatient-data-Jul-25.xlsm")
    'inpatient'
    >>> _extract_service_type("021225_133523_FFT_IP_MacroWebfile_Jul-25.xlsm")
    'inpatient'
    >>> _extract_service_type("FFT_AE_Aug-25.xlsm")
    'ae'
    >>> _extract_service_type("FFT_AMB_Jul-25.xlsm")
    'ambulance'
    >>> _extract_service_type("random-file.xlsm")

    """
    filename_lower = filename.lower()

    # Check for inpatient patterns
    if any(pattern in filename_lower for pattern in ["inpatient", "_ip_", "fft_ip"]):
        return "inpatient"

    # Check for A&E patterns
    if any(pattern in filename_lower for pattern in ["_ae_", "fft_ae"]):
        return "ae"

    # Check for ambulance patterns
    if any(pattern in filename_lower for pattern in ["ambulance", "_amb_", "fft_amb"]):
        return "ambulance"

    return None

"""Excel output functions."""

from pathlib import Path
import pandas as pd

from openpyxl import load_workbook
from openpyxl.workbook import Workbook

from src.fft.config import (
    TEMPLATES_DIR,
    TEMPLATE_CONFIG,
    BS_SHEET_CONFIG,
    PERIOD_LABEL_CONFIG,
    TEMPLATE_CONFIG,
    PERCENTAGE_COLUMN_CONFIG,
    OUTPUTS_DIR,
)


def load_template(service_type: str) -> Workbook:
    """Load Excel template for the specified service type, preserving VBA macros.

    Args:
        service_type: 'inpatient', 'ae', or 'ambulance'

    Returns:
        Openpyxl Workbook object with VBA preserved

    Raises:
        KeyError: If service_type is not configured
        FileNotFoundError: If template file doesn't exist

    >>> from src.fft.writers import load_template
    >>> wb = load_template('inpatient')
    >>> type(wb).__name__
    'Workbook'
    >>> 'ICB' in wb.sheetnames
    True

    # Edge case: Unknown service type
    >>> load_template('unknown_service')
    Traceback (most recent call last):
        ...
    KeyError: "Unknown service type: 'unknown_service'"

    # Edge case: Missing template file
    >>> from src.fft.config import TEMPLATE_CONFIG
    >>> TEMPLATE_CONFIG['test_missing'] = {'template_file': 'nonexistent.xlsm'}
    >>> load_template('test_missing')
    Traceback (most recent call last):
        ...
    FileNotFoundError: Template not found: nonexistent.xlsm
    >>> del TEMPLATE_CONFIG['test_missing']
    """
    if service_type not in TEMPLATE_CONFIG:
        raise KeyError(f"Unknown service type: '{service_type}'")

    template_file = TEMPLATE_CONFIG[service_type]["template_file"]
    template_path = TEMPLATES_DIR / template_file

    if not template_path.exists():
        raise FileNotFoundError(f"Template not found: {template_file}")

    return load_workbook(template_path, keep_vba=True)


# %%
def write_dataframe_to_sheet(
    workbook: Workbook,
    df: pd.DataFrame,
    sheet_name: str,
    start_row: int,
    start_col: int = 1,
) -> None:
    """Write DataFrame contents to a specific sheet location.

    Writes data without headers - assumes template already has headers in place.

    Args:
        workbook: Openpyxl Workbook object
        df: DataFrame to write
        sheet_name: Name of target sheet
        start_row: Row number to start writing (1-indexed)
        start_col: Column number to start writing (1-indexed, default 1)

    Returns:
        None (modifies workbook in place)

    Raises:
        KeyError: If sheet_name doesn't exist in workbook

    >>> from src.fft.writers import load_template, write_dataframe_to_sheet
    >>> import pandas as pd
    >>> wb = load_template('inpatient')
    >>> df = pd.DataFrame({
    ...     'ICB Code': ['ABC', 'DEF'],
    ...     'ICB Name': ['Test ICB 1', 'Test ICB 2'],
    ...     'Total Responses': [100, 200]
    ... })
    >>> write_dataframe_to_sheet(wb, df, 'ICB', start_row=15, start_col=1)
    >>> wb['ICB'].cell(row=15, column=1).value
    'ABC'
    >>> wb['ICB'].cell(row=16, column=2).value
    'Test ICB 2'

    # Edge case: Sheet doesn't exist
    >>> write_dataframe_to_sheet(wb, df, 'NonExistent', start_row=15)
    Traceback (most recent call last):
        ...
    KeyError: "Sheet 'NonExistent' not found in workbook"

    # Edge case: Empty DataFrame
    >>> df_empty = pd.DataFrame(columns=['A', 'B'])
    >>> write_dataframe_to_sheet(wb, df_empty, 'ICB', start_row=20)
    >>> wb['ICB'].cell(row=20, column=1).value is None
    True
    """
    if sheet_name not in workbook.sheetnames:
        raise KeyError(f"Sheet '{sheet_name}' not found in workbook")

    sheet = workbook[sheet_name]

    for row_idx, row in enumerate(df.itertuples(index=False), start=start_row):
        for col_idx, value in enumerate(row, start=start_col):
            sheet.cell(row=row_idx, column=col_idx).value = value


# %%
def write_bs_lookup_data(
    workbook: Workbook, ward_df: pd.DataFrame, service_type: str
) -> None:
    """Populate BS sheet lookup tables for dropdown filtering.

    Writes reference lists and linked lists used by VBA macros for filtering.

    The BS sheet has two main sections:
    1. Reference Lists (U:Z) - Full hierarchy from ward data
    2. Linked Lists (AE+) - Deduplicated, sorted lists for each tab's dropdowns

    Args:
        workbook: Openpyxl Workbook object
        ward_df: DataFrame containing full ward-level data with hierarchy columns
        service_type: 'inpatient', 'ae', or 'ambulance'

    Returns:
        None (modifies workbook in place)

    Raises:
        KeyError: If BS sheet doesn't exist or required columns missing

    >>> from src.fft.writers import load_template, write_bs_lookup_data
    >>> import pandas as pd
    >>> wb = load_template('inpatient')
    >>> df = pd.DataFrame({
    ...     'ICB Code': ['ABC', 'ABC', 'DEF'],
    ...     'ICB Name': ['ICB North', 'ICB North', 'ICB South'],
    ...     'Trust Code': ['T01', 'T01', 'T02'],
    ...     'Trust Name': ['Trust A', 'Trust A', 'Trust B'],
    ...     'Site Code': ['S01', 'S02', 'S03'],
    ...     'Site Name': ['Site 1', 'Site 2', 'Site 3'],
    ...     'Ward Name': ['Ward 1', 'Ward 2', 'Ward 3']
    ... })
    >>> write_bs_lookup_data(wb, df, 'inpatient')
    >>> wb['BS'].cell(row=2, column=21).value
    'ABC'

    # Edge case: Missing BS sheet
    >>> from openpyxl import Workbook as NewWorkbook
    >>> wb_no_bs = NewWorkbook()
    >>> write_bs_lookup_data(wb_no_bs, df, 'inpatient')
    Traceback (most recent call last):
        ...
    KeyError: "Sheet 'BS' not found in template workbook"

    # Edge case: Missing required columns
    >>> df_missing = pd.DataFrame({'ICB Code': ['ABC']})
    >>> wb2 = load_template('inpatient')
    >>> write_bs_lookup_data(wb2, df_missing, 'inpatient')
    Traceback (most recent call last):
        ...
    KeyError: "Required column 'ICB Name' not found in DataFrame"
    """
    if "BS" not in workbook.sheetnames:
        raise KeyError("Sheet 'BS' not found in template workbook")

    if service_type not in BS_SHEET_CONFIG:
        raise KeyError(f"Unknown service type: '{service_type}'")

    config = BS_SHEET_CONFIG[service_type]
    sheet = workbook["BS"]

    # Validate required columns exist
    ref_cols = config["reference_columns"]
    for col in ref_cols:
        if col not in ward_df.columns:
            raise KeyError(f"Required column '{col}' not found in DataFrame")

    # 1. Write Reference Lists (full hierarchy to U:Z)
    ref_start_col = config["reference_list_start_col"]
    ref_start_row = config["reference_list_start_row"]

    ref_data = ward_df[ref_cols].copy()
    for row_idx, row in enumerate(ref_data.itertuples(index=False), start=ref_start_row):
        for col_idx, value in enumerate(row, start=ref_start_col):
            sheet.cell(row=row_idx, column=col_idx).value = value

    # 2. Write Linked Lists (deduplicated, sorted for each tab)
    for level, level_config in config["linked_lists"].items():
        start_col = level_config["start_col"]
        columns = level_config["columns"]

        # Extract unique combinations, sort alphabetically
        unique_df = (
            ward_df[columns]
            .drop_duplicates()
            .sort_values(by=columns)
            .reset_index(drop=True)
        )

        # Write each column separately (they get sorted independently)
        for col_offset, col_name in enumerate(columns):
            col_values = (
                unique_df[col_name].drop_duplicates().sort_values().reset_index(drop=True)
            )
            for row_idx, value in enumerate(col_values, start=1):
                sheet.cell(row=row_idx, column=start_col + col_offset).value = value


# %%
def update_period_labels(workbook: Workbook, service_type: str, fft_period: str) -> None:
    """Update period-dependent labels in the workbook.

    Updates cells containing dynamic period text (e.g., "August 2024").

    Args:
        workbook: Openpyxl Workbook object
        service_type: 'inpatient', 'ae', or 'ambulance'
        fft_period: FFT period string (e.g., 'Aug-24')

    Returns:
        None (modifies workbook in place)

    Raises:
        KeyError: If service_type is not configured or sheet doesn't exist

    >>> from src.fft.writers import load_template, update_period_labels
    >>> wb = load_template('inpatient')
    >>> update_period_labels(wb, 'inpatient', 'Aug-24')
    >>> wb['Notes'].cell(row=2, column=1).value
    'Inpatient Friends and Family Test (FFT) Data - Aug-24'

    # Edge case: Unknown service type
    >>> update_period_labels(wb, 'unknown', 'Aug-24')
    Traceback (most recent call last):
        ...
    KeyError: "Unknown service type: 'unknown'"

    # Edge case: Missing sheet
    >>> from openpyxl import Workbook as NewWorkbook
    >>> wb_no_notes = NewWorkbook()
    >>> update_period_labels(wb_no_notes, 'inpatient', 'Aug-24')
    Traceback (most recent call last):
        ...
    KeyError: "Sheet 'Notes' not found in workbook"

    # Edge case: Missing configuration for service type
    >>> from src.fft.config import PERIOD_LABEL_CONFIG
    >>> PERIOD_LABEL_CONFIG['test_missing'] = {
    ...     'notes_title': {
    ...         'sheet': 'Notes',
    ...         'cell': 'A2',
    ...         'template': 'Test FFT Data - {period}',
    ...     }
    ... }
    >>> update_period_labels(wb, 'test_missing', 'Aug-24') # No error should be raised
    >>> del PERIOD_LABEL_CONFIG['test_missing']

    # Edge case: Empty configuration for service type
    >>> PERIOD_LABEL_CONFIG['test_empty'] = {}
    >>> update_period_labels(wb, 'test_empty', 'Aug-24')  # No error should be raised
    >>> del PERIOD_LABEL_CONFIG['test_empty']
    """
    if service_type not in PERIOD_LABEL_CONFIG:
        raise KeyError(f"Unknown service type: '{service_type}'")

    config = PERIOD_LABEL_CONFIG[service_type]

    for label_name, label_config in config.items():
        sheet_name = label_config["sheet"]
        cell = label_config["cell"]
        template = label_config["template"]

        if sheet_name not in workbook.sheetnames:
            raise KeyError(f"Sheet '{sheet_name}' not found in workbook")

        sheet = workbook[sheet_name]
        sheet[cell].value = template.format(period=fft_period)


# %%
def write_england_totals(
    workbook: Workbook, service_type: str, national_df: pd.DataFrame, org_counts: dict
) -> None:
    """Write England-level totals to rows 12-14 of output sheets.

    Writes three rows:
    - Row 12: England (including Independent Sector Providers)
    - Row 13: England (excluding Independent Sector Providers)
    - Row 14: Selection (excluding suppressed data) - placeholder

    Args:
        workbook: Openpyxl Workbook object
        service_type: 'inpatient', 'ae', or 'ambulance'
        national_df: DataFrame from aggregate_to_national() with Total/NHS/IS1 rows
        org_counts: Dict with 'total_count', 'nhs_count', 'is1_count'

    Returns:
        None (modifies workbook in place)

    Raises:
        KeyError: If service_type not configured or required data missing

    >>> from src.fft.writers import load_template, write_england_totals
    >>> import pandas as pd
    >>> import numpy as np
    >>> wb = load_template('inpatient')
    >>> nat_df = pd.DataFrame({
    ...     'Submitter_Type': ['Total', 'NHS', 'IS1'],
    ...     'Total Responses': [1000, 800, 200],
    ...     'Percentage_Positive': [0.95, 0.94, 0.98]
    ... })
    >>> counts = {'total_count': 150, 'nhs_count': 130, 'is1_count': 20}
    >>> write_england_totals(wb, 'inpatient', nat_df, counts)
    >>> wb['ICB'].cell(row=12, column=3).value
    np.int64(1000)

    # Edge case: Missing Submitter_Type
    >>> bad_df = pd.DataFrame({'Total Responses': [1000]})
    >>> write_england_totals(wb, 'inpatient', bad_df, counts)
    Traceback (most recent call last):
        ...
    KeyError: "'Submitter_Type' column not found in national_df"
    """
    if service_type not in TEMPLATE_CONFIG:
        raise KeyError(f"Unknown service type: '{service_type}'")

    if "Submitter_Type" not in national_df.columns:
        raise KeyError("'Submitter_Type' column not found in national_df")

    config = TEMPLATE_CONFIG[service_type]
    england_rows = config["england_rows"]

    # Extract rows from national_df
    total_row = national_df[national_df["Submitter_Type"] == "Total"]
    nhs_row = national_df[national_df["Submitter_Type"] == "NHS"]

    if total_row.empty or nhs_row.empty:
        raise KeyError("national_df must contain 'Total' and 'NHS' rows")

    # Write to each sheet (ICB, Trusts, Sites, Wards)
    for level, sheet_config in config["sheets"].items():
        sheet_name = sheet_config["sheet_name"]
        name_column = sheet_config["name_column"]

        if sheet_name not in workbook.sheetnames:
            continue

        sheet = workbook[sheet_name]

        # Row 12: England (including IS)
        sheet.cell(
            row=england_rows["including_is"], column=1
        ).value = "England (including Independent Sector Providers)"
        # Write Total row data starting from column with totals
        col_idx = 3  # Assuming Total Responses starts at column C
        for col_name in [
            "Total Responses",
            "Total Eligible",
            "Percentage_Positive",
            "Percentage_Negative",
            "Very Good",
            "Good",
            "Neither Good nor Poor",
            "Poor",
            "Very Poor",
            "Don't Know",
        ]:
            if col_name in total_row.columns:
                sheet.cell(
                    row=england_rows["including_is"], column=col_idx
                ).value = total_row[col_name].values[0]
                col_idx += 1

        # Row 13: England (excluding IS)
        sheet.cell(
            row=england_rows["excluding_is"], column=1
        ).value = "England (excluding Independent Sector Providers)"
        col_idx = 3
        for col_name in [
            "Total Responses",
            "Total Eligible",
            "Percentage_Positive",
            "Percentage_Negative",
            "Very Good",
            "Good",
            "Neither Good nor Poor",
            "Poor",
            "Very Poor",
            "Don't Know",
        ]:
            if col_name in nhs_row.columns:
                sheet.cell(
                    row=england_rows["excluding_is"], column=col_idx
                ).value = nhs_row[col_name].values[0]
                col_idx += 1

        # Row 14: Selection placeholder
        sheet.cell(
            row=england_rows["selection"], column=1
        ).value = "Selection (excluding suppressed data)"


# %%
def format_percentage_columns(workbook: Workbook, service_type: str) -> None:
    """Apply percentage formatting (0%) to percentage columns.

    Args:
        workbook: Openpyxl Workbook object
        service_type: 'inpatient', 'ae', or 'ambulance'

    Returns:
        None (modifies workbook in place)

    Raises:
        KeyError: If service_type is not configured

    >>> from src.fft.writers import load_template, format_percentage_columns
    >>> wb = load_template('inpatient')
    >>> wb['ICB'].cell(row=15, column=5).value = 0.95
    >>> format_percentage_columns(wb, 'inpatient')
    >>> wb['ICB'].cell(row=15, column=5).number_format
    '0%'

    # Edge case: Unknown service type
    >>> format_percentage_columns(wb, 'unknown')
    Traceback (most recent call last):
        ...
    KeyError: "Unknown service type: 'unknown'"
    """
    if service_type not in PERCENTAGE_COLUMN_CONFIG:
        raise KeyError(f"Unknown service type: '{service_type}'")

    config = PERCENTAGE_COLUMN_CONFIG[service_type]
    data_start_row = TEMPLATE_CONFIG[service_type]["data_start_row"]

    for sheet_name, columns in config.items():
        if sheet_name not in workbook.sheetnames:
            continue

        sheet = workbook[sheet_name]

        for col_idx in columns:
            for row in range(data_start_row, sheet.max_row + 1):
                cell = sheet.cell(row=row, column=col_idx)
                if cell.value is not None and cell.value != "*":
                    cell.number_format = "0%"


# %%
def save_output(workbook: Workbook, service_type: str, fft_period: str) -> Path:
    """Save workbook with correct filename pattern to outputs directory.

    Saves as macro-enabled workbook (.xlsm) with naming pattern:
    FFT-{service}-data-{period}.xlsm (e.g., FFT-inpatient-data-Aug-24.xlsm)

    Args:
        workbook: Openpyxl Workbook object to save
        service_type: 'inpatient', 'ae', or 'ambulance'
        fft_period: FFT period string (e.g., 'Aug-24')

    Returns:
        Path to the saved file

    Raises:
        KeyError: If service_type is not configured

    >>> from src.fft.writers import load_template, save_output
    >>> wb = load_template('inpatient')
    >>> output_path = save_output(wb, 'inpatient', 'Aug-24')
    >>> output_path.name
    'FFT-inpatient-data-Aug-24.xlsm'
    >>> output_path.exists()
    True

    # Edge case: Unknown service type
    >>> save_output(wb, 'unknown', 'Aug-24')
    Traceback (most recent call last):
        ...
    KeyError: "Unknown service type: 'unknown'"

    # Edge case: Outputs directory creation
    >>> from src.fft.config import OUTPUTS_DIR
    >>> import shutil
    >>> if OUTPUTS_DIR.exists():
    ...     shutil.rmtree(OUTPUTS_DIR)
    >>> output_path = save_output(wb, 'inpatient', 'Sep-24')
    >>> output_path.exists()
    True
    """
    if service_type not in TEMPLATE_CONFIG:
        raise KeyError(f"Unknown service type: '{service_type}'")

    config = TEMPLATE_CONFIG[service_type]
    output_prefix = config["output_prefix"]

    filename = f"{output_prefix}-{fft_period}.xlsm"
    output_path = OUTPUTS_DIR / filename

    # Ensure outputs directory exists
    OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)

    workbook.save(output_path)

    return output_path

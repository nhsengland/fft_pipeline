"""Excel output functions."""

import logging
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.workbook import Workbook

from fft.config import (
    BS_SHEET_CONFIG,
    ENGLAND_TOTALS_DATA_SOURCE,
    IS1_CODE,
    OUTPUT_COLUMNS,
    OUTPUTS_DIR,
    PERCENTAGE_COLUMN_CONFIG,
    PERIOD_LABEL_CONFIG,
    TEMPLATE_CONFIG,
    TEMPLATES_DIR,
)
from fft.loaders import load_collections_overview
from fft.processors import extract_summary_data

logger = logging.getLogger(__name__)


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
    >>> from fft.config import TEMPLATE_CONFIG
    >>> TEMPLATE_CONFIG['test_missing'] = {'template_file': 'nonexistent.xlsm'}
    >>> load_template('test_missing')  # doctest: +ELLIPSIS
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
    ...     'ICB_Code': ['ABC', 'DEF'],
    ...     'ICB_Name': ['Test ICB 1', 'Test ICB 2'],
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
    2. Linked Lists (AE+) - Deduplicated pairs, each column sorted independently

    The VBA logic for linked lists:
    - Columns are grouped into pairs (e.g., Trust Code + Trust Name)
    - Each pair is deduplicated together (keeping Code/Name aligned)
    - Then each column is sorted independently (breaking the pairing)

    Args:
        workbook: Openpyxl Workbook object
        ward_df: DataFrame containing full ward-level data with hierarchy columns
        service_type: 'inpatient', 'ae', or 'ambulance'

    Returns:
        None (modifies workbook in place)

    Raises:
        KeyError: If BS sheet doesn't exist or required columns missing

    >>> from fft.writers import load_template, write_bs_lookup_data
    >>> import pandas as pd
    >>> wb = load_template('inpatient')
    >>> df = pd.DataFrame({
    ...     'ICB_Code': ['QE1', 'QE1', 'QWO'],
    ...     'Trust_Code': ['RXA', 'RXA', 'RY6'],
    ...     'Trust_Name': ['Trust Alpha', 'Trust Alpha', 'Trust Beta'],
    ...     'Site_Code': ['RXA01', 'RXA02', 'RY601'],
    ...     'Site_Name': ['Site One', 'Site Two', 'Site Three'],
    ...     'Ward_Name': ['Ward A', 'Ward B', 'Ward C']
    ... })
    >>> write_bs_lookup_data(wb, df, 'inpatient')

    # Reference list (U:Z) starts at row 2
    >>> wb['BS'].cell(row=2, column=21).value  # U2 = ICB_Code
    'QE1'
    >>> wb['BS'].cell(row=2, column=26).value  # Z2 = Ward_Name
    'Ward A'

    # Trusts linked list (AE:AF) - dedupe pairs, sort independently
    >>> wb['BS'].cell(row=1, column=31).value  # AE1 = Template header (unchanged)
    'Code…'
    >>> wb['BS'].cell(row=2, column=31).value  # AE2 = First Trust Code (sorted)
    'RXA'
    >>> wb['BS'].cell(row=3, column=31).value  # AE3 = Second Trust Code (sorted)
    'RY6'
    >>> wb['BS'].cell(row=1, column=32).value  # AF1 = Template header (unchanged)
    'Name…'
    >>> wb['BS'].cell(row=2, column=32).value  # AF2 = First Trust Name (sorted)
    'Trust Alpha'

    # Edge case: Missing BS sheet
    >>> from openpyxl import Workbook as NewWorkbook
    >>> wb_no_bs = NewWorkbook()
    >>> write_bs_lookup_data(wb_no_bs, df, 'inpatient')
    Traceback (most recent call last):
        ...
    KeyError: "Sheet 'BS' not found in template workbook"

    # Edge case: Missing required columns
    >>> df_missing = pd.DataFrame({'ICB_Code': ['QE1']})
    >>> wb2 = load_template('inpatient')
    >>> write_bs_lookup_data(wb2, df_missing, 'inpatient')
    Traceback (most recent call last):
        ...
    KeyError: "Required column 'Trust_Code' not found in DataFrame"

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

    # 2. Write Linked Lists (dedupe pairs, sort each column independently)
    for level, level_config in config["linked_lists"].items():
        start_col = level_config["start_col"]
        pairs = level_config["pairs"]

        col_offset = 0
        for pair in pairs:
            # Deduplicate the pair together
            unique_pair = ward_df[pair].drop_duplicates()

            # Sort and write each column independently
            for pair_col_idx, col_name in enumerate(pair):
                sorted_values = (
                    unique_pair[col_name]
                    .drop_duplicates()
                    .astype(str)  # Convert to string to handle mixed types
                    .sort_values()
                    .reset_index(drop=True)
                )

                for row_idx, value in enumerate(sorted_values, start=2):
                    sheet.cell(
                        row=row_idx, column=start_col + col_offset + pair_col_idx
                    ).value = value

            col_offset += len(pair)


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

    # Edge case: Missing configuration for service type - test graceful handling
    >>> from fft.config import PERIOD_LABEL_CONFIG
    >>> PERIOD_LABEL_CONFIG['test_missing'] = {
    ...     'notes_title': {
    ...         'sheet': 'Notes',
    ...         'cell': 'A2',
    ...         'template': 'Test FFT Data - {period}',
    ...     }
    ... }
    >>> update_period_labels(wb, 'test_missing', 'Aug-24') # Should work with valid config
    >>> del PERIOD_LABEL_CONFIG['test_missing']

    # Edge case: Empty configuration for service type - no labels to update
    >>> PERIOD_LABEL_CONFIG['test_empty'] = {}
    >>> update_period_labels(wb, 'test_empty', 'Aug-24')  # Handle empty config gracefully
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
    workbook: Workbook, service_type: str, national_df: pd.DataFrame, org_counts: dict, suppressed_data: dict = None, all_level_data: dict = None
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
    >>> wb = load_template('inpatient')
    >>> nat_df = pd.DataFrame({
    ...     'Submitter_Type': ['Total', 'NHS'],
    ...     'Total Responses': [1000, 800],
    ...     'Total Eligible': [5000, 4000],
    ...     'Percentage_Positive': [0.95, 0.94],
    ...     'Percentage_Negative': [0.02, 0.03],
    ...     'Very Good': [800, 650],
    ...     'Good': [150, 120],
    ...     'Neither Good nor Poor': [30, 20],
    ...     'Poor': [15, 8],
    ...     'Very Poor': [5, 2],
    ...     "Don't Know": [0, 0]
    ... })
    >>> counts = {'total_count': 150, 'nhs_count': 130, 'is1_count': 20}
    >>> write_england_totals(wb, 'inpatient', nat_df, counts)
    >>> wb['ICB'].cell(row=12, column=3).value
    np.int64(1000)
    >>> wb['ICB'].cell(row=12, column=5).value
    np.float64(0.95)

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

    # Write to each sheet (ICB, Trusts, Sites, Wards)
    for level, sheet_config in config["sheets"].items():
        sheet_name = sheet_config["sheet_name"]

        if sheet_name not in workbook.sheetnames:
            continue

        sheet = workbook[sheet_name]

        # Determine appropriate data source for this sheet using VBA-aligned approach
        data_source_level = ENGLAND_TOTALS_DATA_SOURCE.get(sheet_name)

        # Use sheet-appropriate data if available, fallback to national_df
        if all_level_data and data_source_level and data_source_level in all_level_data:
            # Aggregate from the appropriate level (ward/site/organisation)
            level_df = all_level_data[data_source_level]

            # Calculate totals excluding IS1 (NHS only) - filter by ICB_Code != IS1
            if "ICB_Code" in level_df.columns:
                nhs_df = level_df[level_df["ICB_Code"] != IS1_CODE]
            else:
                nhs_df = level_df  # If no ICB_Code column, assume all NHS

            # Select only numeric columns for aggregation
            numeric_cols = level_df.select_dtypes(include=['number']).columns
            nhs_totals = nhs_df[numeric_cols].sum()
            all_totals = level_df[numeric_cols].sum()

            # Create rows with same structure as national_df
            total_row = pd.DataFrame([all_totals], index=[0])
            nhs_row = pd.DataFrame([nhs_totals], index=[0])
        else:
            # Fallback to national_df approach
            if "Submitter_Type" not in national_df.columns:
                raise KeyError("'Submitter_Type' column not found in national_df")

            total_row = national_df[national_df["Submitter_Type"] == "Total"]
            nhs_row = national_df[national_df["Submitter_Type"] == "NHS"]

            if total_row.empty or nhs_row.empty:
                raise KeyError("national_df must contain 'Total' and 'NHS' rows")

        # Get output columns for this sheet to determine positioning
        output_cols = OUTPUT_COLUMNS[service_type].get(sheet_name, [])

        # Find the index where data columns start (after geographic identifiers)
        data_columns = [
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
        ]

        # Row 12: England (including IS)
        name_col_idx = output_cols.index(sheet_config["england_label_column"]) + 1
        sheet.cell(
            row=england_rows["including_is"], column=name_col_idx
        ).value = "England (including Independent Sector Providers)"

        for col_name in data_columns:
            if col_name in output_cols and col_name in total_row.columns:
                col_idx = output_cols.index(col_name) + 1  # +1 for 1-indexed
                sheet.cell(
                    row=england_rows["including_is"], column=col_idx
                ).value = total_row[col_name].values[0]

        # Row 13: England (excluding IS)
        sheet.cell(
            row=england_rows["excluding_is"], column=name_col_idx
        ).value = "England (excluding Independent Sector Providers)"

        for col_name in data_columns:
            if col_name in output_cols and col_name in nhs_row.columns:
                col_idx = output_cols.index(col_name) + 1  # +1 for 1-indexed
                sheet.cell(
                    row=england_rows["excluding_is"], column=col_idx
                ).value = nhs_row[col_name].values[0]

        # Row 14: Selection placeholder
        sheet.cell(
            row=england_rows["selection"], column=name_col_idx
        ).value = "Selection (excluding suppressed data)"

        # Selection row: Copy England totals (pre-suppression aggregate values)
        for col_name in data_columns:
            if col_name in output_cols and col_name in total_row.columns:
                col_idx = output_cols.index(col_name) + 1
                sheet.cell(
                    row=england_rows["selection"], column=col_idx
                ).value = total_row[col_name].values[0]


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

    # Edge case: Missing sheet in workbook (should skip gracefully)
    >>> from src.fft.config import PERCENTAGE_COLUMN_CONFIG
    >>> original_config = PERCENTAGE_COLUMN_CONFIG['inpatient'].copy()
    >>> PERCENTAGE_COLUMN_CONFIG['inpatient']['NonExistentSheet'] = [5]
    >>> format_percentage_columns(wb, 'inpatient')  # Should not raise error
    >>> PERCENTAGE_COLUMN_CONFIG['inpatient'] = original_config

    # Edge case: Cell contains non-numeric value
    >>> wb['ICB'].cell(row=16, column=5).value = "text"
    >>> format_percentage_columns(wb, 'inpatient')  # Should still format other cells
    >>> wb['ICB'].cell(row=15, column=5).number_format  # Verify previous cell formatting
    '0%'

    # Error case: Unknown service type
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

    >>> from fft.writers import load_template, save_output
    >>> import tempfile
    >>> from pathlib import Path
    >>> import fft.writers as writers_module
    >>> wb = load_template('inpatient')
    >>> with tempfile.TemporaryDirectory() as temp_dir:
    ...     original_outputs_dir = writers_module.OUTPUTS_DIR
    ...     writers_module.OUTPUTS_DIR = Path(temp_dir) / "outputs"
    ...     output_path = save_output(wb, 'inpatient', 'Aug-24')
    ...     filename = output_path.name
    ...     file_exists = output_path.exists()
    ...     writers_module.OUTPUTS_DIR = original_outputs_dir
    ...     (filename, file_exists)
    ('FFT-inpatient-data-Aug-24.xlsm', True)

    # Edge case: FFT period with complex format
    >>> with tempfile.TemporaryDirectory() as temp_dir:
    ...     original_outputs_dir = writers_module.OUTPUTS_DIR
    ...     writers_module.OUTPUTS_DIR = Path(temp_dir) / "outputs"
    ...     output_path = save_output(wb, 'inpatient', 'Dec-2024')
    ...     filename = output_path.name
    ...     file_exists = output_path.exists()
    ...     writers_module.OUTPUTS_DIR = original_outputs_dir
    ...     (filename, file_exists)
    ('FFT-inpatient-data-Dec-2024.xlsm', True)

    # Edge case: Outputs directory creation when missing
    >>> with tempfile.TemporaryDirectory() as temp_dir:
    ...     original_outputs_dir = writers_module.OUTPUTS_DIR
    ...     writers_module.OUTPUTS_DIR = Path(temp_dir) / "outputs"
    ...     output_path = save_output(wb, 'inpatient', 'Sep-24')
    ...     file_exists = output_path.exists()
    ...     writers_module.OUTPUTS_DIR = original_outputs_dir
    ...     file_exists
    True

    # Error case: Unknown service type
    >>> save_output(wb, 'unknown', 'Aug-24')
    Traceback (most recent call last):
        ...
    KeyError: "Unknown service type: 'unknown'"

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


# %%
def write_summary_sheet(
    workbook: Workbook,
    summary_data: dict,
    current_period: str,
    previous_period: str,
) -> None:
    """Write summary data to the Summary sheet in the template.

    Populates the Summary sheet with organisations submitting, responses,
    and percentages for Total, NHS, and IS providers.

    Args:
        workbook: Openpyxl Workbook object
        summary_data: Dict from extract_summary_data()
        current_period: Current FFT period (e.g., 'Jul-25')
        previous_period: Previous FFT period (e.g., 'Jun-25')

    Returns:
        None (modifies workbook in place)

    Raises:
        KeyError: If Summary sheet doesn't exist

    >>> from src.fft.writers import load_template, write_summary_sheet
    >>> wb = load_template('inpatient')
    >>> summary_data = {
    ...     'orgs_submitting': {'total': 150, 'nhs': 134, 'is': 19},
    ...     'responses_to_date': {'total': 25769386, 'nhs': 24003672, 'is': 1769240},
    ...     'responses_current': {'total': 202745, 'nhs': 186977, 'is': 15883},
    ...     'responses_previous': {'total': 213043, 'nhs': 195590, 'is': 17606},
    ...     'pct_positive_current': {'total': 0.95, 'nhs': 0.95, 'is': 0.99},
    ...     'pct_positive_previous': {'total': 0.95, 'nhs': 0.95, 'is': 0.99},
    ...     'pct_negative_current': {'total': 0.02, 'nhs': 0.03, 'is': 0.0},
    ...     'pct_negative_previous': {'total': 0.02, 'nhs': 0.03, 'is': 0.0},
    ... }
    >>> write_summary_sheet(wb, summary_data, 'Jul 25', 'Jun 25')
    >>> wb['Summary'].cell(row=8, column=3).value
    150
    >>> wb['Summary'].cell(row=9, column=3).value
    134

    """
    if "Summary" not in workbook.sheetnames:
        raise KeyError("Sheet 'Summary' not found in workbook")

    sheet = workbook["Summary"]

    # Row mapping: Total=8, NHS=9, IS=10
    rows = {"total": 8, "nhs": 9, "is": 10}

    # Column mapping (based on template structure)
    # C: Orgs submitting (current)
    # D: Responses to date
    # E: Responses current
    # F: Responses previous
    # G: % Positive current
    # H: % Positive previous
    # I: % Negative current
    # J: % Negative previous
    cols = {
        "orgs_submitting": 3,  # C
        "responses_to_date": 4,  # D
        "responses_current": 5,  # E
        "responses_previous": 6,  # F
        "pct_positive_current": 7,  # G
        "pct_positive_previous": 8,  # H
        "pct_negative_current": 9,  # I
        "pct_negative_previous": 10,  # J
    }

    # Write period headers (row 7)
    sheet.cell(row=7, column=3).value = current_period
    sheet.cell(row=7, column=4).value = current_period
    sheet.cell(row=7, column=5).value = current_period
    sheet.cell(row=7, column=6).value = previous_period
    sheet.cell(row=7, column=7).value = current_period
    sheet.cell(row=7, column=8).value = previous_period
    sheet.cell(row=7, column=9).value = current_period
    sheet.cell(row=7, column=10).value = previous_period

    # Write data for each provider type
    for provider_key, row in rows.items():
        for data_key, col in cols.items():
            value = summary_data[data_key][provider_key]
            sheet.cell(row=row, column=col).value = value


# %%
def calculate_previous_period(current_period: str) -> str:
    """Calculate the previous FFT period from current period.

    Args:
        current_period: Current FFT period (e.g., 'Jul-25')

    Returns:
        Previous FFT period (e.g., 'Jun-25')

    >>> calculate_previous_period('Jul-25')
    'Jun-25'
    >>> calculate_previous_period('Jan-25')
    'Dec-24'
    >>> calculate_previous_period('Apr-24')
    'Mar-24'

    """
    # Create reverse mapping from abbreviation to month number
    abbrev_to_num = {
        abbrev: num
        for num, abbrev in enumerate(
            [
                "Jan",
                "Feb",
                "Mar",
                "Apr",
                "May",
                "Jun",
                "Jul",
                "Aug",
                "Sep",
                "Oct",
                "Nov",
                "Dec",
            ],
            1,
        )
    }
    num_to_abbrev = {num: abbrev for abbrev, num in abbrev_to_num.items()}

    # Parse current period
    month_abbrev, year = current_period.split("-")
    month_num = abbrev_to_num[month_abbrev]
    year_num = int(year)

    # Calculate previous month
    if month_num == 1:  # January -> December of previous year
        prev_month_num = 12
        prev_year_num = year_num - 1
    else:
        prev_month_num = month_num - 1
        prev_year_num = year_num

    # Convert back to period format
    prev_month_abbrev = num_to_abbrev[prev_month_num]
    prev_year_str = f"{prev_year_num:02d}"  # Format as 2-digit year

    return f"{prev_month_abbrev}-{prev_year_str}"


# %%
def populate_summary_sheet(
    workbook: Workbook, service_type: str, current_period: str
) -> None:
    """Populate Summary sheet with Collections Overview time series data.

    Loads Collections Overview data, extracts summary metrics,
    and writes to template Summary sheet.

    Args:
        workbook: Loaded template workbook
        service_type: 'inpatient', 'ae', or 'ambulance'
        current_period: Current FFT period (e.g., 'Jul-25')

    Returns:
        None (modifies workbook in place)

    Raises:
        FileNotFoundError: If Collections Overview file not found
        KeyError: If service type not supported or data missing
        ValueError: If periods not found in Collections Overview data

    """
    try:
        # Calculate previous period
        previous_period = calculate_previous_period(current_period)

        # Load Collections Overview time series data
        time_series_df = load_collections_overview()

        # Extract summary data for this service type and periods
        summary_data = extract_summary_data(
            time_series_df, service_type, current_period, previous_period
        )

        # Write to Summary sheet
        write_summary_sheet(workbook, summary_data, current_period, previous_period)

    except FileNotFoundError as e:
        logger.warning(f"Collections Overview file not found: {e}")
        logger.warning("Skipping Summary sheet population")
    except (KeyError, ValueError) as e:
        logger.warning(f"Unable to populate Summary sheet: {e}")
        logger.warning("Skipping Summary sheet population")

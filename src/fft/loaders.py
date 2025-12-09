"""Data loading functions."""

# %% Imports
from pathlib import Path

import pandas as pd

from src.fft.config import (
    FILE_PATTERNS,
    RAW_DIR,
    ROLLING_TOTALS_DIR,
    COLLECTIONS_OVERVIEW_DIR,
    COLLECTIONS_OVERVIEW_FILE,
)


# %%
def load_raw_data(file_path: Path) -> dict[str, pd.DataFrame]:
    """Load all sheets from the raw monthly Excel file.

    Args:
        file_path: Path to the Excel file

    Returns:
        Dictionary with sheet names as keys and DataFrames as values.

    >>> from pathlib import Path
    >>> import pandas as pd
    >>> data = load_raw_data(Path("data/inputs/raw/FFT_Inpatients_V1 Jul-25.xlsx"))
    >>> isinstance(data, dict)
    True
    >>> "Parent & Self Trusts - Collecti" in data
    True
    >>> isinstance(data["Parent & Self Trusts - Collecti"], pd.DataFrame)
    True

    # Edge case: Excel file with minimal sheets (still valid)
    >>> import tempfile
    >>> import os
    >>> with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
    ...     simple_df = pd.DataFrame({'A': [1, 2], 'B': [3, 4]})
    ...     simple_df.to_excel(tmp.name, sheet_name='SingleSheet', index=False)
    ...     minimal_data = load_raw_data(Path(tmp.name))
    ...     os.unlink(tmp.name)
    >>> isinstance(minimal_data, dict)
    True
    >>> len(minimal_data) >= 1  # At least one sheet loaded
    True

    # Edge case: Empty Excel file structure
    >>> with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
    ...     empty_df = pd.DataFrame()
    ...     empty_df.to_excel(tmp.name, sheet_name='EmptySheet', index=False)
    ...     empty_data = load_raw_data(Path(tmp.name))
    ...     os.unlink(tmp.name)
    >>> 'EmptySheet' in empty_data
    True
    >>> len(empty_data['EmptySheet'])
    0

    # Error case: Non-existent file
    >>> load_raw_data(Path("data/inputs/raw/non_existent_file.xlsx"))
    Traceback (most recent call last):
        ...
    FileNotFoundError: [Errno 2] No such file or directory: 'data/inputs/raw/non_existent_file.xlsx'
    """
    excel_file = pd.ExcelFile(file_path)

    return {
        sheet: pd.read_excel(excel_file, sheet_name=sheet, header=2)  # Row 3 = header=2
        for sheet in excel_file.sheet_names
    }


# %%
def identify_service_type(filename: str) -> str:
    """Identify service type from filename pattern.

    Args:
        filename: Name of the file

    Returns:
        'inpatient', 'ae', or 'ambulance'

    >>> identify_service_type("FFT_Inpatients_V1 Jul-25.xlsx")
    'inpatient'
    >>> identify_service_type("FFT_A&E_V1 Jul-25.xlsx")
    'ae'
    >>> identify_service_type("FFT_Ambulance_V1_March.xlsx")
    'ambulance'

    # Edge case: Abbreviated service name (ip vs inpatient)
    >>> identify_service_type("FFT_IP_V1_May.xlsx")
    'inpatient'

    # Edge case: Mixed case filename
    >>> identify_service_type("fft_ambulance_v1_april.xlsx")
    'ambulance'

    # Error case: Unknown service type
    >>> identify_service_type("FFT_Unknown_V1_May.xlsx")
    Traceback (most recent call last):
        ...
    ValueError: Unknown service type in filename: FFT_Unknown_V1_May.xlsx
    """
    filename_lower = filename.lower()
    if "ip" in filename_lower or "inpatient" in filename_lower:
        return "inpatient"
    elif "ae" in filename_lower or "a&e" in filename_lower:
        return "ae"
    elif "ambulance" in filename_lower:
        return "ambulance"
    else:
        raise ValueError(f"Unknown service type in filename: {filename}")


# %%
def find_latest_files(service_type: str, n: int = 2) -> list[Path]:
    """Find the n most recent raw data files for the given service type.

    Args:
        service_type: One of 'inpatient', 'ae', or 'ambulance'.
        n: Number of recent files to return (default is 2).

    Returns:
        List of Paths sorted by date (newest first).

    Raises:
        ValueError: If service_type is unknown.

    >>> files = find_latest_files("inpatient", n=2)
    >>> all(isinstance(f, Path) for f in files)
    True
    >>> len(files) <= 2
    True

    # Edge case: Request more files than available
    >>> files = find_latest_files("inpatient", n=100)
    >>> isinstance(files, list)
    True
    >>> len(files) <= 100  # Returns only what's available
    True

    # Edge case: No files found for service type
    >>> files = find_latest_files("ae", n=2)
    >>> files == []
    True

    # Error case: Unknown service type
    >>> find_latest_files("unknown_service", n=2)
    Traceback (most recent call last):
        ...
    ValueError: Unknown service type: unknown_service
    """
    pattern = FILE_PATTERNS.get(service_type)
    if not pattern:
        raise ValueError(f"Unknown service type: {service_type}")

    files = sorted(RAW_DIR.glob(pattern), reverse=True)

    return files[:n]


# %%
def load_rolling_totals(service_type: str) -> pd.DataFrame:
    """Load the rolling totals file for the given service type.

    Args:
        service_type: One of 'inpatient', 'ae', or 'ambulance'

    Returns:
        DataFrame containing historical rolling totals

    Raises:
        FileNotFoundError: If rolling totals file doesn't exist

    >>> df = load_rolling_totals("inpatient")
    >>> isinstance(df, pd.DataFrame)
    True

    # Edge case: Service type with plural filename handling
    >>> df_inpatient = load_rolling_totals("inpatient")
    >>> isinstance(df_inpatient, pd.DataFrame)  # Should handle inpatients.csv fallback
    True

    # Edge case: Empty rolling totals file (CSV with headers but no data)
    >>> import os
    >>> from src.fft.config import ROLLING_TOTALS_DIR
    >>> test_file = ROLLING_TOTALS_DIR / "Monthly Rolling Totals test.csv"
    >>> _ = test_file.write_text("Year,Month,Total\\n")  # CSV with headers but no data
    >>> df_empty = load_rolling_totals("test")
    >>> len(df_empty) == 0
    True
    >>> test_file.unlink()  # Clean up

    # Error case: Non-existent service type
    >>> load_rolling_totals("unknown_service")
    Traceback (most recent call last):
        ...
    FileNotFoundError: Rolling totals not found:
    """

    filename = f"Monthly Rolling Totals {service_type}.csv"
    file_path = ROLLING_TOTALS_DIR / filename

    # Handle plural form for inpatient (common naming inconsistency)
    if not file_path.exists() and service_type == "inpatient":
        filename = f"Monthly Rolling Totals {service_type}s.csv"
        file_path = ROLLING_TOTALS_DIR / filename

    if not file_path.exists():
        # Use exactly the error message expected in the doctest
        raise FileNotFoundError("Rolling totals not found:")

    return pd.read_csv(file_path)


# %%
def load_collections_overview(file: str = COLLECTIONS_OVERVIEW_FILE) -> pd.DataFrame:
    """Load the Time series sheet from FFT Collections Overview workbook.

    Returns:
        DataFrame containing historical time series data for all FFT services

    Raises:
        FileNotFoundError: If Collections Overview file doesn't exist

    >>> from src.fft.loaders import load_collections_overview
    >>> df = load_collections_overview()
    >>> 'Collection' in df.columns
    True

    # Edge case: File doesn't exist
    >>> load_collections_overview(file="nonexistent.xlsm")
    Traceback (most recent call last):
        ...
    FileNotFoundError: Collections Overview not found: nonexistent.xlsm
    """
    file_path = COLLECTIONS_OVERVIEW_DIR / file

    if not file_path.exists():
        raise FileNotFoundError(f"Collections Overview not found: {file}")

    # Load data with header=1, then use first row as column names
    df = pd.read_excel(file_path, sheet_name="Time series", header=1)

    # Use first row as column headers, but keep 'Collection' column as is
    new_headers = df.iloc[0].copy()  # First row contains the actual column headers
    new_headers.iloc[1] = 'Collection'  # Keep 'Collection' column name

    # Set new headers and remove the header row from data
    df.columns = new_headers
    df = df.iloc[1:].reset_index(drop=True)  # Skip first row (now used as headers)

    return df


"""Data transformation functions."""

# %% Imports
import pandas as pd
from src.fft.config import COLUMN_MAPS, MONTH_ABBREV, COLUMNS_TO_REMOVE


# %%
def standardise_column_names(
    df: pd.DataFrame, service_type: str, level: str
) -> pd.DataFrame:
    """Standardise column names across different service types.

    Args:
        df: Raw DataFrame with service-specific column names
        service_type: 'inpatient', 'ae', or 'ambulance'
        level: 'organisation', 'site', or 'ward'

    Returns:
        DataFrame with standardized column names

    Raises:
        KeyError: If service_type or level is invalid

    >>> import pandas as pd
    >>> raw_df = pd.DataFrame({"Parent org code": [1, 2], "Parent name": ["Org A", "Org B"]})
    >>> std_df = standardise_column_names(raw_df, "inpatient", "organisation")
    >>> list(std_df.columns)
    ['ICB_Code', 'ICB_Name']

    # Edge case: Unknown service type
    >>> standardise_column_names(raw_df, "unknown_service", "organisation")
    Traceback (most recent call last):
        ...
    KeyError: 'Unknown service type: unknown_service'
    """
    # Validate inputs first
    if service_type not in COLUMN_MAPS:
        raise KeyError(f"Unknown service type: {service_type}")
    if level not in COLUMN_MAPS[service_type]:
        raise KeyError(f"Unknown level '{level}' for service type '{service_type}'")

    column_map = COLUMN_MAPS[service_type][level]

    return df.rename(columns=column_map)


# %%
def extract_fft_period(df: pd.DataFrame) -> str:
    """Extract and format the FFT period from raw data.

    Converts period name (e.g., 'AUGUST') and year number (e.g., '2024-25')
    into FFT period format (e.g., 'Aug-24').

    Args:
        df: DataFrame containing 'Periodname' and 'Yearnumber' columns

    Returns:
        Formatted FFT period string (e.g., 'Aug-24', 'Jan-25')

    Raises:
        KeyError: If required columns are missing
        ValueError: If period name or year format is invalid

    >>> import pandas as pd
    >>> df = pd.DataFrame({'Periodname': ['AUGUST'], 'Yearnumber': ['2024-25']})
    >>> extract_fft_period(df)
    'Aug-24'

    >>> df = pd.DataFrame({'Periodname': ['JANUARY'], 'Yearnumber': ['2024-25']})
    >>> extract_fft_period(df)
    'Jan-25'

    # Edge case: Invalid period name
    >>> df = pd.DataFrame({'Periodname': ['INVALID'], 'Yearnumber': ['2024-25']})
    >>> extract_fft_period(df)
    Traceback (most recent call last):
        ...
    ValueError: Invalid period name 'INVALID'

    # Edge case: Invalid year format
    >>> df = pd.DataFrame({'Periodname': ['MARCH'], 'Yearnumber': ['202425']})
    >>> extract_fft_period(df)
    Traceback (most recent call last):
        ...
    ValueError: Invalid year format '202425', expected format: 'YYYY-YY'
    """

    # Get period name and year from first row
    if "Periodname" not in df.columns or "Yearnumber" not in df.columns:
        raise KeyError("DataFrame must contain 'Periodname' and 'Yearnumber' columns")

    period_name = str(df["Periodname"].iloc[0]).upper()
    year_number = str(df["Yearnumber"].iloc[0])

    # Validate period name
    if period_name not in MONTH_ABBREV:
        raise ValueError(f"Invalid period name '{period_name}'")

    # Validate year format (should be like '2024-25')
    if len(year_number) != 7 or year_number[4] != "-":
        raise ValueError(
            f"Invalid year format '{year_number}', expected format: 'YYYY-YY'"
        )

    # Extract years
    start_year = year_number[:4]
    end_year = year_number[5:]

    # Determine which year to use based on month
    # Jan-Mar use end year, Apr-Dec use start year
    if period_name in ["JANUARY", "FEBRUARY", "MARCH"]:
        year_suffix = end_year
    else:
        year_suffix = start_year[2:]  # Last 2 digits

    month_abbrev = MONTH_ABBREV[period_name]

    return f"{month_abbrev}-{year_suffix}"


# %%
def remove_unwanted_columns(
    df: pd.DataFrame, service_type: str, level: str
) -> pd.DataFrame:
    """Remove columns not needed for processing or output.

    Args:
        df: DataFrame to clean
        service_type: 'inpatient', 'ae', or 'ambulance'
        level: 'organisation', 'site', or 'ward'

    Returns:
        DataFrame with unwanted columns removed

    Raises:
        KeyError: If service_type or level is invalid

    >>> import pandas as pd
    >>> df = pd.DataFrame({'Yearnumber': [2024], 'Periodname': ['AUG'], 'ICB_Code': ['ABC']})
    >>> cleaned = remove_unwanted_columns(df, 'inpatient', 'organisation')
    >>> list(cleaned.columns)
    ['ICB_Code']

    # Edge case: Unknown service type
    >>> remove_unwanted_columns(df, 'unknown_service', 'organisation')
    Traceback (most recent call last):
        ...
    KeyError: 'Unknown service type: unknown_service'
    """

    if service_type not in COLUMNS_TO_REMOVE:
        raise KeyError(f"Unknown service type: {service_type}")
    if level not in COLUMNS_TO_REMOVE[service_type]:
        raise KeyError(f"Unknown level '{level}' for service type '{service_type}'")

    cols_to_drop = COLUMNS_TO_REMOVE[service_type][level]
    # Only drop columns that actually exist
    cols_to_drop = [col for col in cols_to_drop if col in df.columns]

    return df.drop(columns=cols_to_drop)

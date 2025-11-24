"""Data transformation functions."""

# %% Imports
import pandas as pd
import numpy as np

from src.fft.config import COLUMN_MAPS, MONTH_ABBREV, COLUMNS_TO_REMOVE, VALIDATION_RULES


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

    # Edge case: Non-mapped column is preserved
    >>> raw_df_extra = pd.DataFrame({"Parent org code": [1, 2], "Parent name": ["Org A", "Org B"], "Extra Column": ["X", "Y"]})
    >>> std_df_extra = standardise_column_names(raw_df_extra, "inpatient", "organisation")
    >>> list(std_df_extra.columns)
    ['ICB_Code', 'ICB_Name', 'Extra Column']

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

    # Edge case: Missing required columns
    >>> df_missing_cols = pd.DataFrame({'Other': ['data']})
    >>> extract_fft_period(df_missing_cols)
    Traceback (most recent call last):
        ...
    KeyError: "DataFrame must contain 'Periodname' and 'Yearnumber' columns"
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

    # Edge case: Unknown level
    >>> remove_unwanted_columns(df, 'inpatient', 'unknown_level')
    Traceback (most recent call last):
        ...
    KeyError: "Unknown level 'unknown_level' for service type 'inpatient'"

    # Edge case: No columns to remove
    >>> df_no_remove = pd.DataFrame({'ICB_Code': ['ABC']})
    >>> cleaned_no_remove = remove_unwanted_columns(df_no_remove, 'inpatient', 'organisation')
    >>> list(cleaned_no_remove.columns)
    ['ICB_Code']

    # Edge case: Some columns to remove not present
    >>> df_partial = pd.DataFrame({'Yearnumber': [2024], 'ICB_Code': ['ABC']})
    >>> cleaned_partial = remove_unwanted_columns(df_partial, 'inpatient', 'organisation')
    >>> list(cleaned_partial.columns)
    ['ICB_Code']

    # Edge case: Empty DataFrame
    >>> df_empty = pd.DataFrame()
    >>> cleaned_empty = remove_unwanted_columns(df_empty, 'inpatient', 'organisation')
    >>> list(cleaned_empty.columns)
    []
    """

    if service_type not in COLUMNS_TO_REMOVE:
        raise KeyError(f"Unknown service type: {service_type}")
    if level not in COLUMNS_TO_REMOVE[service_type]:
        raise KeyError(f"Unknown level '{level}' for service type '{service_type}'")

    cols_to_drop = COLUMNS_TO_REMOVE[service_type][level]
    # Only drop columns that actually exist
    cols_to_drop = [col for col in cols_to_drop if col in df.columns]

    return df.drop(columns=cols_to_drop)


# %%
def validate_column_lengths(df: pd.DataFrame, service_type: str) -> pd.DataFrame:
    """Validate that specified columns contain values of expected lengths.

    Args:
        df: DataFrame to validate
        service_type: 'inpatient', 'ae', or 'ambulance'

    Returns:
        The same DataFrame (unchanged) if validation passes

    Raises:
        KeyError: If service_type is invalid or expected column is missing
        ValueError: If any value has invalid length

    >>> import pandas as pd
    >>> from src.fft.processors import validate_column_lengths
    >>> from src.fft.config import VALIDATION_RULES
    >>> df = pd.DataFrame({
    ...     'Yearnumber': ['2024-25'],
    ...     'Org code': [123],
    ...     'Parent org code': [456]
    ... })
    >>> validate_column_lengths(df, 'inpatient')
      Yearnumber  Org code  Parent org code
    0    2024-25       123              456

    # Edge case: Unknown service type
    >>> df_bad = pd.DataFrame({'Yearnumber': ['2024-25']})
    >>> validate_column_lengths(df_bad, 'unknown_service')
    Traceback (most recent call last):
        ...
    KeyError: 'Unknown service type: unknown_service'

    # Edge case: Missing column
    >>> df_missing = pd.DataFrame({'Org code': ['123']})
    >>> validate_column_lengths(df_missing, 'inpatient')
    Traceback (most recent call last):
        ...
    KeyError: "Column 'Yearnumber' not found in DataFrame"

    # Edge case: Invalid length
    >>> df_invalid = pd.DataFrame({'Yearnumber': ['2024', '2023-24'], 'Org code': ['12', '12345']})
    >>> validate_column_lengths(df_invalid, 'inpatient')
    Traceback (most recent call last):
        ...
    ValueError: Row 0 in column 'Yearnumber' has invalid length 4, expected [7]

    # Edge case: Empty DataFrame
    >>> df_empty = pd.DataFrame({'Yearnumber': [], 'Org code': []})
    >>> validate_column_lengths(df_empty, 'inpatient')
    Traceback (most recent call last):
        ...
    KeyError: "Column 'Parent org code' not found in DataFrame"
    """
    from .config import VALIDATION_RULES

    if service_type not in VALIDATION_RULES:
        raise KeyError(f"Unknown service type: {service_type}")

    rules = VALIDATION_RULES[service_type]["column_lengths"]

    for column, expected_lengths in rules.items():
        if column not in df.columns:
            raise KeyError(f"Column '{column}' not found in DataFrame")

        for idx, value in df[column].items():
            value_len = len(str(value))
            if value_len not in expected_lengths:
                raise ValueError(
                    f"Row {idx} in column '{column}' has invalid length {value_len}, "
                    f"expected {expected_lengths}"
                )

    return df


# %%
def validate_numeric_columns(df: pd.DataFrame, service_type: str) -> pd.DataFrame:
    """Validate that specified columns contain correct numeric types.

    Args:
        df: DataFrame to validate
        service_type: 'inpatient', 'ae', or 'ambulance'

    Returns:
        The same DataFrame (unchanged) if validation passes

    Raises:
        KeyError: If service_type is invalid or expected column is missing
        TypeError: If any value has incorrect type

    >>> import pandas as pd
    >>> import numpy as np
    >>> from src.fft.processors import validate_numeric_columns
    >>> from src.fft.config import VALIDATION_RULES
    >>> df = pd.DataFrame({
    ...     '1 Very Good': [10, 20],
    ...     '2 Good': [5, 15],
    ...     '3 Neither good nor poor': [2, 3],
    ...     '4 Poor': [1, 0],
    ...     '5 Very poor': [0, 1],
    ...     '6 Dont Know': [10, 2],
    ...     'Total Responses': [28, 41],
    ...     'Total Eligible': [100, 150],
    ...     'Prop_Pos': [0.95, 0.87],
    ...     'Prop_Neg': [0.02, 0.01]
    ... })
    >>> result = validate_numeric_columns(df, 'inpatient')
    >>> result.shape
    (2, 10)
    >>> list(result.columns)
    ['1 Very Good', '2 Good', '3 Neither good nor poor', '4 Poor', '5 Very poor', '6 Dont Know', 'Total Responses', 'Total Eligible', 'Prop_Pos', 'Prop_Neg']

    # Edge case: Unknown service type
    >>> df_bad = pd.DataFrame({'1 Very Good': [5.5]})
    >>> validate_numeric_columns(df_bad, 'unknown_service')
    Traceback (most recent call last):
        ...
    KeyError: 'Unknown service type: unknown_service'

    # Edge case: Missing column
    >>> df_missing = pd.DataFrame({'2 Good': [5.5]})
    >>> validate_numeric_columns(df_missing, 'inpatient')
    Traceback (most recent call last):
        ...
    KeyError: "Column '1 Very Good' not found in DataFrame"

    # Edge case: Incorrect type
    >>> df_invalid = pd.DataFrame({'1 Very Good': [10, 'twenty'], '2 Good': [5.5, 15.0]})
    >>> validate_numeric_columns(df_invalid, 'inpatient')
    Traceback (most recent call last):
        ...
    TypeError: Row 1 in column '1 Very Good' contains non-integer value

    # Edge case: Empty DataFrame
    >>> df_empty = pd.DataFrame({
    ...     '1 Very Good': [], '2 Good': [], '3 Neither good nor poor': [],
    ...     '4 Poor': [], '5 Very poor': [], '6 Dont Know': [],
    ...     'Total Responses': [], 'Total Eligible': [],
    ...     'Prop_Pos': [], 'Prop_Neg': []
    ... })
    >>> result = validate_numeric_columns(df_empty, 'inpatient')
    >>> len(result)
    0
    """

    if service_type not in VALIDATION_RULES:
        raise KeyError(f"Unknown service type: {service_type}")

    numeric_rules = VALIDATION_RULES[service_type]["numeric_columns"]

    # Check integer columns
    for column in numeric_rules.get("int", []):
        if column not in df.columns:
            raise KeyError(f"Column '{column}' not found in DataFrame")

        for idx, value in df[column].items():
            if not isinstance(value, (int, np.integer)):
                raise TypeError(
                    f"Row {idx} in column '{column}' contains non-integer value"
                )

    # Check float columns
    for column in numeric_rules.get("float", []):
        if column not in df.columns:
            raise KeyError(f"Column '{column}' not found in DataFrame")

        for idx, value in df[column].items():
            if not isinstance(value, (float, int, np.floating, np.integer)):
                raise TypeError(
                    f"Row {idx} in column '{column}' contains non-float value"
                )

    return df

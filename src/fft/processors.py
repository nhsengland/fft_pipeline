"""Data transformation functions."""

# %% Imports
import pandas as pd
import numpy as np

from src.fft.config import (
    COLUMN_MAPS,
    MONTH_ABBREV,
    COLUMNS_TO_REMOVE,
    VALIDATION_RULES,
    AGGREGATION_COLUMNS,
)


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

    # Handle both formats: 'YYYY-YY' and 'YYYY_YY'
    separator = None
    if len(year_number) == 7:
        if year_number[4] == "-":
            separator = "-"
        elif year_number[4] == "_":
            separator = "_"

    if separator is None:
        raise ValueError(
            f"Invalid year format '{year_number}', expected format: 'YYYY-YY' or 'YYYY_YY'"
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


# %% Aggregation
def _aggregate_by_level(df: pd.DataFrame, group_by_cols: list[str]) -> pd.DataFrame:
    """Helper function to aggregate data by specified grouping columns.

    This is an internal function used by aggregate_to_icb, aggregate_to_trust, etc.

    Args:
        df: DataFrame to aggregate
        group_by_cols: Columns to group by (e.g., ['ICB_Code', 'ICB_Name'])

    Returns:
        Aggregated DataFrame with recalculated percentages

    Raises:
        KeyError: If any group_by column is missing
    """
    # Check required columns exist
    missing_cols = [col for col in group_by_cols if col not in df.columns]
    if missing_cols:
        raise KeyError(f"DataFrame missing required columns: {missing_cols}")

    # Determine which columns to sum (only those that exist in df)
    cols_to_sum = []
    for col_group in ["likert_responses", "totals", "collection_modes"]:
        cols_to_sum.extend(
            [col for col in AGGREGATION_COLUMNS[col_group] if col in df.columns]
        )

    # Group and sum
    agg_df = df.groupby(group_by_cols, as_index=False)[cols_to_sum].sum()

    # Recalculate percentage positive (Very Good + Good) / Total Responses
    if all(col in agg_df.columns for col in ["1 Very Good", "2 Good", "Total Responses"]):
        agg_df["Percentage_Positive"] = (
            (agg_df["1 Very Good"] + agg_df["2 Good"]) / agg_df["Total Responses"]
        ).round(4)

    # Recalculate percentage negative (Poor + Very Poor) / Total Responses
    if all(col in agg_df.columns for col in ["4 Poor", "5 Very poor", "Total Responses"]):
        agg_df["Percentage_Negative"] = (
            (agg_df["4 Poor"] + agg_df["5 Very poor"]) / agg_df["Total Responses"]
        ).round(4)

    return agg_df


def aggregate_to_icb(df: pd.DataFrame) -> pd.DataFrame:
    """Aggregate organisation/trust level data to ICB level.

    Groups by ICB Code and Name, sums all response counts, and recalculates percentages.

    Args:
        df: DataFrame with trust-level data (must have ICB_Code, ICB_Name)

    Returns:
        DataFrame aggregated to ICB level with recalculated percentages

    Raises:
        KeyError: If required columns are missing

    >>> import pandas as pd
    >>> import numpy as np
    >>> from src.fft.processors import aggregate_to_icb
    >>> from src.fft.config import AGGREGATION_COLUMNS
    >>> df = pd.DataFrame({
    ...     'ICB_Code': ['ABC', 'ABC', 'DEF'],
    ...     'ICB_Name': ['ICB North', 'ICB North', 'ICB South'],
    ...     '1 Very Good': [10, 5, 8],
    ...     '2 Good': [3, 2, 4],
    ...     '3 Neither good nor poor': [1, 0, 1],
    ...     '4 Poor': [0, 1, 0],
    ...     '5 Very poor': [0, 0, 1],
    ...     '6 Dont Know': [1, 1, 0],
    ...     'Total Responses': [15, 9, 14],
    ...     'Total Eligible': [100, 50, 80]
    ... })
    >>> result = aggregate_to_icb(df)
    >>> result[result['ICB_Code'] == 'ABC']['Total Responses'].values[0]
    np.int64(24)
    >>> result[result['ICB_Code'] == 'ABC']['1 Very Good'].values[0]
    np.int64(15)
    >>> result[result['ICB_Code'] == 'DEF']['Total Responses'].values[0]
    np.int64(14)

    # Edge case: Missing required columns
    >>> df_missing = pd.DataFrame({
    ...     'ICB_Name': ['ICB North'],
    ...     '1 Very Good': [10],
    ...     'Total Responses': [15]
    ... })
    >>> aggregate_to_icb(df_missing)
    Traceback (most recent call last):
        ...
    KeyError: "DataFrame missing required columns: ['ICB_Code']"

    # Edge case: Empty DataFrame
    >>> df_empty = pd.DataFrame(columns=['ICB_Code', 'ICB_Name', '1 Very Good', 'Total Responses'])
    >>> result_empty = aggregate_to_icb(df_empty)
    >>> len(result_empty)
    0
    """
    return _aggregate_by_level(df, ["ICB_Code", "ICB_Name"])


def aggregate_to_trust(df: pd.DataFrame) -> pd.DataFrame:
    """Aggregate site level data to Trust level.

    Groups by Trust Code and Name, sums all response counts, and recalculates percentages.

    Args:
        df: DataFrame with site-level data (must have Trust_Code, Trust_Name)

    Returns:
        DataFrame aggregated to Trust level with recalculated percentages

    Raises:
        KeyError: If required columns are missing

    >>> import pandas as pd
    >>> import numpy as np
    >>> from src.fft.processors import aggregate_to_trust
    >>> from src.fft.config import AGGREGATION_COLUMNS
    >>> df = pd.DataFrame({
    ...     'Trust_Code': ['T01', 'T01', 'T02'],
    ...     'Trust_Name': ['Trust A', 'Trust A', 'Trust B'],
    ...     '1 Very Good': [10, 5, 8],
    ...     '2 Good': [3, 2, 4],
    ...     '3 Neither good nor poor': [1, 0, 1],
    ...     '4 Poor': [0, 1, 0],
    ...     '5 Very poor': [0, 0, 1],
    ...     '6 Dont Know': [1, 1, 0],
    ...     'Total Responses': [15, 9, 14],
    ...     'Total Eligible': [100, 50, 80]
    ... })
    >>> result = aggregate_to_trust(df)
    >>> result[result['Trust_Code'] == 'T01']['Total Responses'].values[0]
    np.int64(24)
    >>> result[result['Trust_Code'] == 'T01']['1 Very Good'].values[0]
    np.int64(15)
    >>> result[result['Trust_Code'] == 'T02']['Total Responses'].values[0]
    np.int64(14)

    # Edge case: Missing required columns
    >>> df_missing = pd.DataFrame({
    ...     'Trust_Name': ['Trust A'],
    ...     '1 Very Good': [10],
    ...     'Total Responses': [15]
    ... })
    >>> aggregate_to_trust(df_missing)
    Traceback (most recent call last):
        ...
    KeyError: "DataFrame missing required columns: ['Trust_Code']"

    # Edge case: Empty DataFrame
    >>> df_empty = pd.DataFrame(columns=['Trust_Code', 'Trust_Name', '1 Very Good', 'Total Responses'])
    >>> result_empty = aggregate_to_trust(df_empty)
    >>> len(result_empty)
    0
    """
    return _aggregate_by_level(df, ["Trust_Code", "Trust_Name"])


def aggregate_to_site(df: pd.DataFrame) -> pd.DataFrame:
    """Aggregate ward level data to Site level.

    Groups by Site Code and Name, sums all response counts, and recalculates percentages.

    Args:
        df: DataFrame with ward-level data (must have Site_Code, Site_Name)

    Returns:
        DataFrame aggregated to Site level with recalculated percentages

    Raises:
        KeyError: If required columns are missing

    >>> import pandas as pd
    >>> import numpy as np
    >>> from src.fft.processors import aggregate_to_site
    >>> df = pd.DataFrame({
    ...     'Site_Code': ['S01', 'S01', 'S02'],
    ...     'Site_Name': ['Site A', 'Site A', 'Site B'],
    ...     '1 Very Good': [10, 5, 8],
    ...     '2 Good': [3, 2, 4],
    ...     '3 Neither good nor poor': [1, 0, 1],
    ...     '4 Poor': [0, 1, 0],
    ...     '5 Very poor': [0, 0, 1],
    ...     '6 Dont Know': [1, 1, 0],
    ...     'Total Responses': [15, 9, 14],
    ...     'Total Eligible': [100, 50, 80]
    ... })
    >>> result = aggregate_to_site(df)
    >>> result[result['Site_Code'] == 'S01']['Total Responses'].values[0]
    np.int64(24)
    >>> result[result['Site_Code'] == 'S01']['1 Very Good'].values[0]
    np.int64(15)
    >>> result[result['Site_Code'] == 'S02']['Total Responses'].values[0]
    np.int64(14)

    # Edge case: Missing required columns
    >>> df_missing = pd.DataFrame({
    ...     'Site_Name': ['Site A'],
    ...     '1 Very Good': [10],
    ...     'Total Responses': [15]
    ... })
    >>> aggregate_to_site(df_missing)
    Traceback (most recent call last):
        ...
    KeyError: "DataFrame missing required columns: ['Site_Code']"

    # Edge case: Empty DataFrame
    >>> df_empty = pd.DataFrame(columns=['Site_Code', 'Site_Name', '1 Very Good', 'Total Responses'])
    >>> result_empty = aggregate_to_site(df_empty)
    >>> len(result_empty)
    0
    """
    return _aggregate_by_level(df, ["Site_Code", "Site_Name"])


def aggregate_to_national(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """Aggregate data to national level (NHS vs Independent providers).

    Creates three aggregations: Total (all), NHS only, and Independent (IS1) only.
    Also counts number of submitting organisations in each category.

    Args:
        df: DataFrame with ICB_Code column

    Returns:
        Tuple of (aggregated_df, org_counts) where:
        - aggregated_df has rows for 'Total', 'NHS', 'IS1'
        - org_counts is dict with keys 'nhs_count', 'is1_count', 'total_count'

    Raises:
        KeyError: If ICB_Code column is missing

    >>> import pandas as pd
    >>> import numpy as np
    >>> from src.fft.processors import aggregate_to_national
    >>> from src.fft.config import AGGREGATION_COLUMNS
    >>> df = pd.DataFrame({
    ...     'ICB_Code': ['ABC', 'DEF', 'IS1', 'IS1'],
    ...     'Trust_Code': ['T01', 'T02', 'T03', 'T04'],
    ...     '1 Very Good': [10, 5, 8, 3],
    ...     '2 Good': [3, 2, 4, 1],
    ...     '3 Neither good nor poor': [1, 0, 1, 0],
    ...     '4 Poor': [0, 1, 0, 1],
    ...     '5 Very poor': [0, 0, 1, 0],
    ...     '6 Dont Know': [1, 1, 0, 0],
    ...     'Total Responses': [15, 9, 14, 5],
    ...     'Total Eligible': [100, 50, 80, 30]
    ... })
    >>> result_df, counts = aggregate_to_national(df)
    >>> counts['nhs_count']
    2
    >>> counts['is1_count']
    2
    >>> counts['total_count']
    4
    >>> result_df[result_df['Submitter_Type'] == 'Total']['Total Responses'].values[0]
    np.int64(43)
    >>> result_df[result_df['Submitter_Type'] == 'NHS']['Total Responses'].values[0]
    np.int64(24)
    >>> result_df[result_df['Submitter_Type'] == 'IS1']['Total Responses'].values[0]
    np.int64(19)

    # Edge case: Missing ICB_Code
    >>> df_missing = pd.DataFrame({'Trust_Code': ['T01']})
    >>> aggregate_to_national(df_missing)
    Traceback (most recent call last):
        ...
    KeyError: "DataFrame must contain 'ICB_Code' column"

    # Edge case: Only NHS providers
    >>> df_nhs_only = pd.DataFrame({
    ...     'ICB_Code': ['ABC', 'DEF'],
    ...     '1 Very Good': [10, 5],
    ...     '2 Good': [3, 2],
    ...     '3 Neither good nor poor': [1, 0],
    ...     '4 Poor': [0, 1],
    ...     '5 Very poor': [0, 0],
    ...     '6 Dont Know': [1, 1],
    ...     'Total Responses': [15, 9],
    ...     'Total Eligible': [100, 50]
    ... })
    >>> result_df_nhs, counts_nhs = aggregate_to_national(df_nhs_only)
    >>> counts_nhs['is1_count']
    0
    >>> len(result_df_nhs)
    2

    # Edge case: Only IS1 providers
    >>> df_is1_only = pd.DataFrame({
    ...     'ICB_Code': ['IS1', 'IS1'],
    ...     '1 Very Good': [8, 3],
    ...     '2 Good': [4, 1],
    ...     '3 Neither good nor poor': [1, 0],
    ...     '4 Poor': [0, 1],
    ...     '5 Very poor': [1, 0],
    ...     '6 Dont Know': [0, 0],
    ...     'Total Responses': [14, 5],
    ...     'Total Eligible': [80, 30]
    ... })
    >>> result_df_is1, counts_is1 = aggregate_to_national(df_is1_only)
    >>> counts_is1['nhs_count']
    0
    >>> len(result_df_is1)
    2

    # Edge case: Empty DataFrame
    >>> df_empty = pd.DataFrame(columns=['ICB_Code', '1 Very Good', 'Total Responses'])
    >>> result_df_empty, counts_empty = aggregate_to_national(df_empty)
    >>> counts_empty['nhs_count']
    0
    >>> counts_empty['is1_count']
    0
    >>> counts_empty['total_count']
    0
    >>> len(result_df_empty)
    0
    """
    if "ICB_Code" not in df.columns:
        raise KeyError("DataFrame must contain 'ICB_Code' column")

    # Count organizations before transformation
    nhs_count = df[df["ICB_Code"] != "IS1"].shape[0]
    is1_count = df[df["ICB_Code"] == "IS1"].shape[0]
    total_count = df.shape[0]

    org_counts = {
        "nhs_count": nhs_count,
        "is1_count": is1_count,
        "total_count": total_count,
    }

    # Create working copy and add Submitter_Type
    work_df = df.copy()
    work_df["Submitter_Type"] = work_df["ICB_Code"].apply(
        lambda x: "IS1" if x == "IS1" else "NHS"
    )

    # Determine which columns to sum
    cols_to_sum = []
    for col_group in ["likert_responses", "totals", "collection_modes"]:
        cols_to_sum.extend(
            [col for col in AGGREGATION_COLUMNS[col_group] if col in work_df.columns]
        )

    # Aggregate by Submitter_Type
    agg_df = work_df.groupby("Submitter_Type", as_index=False)[cols_to_sum].sum()

    # Create Total row (sum of NHS + IS1)
    if len(agg_df) > 0:
        total_row = agg_df[cols_to_sum].sum().to_frame().T
        total_row["Submitter_Type"] = "Total"
        agg_df = pd.concat([total_row, agg_df], ignore_index=True)

    # Recalculate percentages
    if all(col in agg_df.columns for col in ["1 Very Good", "2 Good", "Total Responses"]):
        agg_df["Percentage_Positive"] = (
            (agg_df["1 Very Good"] + agg_df["2 Good"]) / agg_df["Total Responses"]
        ).round(4)

    if all(col in agg_df.columns for col in ["4 Poor", "5 Very poor", "Total Responses"]):
        agg_df["Percentage_Negative"] = (
            (agg_df["4 Poor"] + agg_df["5 Very poor"]) / agg_df["Total Responses"]
        ).round(4)

    return agg_df, org_counts

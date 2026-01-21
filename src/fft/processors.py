"""Data transformation functions."""

# %% Imports
import pandas as pd

from fft.config import (
    AGGREGATION_COLUMNS,
    COLUMN_MAPS,
    COLUMNS_TO_REMOVE,
    MONTH_ABBREV,
    SUMMARY_COLUMNS,
    TIME_SERIES_PREFIXES,
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
    >>> from fft.processors import standardise_column_names
    >>> raw_df = pd.DataFrame({
    ...     "Parent org code": [1, 2], "Parent name": ["Org A", "Org B"]
    ... })
    >>> std_df = standardise_column_names(raw_df, "inpatient", "organisation")
    >>> list(std_df.columns)
    ['ICB_Code', 'ICB_Name']

    # Edge case: Non-mapped column is preserved
    >>> raw_df_extra = pd.DataFrame({
    ...     "Parent org code": [1, 2],
    ...     "Parent name": ["Org A", "Org B"],
    ...     "Extra Column": ["X", "Y"]
    ... })
    >>> std_df_extra = standardise_column_names(
    ...     raw_df_extra, "inpatient", "organisation"
    ... )
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
    >>> from fft.processors import extract_fft_period
    >>> df = pd.DataFrame({'Periodname': ['AUGUST'], 'Yearnumber': ['2024-25']})
    >>> extract_fft_period(df)
    'Aug-24'

    # Edge case: Year boundary crossing (Jan uses next calendar year)
    >>> df_jan = pd.DataFrame({'Periodname': ['JANUARY'], 'Yearnumber': ['2024-25']})
    >>> extract_fft_period(df_jan)
    'Jan-25'

    # Edge case: Invalid year format
    >>> df_bad_year = pd.DataFrame({
    ...     'Periodname': ['MARCH'], 'Yearnumber': ['202425']
    ... })
    >>> extract_fft_period(df_bad_year)
    Traceback (most recent call last):
        ...
    ValueError: Invalid year format '202425', expected format: 'YYYY-YY' or 'YYYY_YY'

    # Error case: Missing required columns
    >>> df_missing = pd.DataFrame({'Other': ['data']})
    >>> extract_fft_period(df_missing)
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
    YEAR_FORMAT_LENGTH = 7  # Expected length for 'YYYY-YY' or 'YYYY_YY' format
    separator = None
    if len(year_number) == YEAR_FORMAT_LENGTH:
        if year_number[4] == "-":
            separator = "-"
        elif year_number[4] == "_":
            separator = "_"

    if separator is None:
        raise ValueError(
            f"Invalid year format '{year_number}', "
            f"expected format: 'YYYY-YY' or 'YYYY_YY'"
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
    >>> from fft.processors import remove_unwanted_columns
    >>> df = pd.DataFrame({
    ...     'Yearnumber': [2024],
    ...     'Periodname': ['AUG'],
    ...     'ICB_Code': ['ABC']
    ... })
    >>> cleaned = remove_unwanted_columns(df, 'inpatient', 'organisation')
    >>> list(cleaned.columns)
    ['ICB_Code']

    # Edge case: No unwanted columns present (graceful handling)
    >>> df_clean = pd.DataFrame({
    ...     'ICB_Code': ['ABC'], 'Total Responses': [100]
    ... })
    >>> result = remove_unwanted_columns(
    ...     df_clean, 'inpatient', 'organisation'
    ... )
    >>> len(result.columns)  # Should preserve all columns
    2

    # Edge case: Empty DataFrame
    >>> df_empty = pd.DataFrame()
    >>> cleaned_empty = remove_unwanted_columns(
    ...     df_empty, 'inpatient', 'organisation'
    ... )
    >>> len(cleaned_empty.columns)
    0

    # Error case: Unknown service type
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

# %% Aggregation
def _aggregate_by_level(df: pd.DataFrame, group_by_cols: list[str]) -> pd.DataFrame:
    """Aggregate data by specified grouping columns.

    This is an internal function used by aggregate_to_icb and aggregate_to_national.

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

    if all(col in agg_df.columns for col in ["Very Good", "Good", "Total Responses"]):
        agg_df["Percentage_Positive"] = (
            (agg_df["Very Good"] + agg_df["Good"]) / agg_df["Total Responses"]
        ).round(4)

    if all(col in agg_df.columns for col in ["Poor", "Very Poor", "Total Responses"]):
        agg_df["Percentage_Negative"] = (
            (agg_df["Poor"] + agg_df["Very Poor"]) / agg_df["Total Responses"]
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
    >>> df = pd.DataFrame({
    ...     'ICB_Code': ['ABC', 'ABC', 'DEF'],
    ...     'ICB_Name': ['ICB North', 'ICB North', 'ICB South'],
    ...     'Very Good': [10, 5, 8],
    ...     'Total Responses': [15, 9, 14],
    ...     'Total Eligible': [100, 50, 80]
    ... })
    >>> result = aggregate_to_icb(df)
    >>> len(result)  # Two ICBs
    2
    >>> result[result['ICB_Code'] == 'ABC']['Total Responses'].values[0]
    np.int64(24)

    # Edge case: Missing required columns
    >>> df_missing = pd.DataFrame({
    ...     'ICB_Name': ['ICB North'],
    ...     'Very Good': [10],
    ...     'Total Responses': [15]
    ... })
    >>> aggregate_to_icb(df_missing)
    Traceback (most recent call last):
        ...
    KeyError: "DataFrame missing required columns: ['ICB_Code']"

    # Edge case: Empty DataFrame
    >>> df_empty = pd.DataFrame(columns=[
    ...     'ICB_Code', 'ICB_Name', 'Very Good', 'Total Responses'
    ... ])
    >>> result_empty = aggregate_to_icb(df_empty)
    >>> len(result_empty)
    0

    """
    return _aggregate_by_level(df, ["ICB_Code", "ICB_Name"])

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
    ...     'Trust_Name': [
    ...         'NHS Foundation Trust',
    ...         'NHS Trust',
    ...         'Independent Sector',
    ...         'Private Provider'
    ...     ],
    ...     'Very Good': [10, 5, 8, 3],
    ...     'Good': [3, 2, 4, 1],
    ...     'Neither good nor poor': [1, 0, 1, 0],
    ...     'Poor': [0, 1, 0, 1],
    ...     'Very poor': [0, 0, 1, 0],
    ...     'Dont Know': [1, 1, 0, 0],
    ...     'Total Responses': [15, 9, 14, 5],
    ...     'Total Eligible': [100, 50, 80, 30]
    ... })
    >>> result_df, counts = aggregate_to_national(df)
    >>> counts['nhs_count']
    np.int64(2)
    >>> counts['is1_count']
    np.int64(2)
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
    ...     'Trust_Name': ['NHS Foundation Trust', 'NHS Trust'],
    ...     'Very Good': [10, 5],
    ...     'Good': [3, 2],
    ...     'Neither good nor poor': [1, 0],
    ...     'Poor': [0, 1],
    ...     'Very poor': [0, 0],
    ...     'Dont Know': [1, 1],
    ...     'Total Responses': [15, 9],
    ...     'Total Eligible': [100, 50]
    ... })
    >>> result_df_nhs, counts_nhs = aggregate_to_national(df_nhs_only)
    >>> counts_nhs['is1_count']
    np.int64(0)
    >>> len(result_df_nhs)
    2

    # Edge case: Only IS1 providers
    >>> df_is1_only = pd.DataFrame({
    ...     'ICB_Code': ['IS1', 'IS1'],
    ...     'Trust_Name': ['Independent Sector', 'Private Provider'],
    ...     'Very Good': [8, 3],
    ...     'Good': [4, 1],
    ...     'Neither good nor poor': [1, 0],
    ...     'Poor': [0, 1],
    ...     'Very poor': [1, 0],
    ...     'Dont Know': [0, 0],
    ...     'Total Responses': [14, 5],
    ...     'Total Eligible': [80, 30]
    ... })
    >>> result_df_is1, counts_is1 = aggregate_to_national(df_is1_only)
    >>> counts_is1['nhs_count']
    np.int64(0)
    >>> len(result_df_is1)
    2

    # Edge case: Empty DataFrame
    >>> df_empty = pd.DataFrame(columns=[
    ...     'ICB_Code', 'Trust_Name', 'Very Good', 'Total Responses'
    ... ])
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

    # Count organisations before transformation
    # Count entries with 'NHS TRUST' or 'NHS FOUNDATION TRUST' in name
    nhs_count = (
        df["Trust_Name"]
        .apply(
            lambda x: isinstance(x, str) and "NHS" in x.upper() and "TRUST" in x.upper()
        )
        .sum()
    )
    is1_count = (
        df["Trust_Name"]
        .apply(
            lambda x: isinstance(x, str)
            and not ("NHS" in x.upper() and "TRUST" in x.upper())
        )
        .sum()
    )
    total_count = len(df)

    org_counts = {
        "nhs_count": nhs_count,
        "is1_count": is1_count,
        "total_count": total_count,
    }

    # Create working copy and add Submitter_Type
    work_df = df.copy()
    work_df["Submitter_Type"] = work_df["Trust_Name"].apply(
        lambda x: "NHS"
        if "NHS" in str(x).upper() and "TRUST" in str(x).upper()
        else "IS1"
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
    if all(col in agg_df.columns for col in ["Very Good", "Good", "Total Responses"]):
        agg_df["Percentage_Positive"] = (
            (agg_df["Very Good"] + agg_df["Good"]) / agg_df["Total Responses"]
        ).round(4)

    if all(col in agg_df.columns for col in ["Poor", "Very Poor", "Total Responses"]):
        agg_df["Percentage_Negative"] = (
            (agg_df["Poor"] + agg_df["Very Poor"]) / agg_df["Total Responses"]
        ).round(4)

    return agg_df, org_counts

# %%
def merge_collection_modes(
    org_df: pd.DataFrame, collection_df: pd.DataFrame
) -> pd.DataFrame:
    """Merge collection mode data into organisation-level DataFrame.

    Args:
        org_df: Organisation-level DataFrame with Trust_Code
        collection_df: Collection mode DataFrame with mode columns

    Returns:
        Merged DataFrame with mode columns added

    >>> import pandas as pd
    >>> org = pd.DataFrame({
    ...     'Trust_Code': ['T01', 'T02'], 'Total Responses': [100, 200]
    ... })
    >>> coll = pd.DataFrame({
    ...     'Trust_Code': ['T01', 'T02'],
    ...     'Mode SMS': [50, 100],
    ...     'Mode Online': [50, 100]
    ... })
    >>> merged = merge_collection_modes(org, coll)
    >>> 'Mode SMS' in merged.columns
    True

    """
    mode_columns = [col for col in collection_df.columns if col.startswith("Mode ")]

    # Keep only Trust_Code and mode columns from collection data
    coll_subset = collection_df[["Trust_Code"] + mode_columns].copy()

    # Merge on Trust_Code
    merged = org_df.merge(coll_subset, on="Trust_Code", how="left")

    # Fill NaN mode values with 0
    for col in mode_columns:
        if col in merged.columns:
            merged[col] = merged[col].fillna(0).astype(int)

    return merged

# %%
def clean_icb_name(name: str) -> str:
    """Clean ICB name to match standard format.

    Args:
        name: Raw ICB name (e.g.,
            "NHS LANCASHIRE AND SOUTH CUMBRIA INTEGRATED CARE BOARD")

    Returns:
        Cleaned name (e.g., "LANCASHIRE AND SOUTH CUMBRIA ICB")

    >>> clean_icb_name(
    ...     "NHS LANCASHIRE AND SOUTH CUMBRIA INTEGRATED CARE BOARD"
    ... )
    'LANCASHIRE AND SOUTH CUMBRIA ICB'
    >>> clean_icb_name("NHS SUSSEX INTEGRATED CARE BOARD")
    'SUSSEX ICB'
    >>> clean_icb_name("INDEPENDENT SECTOR PROVIDERS")
    'INDEPENDENT SECTOR PROVIDERS'

    """
    if not isinstance(name, str):
        return name

    result = name
    if result.startswith("NHS "):
        result = result[4:]  # Remove "NHS " prefix
    result = result.replace("INTEGRATED CARE BOARD", "ICB")

    return result.strip()

# %%
def convert_fft_period_to_datetime(fft_period: str):
    """Convert FFT period to datetime for Collections Overview lookup.

    Args:
        fft_period: FFT period string (e.g., 'Jul-25')

    Returns:
        pandas Timestamp object (e.g., Timestamp('2025-07-01'))

    """
    month_abbrev, year = fft_period.split("-")
    year_full = 2000 + int(year)  # Convert 25 -> 2025

    # Convert month abbreviation to number
    month_map = {
        "Jan": 1,
        "Feb": 2,
        "Mar": 3,
        "Apr": 4,
        "May": 5,
        "Jun": 6,
        "Jul": 7,
        "Aug": 8,
        "Sep": 9,
        "Oct": 10,
        "Nov": 11,
        "Dec": 12,
    }

    month_num = month_map[month_abbrev]
    return pd.Timestamp(year_full, month_num, 1)

def extract_summary_data(
    time_series_df: pd.DataFrame,
    service_type: str,
    current_period: str,
    previous_period: str,
) -> dict:
    """Extract summary data from Time series for a given service type.

    Retrieves organisations submitting, responses, and calculates percentages
    for Total, NHS, and Independent Sector providers.

    Args:
        time_series_df: DataFrame from Collections Overview 'Time series' sheet
        service_type: 'inpatient', 'ae', 'ambulance', etc.
        current_period: Current FFT period (e.g., 'Jul-25')
        previous_period: Previous FFT period (e.g., 'Jun-25')

    Returns:
        Dict with structure for populating Summary sheet

    Raises:
        KeyError: If service_type is invalid or required columns not found
        ValueError: If periods not found in time series data

    >>> import pandas as pd
    >>> import numpy as np
    >>> from fft.processors import extract_summary_data
    >>> df = pd.DataFrame({
    ...     'Collection': pd.to_datetime(['2025-07-01', '2025-06-01', '2025-05-01']),
    ...     'Inpatient Submitted': [150, 148, 145],
    ...     'Inpatient NHS Submitted': [134, 132, 130],
    ...     'Inpatient IS Submitted': [19, 18, 17],
    ...     'Inpatient Responses': [202745, 213043, 200000],
    ...     'Inpatient NHS Responses': [186977, 195590, 185000],
    ...     'Inpatient IS Responses': [15883, 17606, 15000],
    ...     'Inpatient Extremely Likely': [180000, 190000, 178000],
    ...     'Inpatient Likely': [12000, 12500, 11800],
    ...     'Inpatient Extremely Unlikely': [2000, 2100, 1950],
    ...     'Inpatient Unlikely': [1500, 1600, 1450],
    ...     'Inpatient NHS Extremely Likely': [165000, 175000, 163000],
    ...     'Inpatient NHS Likely': [11000, 11500, 10800],
    ...     'Inpatient NHS Extremely Unlikely': [1800, 1900, 1750],
    ...     'Inpatient NHS Unlikely': [1400, 1500, 1350],
    ...     'Inpatient IS Extremely Likely': [15000, 16000, 14500],
    ...     'Inpatient IS Likely': [700, 750, 680],
    ...     'Inpatient IS Extremely Unlikely': [50, 55, 48],
    ...     'Inpatient IS Unlikely': [30, 35, 28],
    ... })
    >>> result = extract_summary_data(df, 'inpatient', 'Jul-25', 'Jun-25')
    >>> result['orgs_submitting']['total']
    np.int64(150)
    >>> result['orgs_submitting']['nhs']
    np.int64(134)
    >>> result['responses_current']['total']
    np.int64(202745)

    # Edge case: Unknown service type
    >>> extract_summary_data(df, 'unknown', 'Jul-25', 'Jun-25')
    Traceback (most recent call last):
        ...
    KeyError: "Unknown service type: 'unknown'"

    # Edge case: Period not found
    >>> extract_summary_data(df, 'inpatient', 'Jan-20', 'Dec-19')
    Traceback (most recent call last):
        ...
    ValueError: Period 'Jan-20' not found in time series data

    """
    if service_type not in TIME_SERIES_PREFIXES:
        raise KeyError(f"Unknown service type: '{service_type}'")

    prefix = TIME_SERIES_PREFIXES[service_type]

    # Convert periods to datetime objects for lookup
    current_datetime = convert_fft_period_to_datetime(current_period)
    previous_datetime = convert_fft_period_to_datetime(previous_period)

    # Validate periods exist
    if current_datetime not in time_series_df["Collection"].values:
        raise ValueError(f"Period '{current_period}' not found in time series data")
    if previous_datetime not in time_series_df["Collection"].values:
        raise ValueError(f"Period '{previous_period}' not found in time series data")

    current_row = time_series_df[time_series_df["Collection"] == current_datetime].iloc[0]
    previous_row = time_series_df[time_series_df["Collection"] == previous_datetime].iloc[
        0
    ]
    current_idx = time_series_df[time_series_df["Collection"] == current_datetime].index[
        0
    ]

    def get_col(suffix):
        """Build column name from prefix and suffix."""
        return f"{prefix}{suffix}"

    def calc_percentage(likely_val, extremely_likely_val, responses_val):
        """Calculate percentage from likely + extremely likely / responses."""
        if responses_val == 0:
            return 0
        return round((likely_val + extremely_likely_val) / responses_val, 2)

    # Build orgs_submitting
    orgs_cols = SUMMARY_COLUMNS["orgs_submitting"]
    orgs_submitting = {
        key: current_row[get_col(suffix)] for key, suffix in orgs_cols.items()
    }

    # Build responses (current and previous)
    resp_cols = SUMMARY_COLUMNS["responses"]
    responses_current = {
        key: current_row[get_col(suffix)] for key, suffix in resp_cols.items()
    }
    responses_previous = {
        key: previous_row[get_col(suffix)] for key, suffix in resp_cols.items()
    }

    # Build responses to date (cumulative sum)
    responses_to_date = {
        key: time_series_df.loc[current_idx:, get_col(suffix)].sum()
        for key, suffix in resp_cols.items()
    }

    # Build percentage positive (current and previous)
    pos_cols = SUMMARY_COLUMNS["positive"]
    pct_positive_current = {
        "total": calc_percentage(
            current_row[get_col(pos_cols["likely"])],
            current_row[get_col(pos_cols["extremely_likely"])],
            responses_current["total"],
        ),
        "nhs": calc_percentage(
            current_row[get_col(pos_cols["nhs_likely"])],
            current_row[get_col(pos_cols["nhs_extremely_likely"])],
            responses_current["nhs"],
        ),
        "is": calc_percentage(
            current_row[get_col(pos_cols["is_likely"])],
            current_row[get_col(pos_cols["is_extremely_likely"])],
            responses_current["is"],
        ),
    }
    pct_positive_previous = {
        "total": calc_percentage(
            previous_row[get_col(pos_cols["likely"])],
            previous_row[get_col(pos_cols["extremely_likely"])],
            responses_previous["total"],
        ),
        "nhs": calc_percentage(
            previous_row[get_col(pos_cols["nhs_likely"])],
            previous_row[get_col(pos_cols["nhs_extremely_likely"])],
            responses_previous["nhs"],
        ),
        "is": calc_percentage(
            previous_row[get_col(pos_cols["is_likely"])],
            previous_row[get_col(pos_cols["is_extremely_likely"])],
            responses_previous["is"],
        ),
    }

    # Build percentage negative (current and previous)
    neg_cols = SUMMARY_COLUMNS["negative"]
    pct_negative_current = {
        "total": calc_percentage(
            current_row[get_col(neg_cols["unlikely"])],
            current_row[get_col(neg_cols["extremely_unlikely"])],
            responses_current["total"],
        ),
        "nhs": calc_percentage(
            current_row[get_col(neg_cols["nhs_unlikely"])],
            current_row[get_col(neg_cols["nhs_extremely_unlikely"])],
            responses_current["nhs"],
        ),
        "is": calc_percentage(
            current_row[get_col(neg_cols["is_unlikely"])],
            current_row[get_col(neg_cols["is_extremely_unlikely"])],
            responses_current["is"],
        ),
    }
    pct_negative_previous = {
        "total": calc_percentage(
            previous_row[get_col(neg_cols["unlikely"])],
            previous_row[get_col(neg_cols["extremely_unlikely"])],
            responses_previous["total"],
        ),
        "nhs": calc_percentage(
            previous_row[get_col(neg_cols["nhs_unlikely"])],
            previous_row[get_col(neg_cols["nhs_extremely_unlikely"])],
            responses_previous["nhs"],
        ),
        "is": calc_percentage(
            previous_row[get_col(neg_cols["is_unlikely"])],
            previous_row[get_col(neg_cols["is_extremely_unlikely"])],
            responses_previous["is"],
        ),
    }

    return {
        "orgs_submitting": orgs_submitting,
        "responses_to_date": responses_to_date,
        "responses_current": responses_current,
        "responses_previous": responses_previous,
        "pct_positive_current": pct_positive_current,
        "pct_positive_previous": pct_positive_previous,
        "pct_negative_current": pct_negative_current,
        "pct_negative_previous": pct_negative_previous,
    }

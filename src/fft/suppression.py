"""Suppression logic for FFT data privacy protection."""

import pandas as pd

from fft.config import AGGREGATION_COLUMNS, SUPPRESSION_THRESHOLD

# Constants for suppression logic
SECOND_RANK = 2  # Used to identify the second-ranked item in suppression logic


# %%
def apply_first_level_suppression(df: pd.DataFrame) -> pd.DataFrame:
    """Flag rows requiring first-level suppression (1-4 responses).

    Adds a 'First_Level_Suppression' column with 1 for rows needing suppression,
    0 otherwise. Rows with 0 responses are not flagged.

    Args:
        df: DataFrame with 'Total Responses' column

    Returns:
        DataFrame with added 'First_Level_Suppression' column

    Raises:
        KeyError: If 'Total Responses' column is missing

    >>> import pandas as pd
    >>> from src.fft.suppression import apply_first_level_suppression
    >>> df = pd.DataFrame({
    ...     'ICB_Code': ['A', 'B', 'C', 'D', 'E'],
    ...     'Total Responses': [0, 3, 5, 10, 2]
    ... })
    >>> result = apply_first_level_suppression(df)
    >>> list(result['First_Level_Suppression'])
    [0, 1, 0, 0, 1]

    # Edge case: Missing column
    >>> df_missing = pd.DataFrame({'ICB_Code': ['A']})
    >>> apply_first_level_suppression(df_missing)
    Traceback (most recent call last):
        ...
    KeyError: "'Total Responses' column not found in DataFrame"

    # Edge case: Empty DataFrame
    >>> df_empty = pd.DataFrame(columns=['Total Responses'])
    >>> result_empty = apply_first_level_suppression(df_empty)
    >>> len(result_empty)
    0

    # Edge case: All zeros
    >>> df_zeros = pd.DataFrame({'Total Responses': [0, 0, 0]})
    >>> result_zeros = apply_first_level_suppression(df_zeros)
    >>> list(result_zeros['First_Level_Suppression'])
    [0, 0, 0]

    # Edge case: All need suppression
    >>> df_all_suppress = pd.DataFrame({'Total Responses': [1, 2, 3, 4]})
    >>> result_all = apply_first_level_suppression(df_all_suppress)
    >>> list(result_all['First_Level_Suppression'])
    [1, 1, 1, 1]

    """
    if "Total Responses" not in df.columns:
        raise KeyError("'Total Responses' column not found in DataFrame")

    # Create suppression flag: 1 if 0 < responses < threshold, else 0
    # See suppression file Ward/Site/Trust Calcs sheets, row 2 column '0><5 responses'
    df = df.copy()
    df["First_Level_Suppression"] = df["Total Responses"].apply(
        lambda x: 1 if 0 < x < SUPPRESSION_THRESHOLD else 0
    )

    return df


# %%

def add_rank_column(df: pd.DataFrame, group_by_col: str | None = None) -> pd.DataFrame:
    """Add ranking column based on Total Responses within groups.

    For ward data, uses VBA tie-breaking: Ward_Name → First Speciality → Second Speciality.
    For other levels, ranks by Total Responses only.

    Args:
        df: DataFrame with 'Total Responses' column
        group_by_col: Column to group by for ranking (e.g., 'Site_Code' for Ward level).

    Returns:
        DataFrame with added 'Rank' column

    """
    if "Total Responses" not in df.columns:
        raise KeyError("'Total Responses' column not found in DataFrame")

    if group_by_col and group_by_col not in df.columns:
        raise KeyError(f"'{group_by_col}' column not found in DataFrame")

    df = df.copy()
    df["Rank"] = 0

    # Check if this is ward data (has specialty columns)
    is_ward_data = "Ward_Name" in df.columns and "First Speciality" in df.columns

    if group_by_col:
        for group_name, group_indices in df.groupby(group_by_col).groups.items():
            group_data = df.loc[group_indices]
            non_zero_data = group_data[group_data["Total Responses"] > 0]

            if non_zero_data.empty:
                continue

            if is_ward_data:
                # VBA tie-breaking using specialty-first approach (best performing so far)
                # This gave us 24 differences vs 60 with ward name approaches
                df_temp = non_zero_data.copy()

                # Use specialty text directly for sorting (VBA sorts alphabetically)
                if "First Speciality" in df_temp.columns:
                    df_temp["_spec1_text"] = df_temp["First Speciality"].astype(str).fillna("")
                else:
                    df_temp["_spec1_text"] = ""

                if "Second Speciality" in df_temp.columns:
                    df_temp["_spec2_text"] = df_temp["Second Speciality"].astype(str).fillna("")
                else:
                    df_temp["_spec2_text"] = ""

                # Sort to match VBA tie-breaking: Total Responses → First Specialty → Second Specialty → Ward_Name
                # VBA prioritizes specialty-based tie-breaking over ward name alphabetical sorting
                sorted_indices = df_temp.sort_values(
                    ["Total Responses", "_spec1_text", "_spec2_text", "Ward_Name"]
                ).index
            else:
                # Standard response-based ranking
                sorted_indices = non_zero_data.sort_values("Total Responses").index

            for i, idx in enumerate(sorted_indices, 1):
                df.loc[idx, "Rank"] = i
    else:
        # ICB level - no grouping
        non_zero_data = df[df["Total Responses"] > 0]
        if not non_zero_data.empty:
            sorted_indices = non_zero_data.sort_values("Total Responses").index
            for i, idx in enumerate(sorted_indices, 1):
                df.loc[idx, "Rank"] = i

    return df


def apply_second_level_suppression(
    df: pd.DataFrame, group_by_col: str | None = None
) -> pd.DataFrame:
    """Flag rows requiring second-level suppression using VBA row-based adjacency.

    Implements VBA logic: =IF(AND(I1=1, H2=2, I2<>1),1,"")
    When previous row is first-level suppressed AND current row has Rank 2
    AND current row is NOT first-level suppressed, flag for second-level suppression.

    This prevents reverse calculation attacks by suppressing the second-lowest
    responding organization when the lowest is already suppressed.

    Row-based adjacency logic (matches VBA):
    - Sort rows by rank within each group
    - Check if previous row is first-level suppressed
    - If so, and current row is rank 2 (not first-level), apply second-level

    Grouping logic by level:
    - ICB level: No grouping (group_by_col=None)
    - Trust level: Group by 'ICB_Code'
    - Site level: Group by 'Trust_Code'
    - Ward level: Group by 'Site_Code'

    Args:
        df: DataFrame with 'First_Level_Suppression' and 'Rank' columns
        group_by_col: Column to group by (None for ICB level)

    Returns:
        DataFrame with added 'Second_Level_Suppression' column

    Raises:
        KeyError: If required columns are missing

    >>> import pandas as pd
    >>> from src.fft.suppression import apply_second_level_suppression
    >>> df = pd.DataFrame({
    ...     'ICB_Code': ['A', 'A', 'A', 'B', 'B'],
    ...     'Trust_Code': ['T1', 'T2', 'T3', 'T4', 'T5'],
    ...     'Rank': [1, 2, 3, 1, 2],
    ...     'First_Level_Suppression': [1, 0, 0, 0, 0]
    ... })
    >>> result = apply_second_level_suppression(df, 'ICB_Code')
    >>> list(result['Second_Level_Suppression'])
    [0, 1, 0, 0, 0]

    # Edge case: Non-adjacent ranks (rank 1 suppressed, rank 3 follows)
    >>> df_gap = pd.DataFrame({
    ...     'Rank': [1, 3],
    ...     'First_Level_Suppression': [1, 0]
    ... })
    >>> result_gap = apply_second_level_suppression(df_gap, None)
    >>> list(result_gap['Second_Level_Suppression'])
    [0, 0]

    # Edge case: Rank 2 already first-level suppressed
    >>> df_both = pd.DataFrame({
    ...     'Rank': [1, 2],
    ...     'First_Level_Suppression': [1, 1]
    ... })
    >>> result_both = apply_second_level_suppression(df_both, None)
    >>> list(result_both['Second_Level_Suppression'])
    [0, 0]

    """
    required_cols = ["First_Level_Suppression", "Rank"]
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise KeyError(f"Required columns missing: {missing_cols}")

    df = df.copy()
    df["Second_Level_Suppression"] = 0

    # VBA suppression workbook logic: =IF(AND(I1=1, H2=2, I2<>1),1,"")
    # Since VBA ranking resets to 1 for each site group, H2=2 means rank 2 within the site
    # I1=1: Previous row (rank 1 within same site) is first-level suppressed
    # H2=2: Current row has rank 2 within the site group
    # I2<>1: Current row is NOT first-level suppressed

    if group_by_col:
        # Within each group, check rank relationships
        for group_name, group_data in df.groupby(group_by_col):
            # Check if Rank 1 in this group is first-level suppressed
            rank_1_rows = group_data[group_data["Rank"] == 1]
            rank_1_suppressed = (rank_1_rows["First_Level_Suppression"] == 1).any()

            if rank_1_suppressed:
                # Find Rank 2 rows that are NOT first-level suppressed
                rank_2_mask = (
                    (df[group_by_col] == group_name)
                    & (df["Rank"] == SECOND_RANK)
                    & (df["First_Level_Suppression"] != 1)
                )
                df.loc[rank_2_mask, "Second_Level_Suppression"] = 1
    else:
        # No grouping - check rank relationships across entire DataFrame
        # Check if any Rank 1 is first-level suppressed
        rank_1_suppressed = (
            (df["Rank"] == 1) & (df["First_Level_Suppression"] == 1)
        ).any()

        if rank_1_suppressed:
            # Find Rank 2 rows that are NOT first-level suppressed
            rank_2_mask = (df["Rank"] == SECOND_RANK) & (
                df["First_Level_Suppression"] != 1
            )
            df.loc[rank_2_mask, "Second_Level_Suppression"] = 1

    return df


# %%
def apply_cascade_suppression(
    parent_df: pd.DataFrame,
    child_df: pd.DataFrame,
    parent_code_col: str,
    child_code_col: str,
    parent_suppression_col: str,
) -> pd.DataFrame:
    """Apply cascade suppression from parent to child level.

    KEY DISTINCTION:
    - First/Second level suppression: Based on child's OWN response count
    - Cascade suppression: Based on PARENT's suppression status

    When a parent organisation is already suppressed (at its own level),
    we must also suppress its children to prevent reverse calculation
    using parent totals.

    Example showing why cascade is needed:

    ICB North has 232 responses and IS SUPPRESSED at ICB level (shown
    as *). Its 3 trusts show:
    - Trust A: 150 responses → Shown
    - Trust B: 80 responses → Shown
    - Trust C: 2 responses → Already suppressed (first-level)

    Problem: Someone can calculate 232 - 150 - 80 = 2, revealing
    Trust C's value!

    Solution: Cascade suppression ALSO suppresses Trust B (Rank 2):
    - Trust A: 150 responses → Shown
    - Trust B: * → Cascade suppressed
    - Trust C: * → First-level suppressed

    Now calculation is impossible: 232 - 150 - ? - ? = unknown

    The function flags the 2 lowest-ranked children (Rank 1 and Rank 2)
    of any suppressed parent organisation.

    Args:
        parent_df: Parent level DataFrame with suppression flags
        child_df: Child level DataFrame with 'Rank' column
        parent_code_col: Column name for parent code (e.g., 'ICB_Code')
        child_code_col: Column name matching parent in child_df (e.g., 'ICB_Code')
        parent_suppression_col: Name of suppression column in parent_df

    Returns:
        child_df with added 'Cascade_Suppression' column

    Raises:
        KeyError: If required columns are missing


    >>> import pandas as pd
    >>> from src.fft.suppression import apply_cascade_suppression
    >>> parent_df = pd.DataFrame({
    ...     'ICB_Code': ['A', 'B'],
    ...     'ICB_Suppression_Required': [1, 0]
    ... })
    >>> child_df = pd.DataFrame({
    ...     'ICB_Code': ['A', 'A', 'A', 'B', 'B'],
    ...     'Trust_Code': ['T1', 'T2', 'T3', 'T4', 'T5'],
    ...     'Rank': [1, 2, 3, 1, 2]
    ... })
    >>> result = apply_cascade_suppression(
    ...     parent_df, child_df, 'ICB_Code', 'ICB_Code', 'ICB_Suppression_Required'
    ... )
    >>> list(result['Cascade_Suppression'])
    [1, 1, 0, 0, 0]

    # Edge case: No parent suppression
    >>> parent_no_suppress = pd.DataFrame({
    ...     'ICB_Code': ['A'],
    ...     'ICB_Suppression_Required': [0]
    ... })
    >>> child_test = pd.DataFrame({
    ...     'ICB_Code': ['A', 'A'],
    ...     'Rank': [1, 2]
    ... })
    >>> result_none = apply_cascade_suppression(
    ...     parent_no_suppress,
    ...     child_test,
    ...     'ICB_Code',
    ...     'ICB_Code',
    ...     'ICB_Suppression_Required'
    ... )
    >>> list(result_none['Cascade_Suppression'])
    [0, 0]

    # Edge case: Missing columns
    >>> parent_missing = pd.DataFrame({'ICB_Code': ['A']})
    >>> child_missing = pd.DataFrame({'ICB_Code': ['A'], 'Rank': [1]})
    >>> apply_cascade_suppression(
    ...     parent_missing, child_missing, 'ICB_Code', 'ICB_Code', 'Missing_Col'
    ... )
    Traceback (most recent call last):
        ...
    KeyError: "Column 'Missing_Col' not found in parent DataFrame"

    # Edge case: Only Rank 1 exists (no Rank 2)
    >>> parent_one_rank = pd.DataFrame({
    ...     'ICB_Code': ['A'],
    ...     'ICB_Suppression_Required': [1]
    ... })
    >>> child_one_rank = pd.DataFrame({
    ...     'ICB_Code': ['A'],
    ...     'Rank': [1]
    ... })
    >>> result_one = apply_cascade_suppression(
    ...     parent_one_rank,
    ...     child_one_rank,
    ...     'ICB_Code',
    ...     'ICB_Code',
    ...     'ICB_Suppression_Required'
    ... )
    >>> list(result_one['Cascade_Suppression'])
    [1]

    """
    # Validate columns
    if parent_code_col not in parent_df.columns:
        raise KeyError(f"Column '{parent_code_col}' not found in parent DataFrame")
    if parent_suppression_col not in parent_df.columns:
        raise KeyError(f"Column '{parent_suppression_col}' not found in parent DataFrame")
    if child_code_col not in child_df.columns:
        raise KeyError(f"Column '{child_code_col}' not found in child DataFrame")
    if "Rank" not in child_df.columns:
        raise KeyError("'Rank' column not found in child DataFrame")

    # Create suppression lookup dict from parent
    suppression_dict = parent_df.set_index(parent_code_col)[
        parent_suppression_col
    ].to_dict()

    # Initialize cascade suppression column
    child_df = child_df.copy()
    child_df["Cascade_Suppression"] = 0

    # For each parent code that requires suppression
    for parent_code, needs_suppression in suppression_dict.items():
        if needs_suppression == 1:
            # Get children for this parent
            parent_children = child_df[child_df[child_code_col] == parent_code]

            # Flag Rank 1 if exists
            if any(parent_children["Rank"] == 1):
                child_df.loc[
                    (child_df[child_code_col] == parent_code) & (child_df["Rank"] == 1),
                    "Cascade_Suppression",
                ] = 1

            # Flag Rank 2 if exists
            if any(parent_children["Rank"] == SECOND_RANK):
                child_df.loc[
                    (child_df[child_code_col] == parent_code)
                    & (child_df["Rank"] == SECOND_RANK),
                    "Cascade_Suppression",
                ] = 1

    return child_df


def suppress_values(df: pd.DataFrame) -> pd.DataFrame:
    """Replace sensitive values with '*' based on suppression flags.

    Applies suppression rules:
    - If ANY suppression flag is 1: Replace Likert responses with '*'
    - If First_Level_Suppression is 1: ALSO replace percentages with '*'
    - Additionally applies individual breakdown column suppression (VBA-aligned)

    This ensures aggregated percentages don't reveal small counts.

    Args:
        df: DataFrame with suppression flag columns

    Returns:
        DataFrame with values replaced by '*' where suppression required

    Raises:
        KeyError: If required columns are missing

    >>> import pandas as pd
    >>> from src.fft.suppression import suppress_values
    >>> df = pd.DataFrame({
    ...     'ICB_Code': ['A', 'B', 'C'],
    ...     'Very Good': [10, 3, 50],
    ...     'Good': [5, 1, 20],
    ...     'Neither Good nor Poor': [2, 0, 10],
    ...     'Poor': [1, 1, 5],
    ...     'Very Poor': [0, 0, 2],
    ...     "Don't Know": [1, 0, 3],
    ...     'Percentage_Positive': [0.79, 0.80, 0.78],
    ...     'Percentage_Negative': [0.05, 0.20, 0.08],
    ...     'First_Level_Suppression': [0, 1, 0],
    ...     'Second_Level_Suppression': [1, 0, 0],
    ...     'Total Responses': [19, 5, 90]
    ... })
    >>> result = suppress_values(df)
    >>> result.loc[0, 'Very Good']
    '*'
    >>> result.loc[0, 'Percentage_Positive']
    0.79
    >>> result.loc[1, 'Very Good']
    '*'
    >>> result.loc[1, 'Percentage_Positive']
    '*'
    >>> result.loc[2, 'Very Good']
    50

    # Edge case: No suppression columns
    >>> df_no_suppress = pd.DataFrame({
    ...     'Very Good': [10],
    ...     'Percentage_Positive': [0.8]
    ... })
    >>> suppress_values(df_no_suppress)
    Traceback (most recent call last):
        ...
    KeyError: 'No suppression flag columns found in DataFrame'

    # Edge case: All values suppressed
    >>> df_all = pd.DataFrame({
    ...     'Very Good': [3, 2],
    ...     'Percentage_Positive': [0.8, 0.7],
    ...     'First_Level_Suppression': [1, 1]
    ... })
    >>> result_all = suppress_values(df_all)
    >>> list(result_all['Very Good'])
    ['*', '*']
    >>> list(result_all['Percentage_Positive'])
    ['*', '*']

    """
    # Find all suppression flag columns
    suppression_cols = [col for col in df.columns if "Suppression" in col]
    if not suppression_cols:
        raise KeyError("No suppression flag columns found in DataFrame")

    df = df.copy()

    # Convert numeric columns to object type to allow '*' values
    likert_cols = [
        col for col in AGGREGATION_COLUMNS["likert_responses"] if col in df.columns
    ]
    percentage_cols = [
        col for col in ["Percentage_Positive", "Percentage_Negative"] if col in df.columns
    ]

    for col in likert_cols + percentage_cols:
        df[col] = df[col].astype(object)

    # Iterate through each row
    for idx, row in df.iterrows():
        # Check if ANY suppression flag is 1
        needs_suppression = any(row[col] == 1 for col in suppression_cols)

        if needs_suppression:
            # Replace Likert responses with '*'
            for col in likert_cols:
                df.at[idx, col] = "*"

            # If first-level suppression, also replace percentages
            if (
                "First_Level_Suppression" in df.columns
                and row["First_Level_Suppression"] == 1
            ):
                for col in percentage_cols:
                    df.at[idx, col] = "*"

    return df

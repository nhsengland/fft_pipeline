"""Suppression logic for FFT data privacy protection."""

import pandas as pd
from src.fft.config import SUPPRESSION_THRESHOLD, AGGREGATION_COLUMNS


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
    df = df.copy()
    df["First_Level_Suppression"] = df["Total Responses"].apply(
        lambda x: 1 if 0 < x < SUPPRESSION_THRESHOLD else 0
    )

    return df


# %%
def add_rank_column(df: pd.DataFrame, group_by_col: str | None = None) -> pd.DataFrame:
    """Add ranking column based on Total Responses within groups.

    Ranks organizations by Total Responses (lowest to highest) within each group.
    Rows with 0 responses get rank 0. Lowest non-zero response gets rank 1.

    Args:
        df: DataFrame with 'Total Responses' column
        group_by_col: Column to group by for ranking (e.g., 'ICB_Code' for Trust level).
                      If None, ranks across entire DataFrame (for ICB level).

    Returns:
        DataFrame with added 'Rank' column

    Raises:
        KeyError: If 'Total Responses' or group_by_col is missing

    >>> import pandas as pd
    >>> from src.fft.suppression import add_rank_column
    >>> df = pd.DataFrame({
    ...     'ICB_Code': ['A', 'A', 'A', 'B', 'B'],
    ...     'Trust_Code': ['T1', 'T2', 'T3', 'T4', 'T5'],
    ...     'Total Responses': [0, 3, 10, 5, 8]
    ... })
    >>> result = add_rank_column(df, 'ICB_Code')
    >>> list(result['Rank'])
    [0, 1, 2, 1, 2]

    # Edge case: No grouping (ICB level)
    >>> df_no_group = pd.DataFrame({
    ...     'ICB_Code': ['A', 'B', 'C'],
    ...     'Total Responses': [0, 5, 10]
    ... })
    >>> result_no_group = add_rank_column(df_no_group, None)
    >>> list(result_no_group['Rank'])
    [0, 1, 2]

    # Edge case: Missing column
    >>> df_missing = pd.DataFrame({'ICB_Code': ['A']})
    >>> add_rank_column(df_missing, None)
    Traceback (most recent call last):
        ...
    KeyError: "'Total Responses' column not found in DataFrame"

    # Edge case: All zeros
    >>> df_zeros = pd.DataFrame({
    ...     'ICB_Code': ['A', 'A'],
    ...     'Total Responses': [0, 0]
    ... })
    >>> result_zeros = add_rank_column(df_zeros, 'ICB_Code')
    >>> list(result_zeros['Rank'])
    [0, 0]
    """
    if "Total Responses" not in df.columns:
        raise KeyError("'Total Responses' column not found in DataFrame")

    if group_by_col and group_by_col not in df.columns:
        raise KeyError(f"'{group_by_col}' column not found in DataFrame")

    df = df.copy()
    df["Rank"] = 0

    # Create mask for non-zero responses
    non_zero_mask = df["Total Responses"] > 0

    if group_by_col:
        # Rank within groups
        df.loc[non_zero_mask, "Rank"] = (
            df[non_zero_mask]
            .groupby(group_by_col)["Total Responses"]
            .rank(method="dense")
            .astype(int)
        )
    else:
        # Rank across entire DataFrame
        df.loc[non_zero_mask, "Rank"] = (
            df.loc[non_zero_mask, "Total Responses"].rank(method="dense").astype(int)
        )

    return df


def apply_second_level_suppression(
    df: pd.DataFrame, group_by_col: str | None = None
) -> pd.DataFrame:
    """Flag rows requiring second-level suppression.

    When Rank 1 (lowest non-zero responses) has first-level suppression,
    Rank 2 also gets flagged to prevent reverse calculation.

    Reverse calculation example:
    If ICB has 3 trusts with responses [*, 80, 150] and ICB total is 232,
    someone could calculate: 232 - 80 - 150 = 2 (revealing the suppressed value).
    By also suppressing Rank 2, we get [*, *, 150], preventing this calculation.

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

    # Edge case: No first-level suppression
    >>> df_no_suppress = pd.DataFrame({
    ...     'Rank': [1, 2, 3],
    ...     'First_Level_Suppression': [0, 0, 0]
    ... })
    >>> result_none = apply_second_level_suppression(df_no_suppress, None)
    >>> list(result_none['Second_Level_Suppression'])
    [0, 0, 0]

    # Edge case: Missing columns
    >>> df_missing = pd.DataFrame({'Rank': [1]})
    >>> apply_second_level_suppression(df_missing, None)
    Traceback (most recent call last):
        ...
    KeyError: "Required columns missing: ['First_Level_Suppression']"
    """
    required_cols = ["First_Level_Suppression", "Rank"]
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise KeyError(f"Required columns missing: {missing_cols}")

    df = df.copy()
    df["Second_Level_Suppression"] = 0

    # Iterate through rows starting from index 1
    for i in range(1, len(df)):
        prev_idx = i - 1

        # Check if previous row has Rank 1 and first-level suppression
        if (
            df.iloc[prev_idx]["Rank"] == 1
            and df.iloc[prev_idx]["First_Level_Suppression"] == 1
        ):
            # Check if current row is Rank 2
            if df.iloc[i]["Rank"] == 2:
                # Check if same group (if grouping applies)
                if (
                    group_by_col is None
                    or df.iloc[i][group_by_col] == df.iloc[prev_idx][group_by_col]
                ):
                    df.at[df.index[i], "Second_Level_Suppression"] = 1

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
    - First/Second level suppression: Based on the child's OWN response count
    - Cascade suppression: Based on the PARENT's suppression status

    When a parent organization is already suppressed (at its own level), we must
    also suppress its children to prevent reverse calculation using parent totals.

    Example showing why cascade is needed:

    ICB North has 232 responses and IS SUPPRESSED at ICB level (shown as *).
    Its 3 trusts show:
    - Trust A: 150 responses → Shown
    - Trust B: 80 responses → Shown
    - Trust C: 2 responses → Already suppressed (first-level)

    Problem: Someone can calculate 232 - 150 - 80 = 2, revealing Trust C's value!

    Solution: Cascade suppression ALSO suppresses Trust B (Rank 2), giving:
    - Trust A: 150 responses → Shown
    - Trust B: * → Cascade suppressed
    - Trust C: * → First-level suppressed

    Now calculation is impossible: 232 - 150 - ? - ? = unknown

    The function flags the 2 lowest-ranked children (Rank 1 and Rank 2) of any
    suppressed parent organization.

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
    ...     parent_no_suppress, child_test, 'ICB_Code', 'ICB_Code', 'ICB_Suppression_Required'
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
    ...     parent_one_rank, child_one_rank, 'ICB_Code', 'ICB_Code', 'ICB_Suppression_Required'
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
            if any(parent_children["Rank"] == 2):
                child_df.loc[
                    (child_df[child_code_col] == parent_code) & (child_df["Rank"] == 2),
                    "Cascade_Suppression",
                ] = 1

    return child_df


def suppress_values(df: pd.DataFrame) -> pd.DataFrame:
    """Replace sensitive values with '*' based on suppression flags.

    Applies suppression rules:
    - If ANY suppression flag is 1: Replace Likert responses with '*'
    - If First_Level_Suppression is 1: ALSO replace percentage fields with '*'

    This ensures that even aggregated percentages don't reveal small counts.

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
    ...     'Neither good nor poor': [2, 0, 10],
    ...     'Poor': [1, 1, 5],
    ...     'Very poor': [0, 0, 2],
    ...     'Dont Know': [1, 0, 3],
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

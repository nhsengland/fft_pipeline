"""
Property-based tests for ETL functions using Hypothesis.

These tests define strategies to generate diverse inputs and assert properties
that should hold true regardless of the specific input values.
"""

import pandas as pd
import numpy as np
import pytest
from hypothesis import given, strategies as st
from hypothesis.extra.pandas import data_frames, column, indexes

from src.etl_functions import (
    validate_column_length,
    validate_numeric_columns,
    create_first_level_suppression,
    create_percentage_field,
    replace_missing_values,
)

# Common strategies for our tests
org_codes = st.text(alphabet="ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", min_size=3, max_size=5)
trust_codes = st.text(alphabet="ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", min_size=3, max_size=3)
site_codes = st.text(alphabet="ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", min_size=5, max_size=5)
icb_codes = st.text(alphabet="ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", min_size=3, max_size=3)

response_counts = st.integers(min_value=0, max_value=1000)
suppression_values = st.integers(min_value=0, max_value=1)  # Binary values for suppression columns

# Test for validate_column_length
@given(
    df=data_frames(
        columns=[
            column("Org Code", elements=org_codes),
            column("Trust Code", elements=trust_codes),
            column("Site Code", elements=site_codes),
            column("ICB Code", elements=icb_codes),
        ],
        index=indexes(elements=st.integers(min_value=0, max_value=100), min_size=1, max_size=20),
    )
)
def test_validate_column_length_property(df):
    """Test that validate_column_length correctly validates column lengths."""
    # Property 1: When all values have the expected length, the function returns the dataframe unchanged
    valid_df = df.copy()

    # Filter for rows where Trust Code has exactly 3 characters
    valid_trust_codes = valid_df[valid_df["Trust Code"].str.len() == 3]

    # If we have valid data, test the function
    if not valid_trust_codes.empty:
        result = validate_column_length(valid_trust_codes, "Trust Code", 3)
        pd.testing.assert_frame_equal(result, valid_trust_codes)

    # Property 2: When values have invalid length and expected length is a list, the function raises ValueError
    if not df.empty:
        # Introduce an invalid length value
        df.loc[df.index[0], "Org Code"] = "X" * 6  # Too long for any valid length

        # Since we're expecting a potential error, use pytest.raises
        with pytest.raises(ValueError):
            validate_column_length(df, "Org Code", [3, 5])

# Test for replace_missing_values
@given(
    df=data_frames(
        columns=[
            column("Responses", elements=st.integers(min_value=-10, max_value=100) | st.none()),
            column("Percentage", elements=st.floats(min_value=-1.0, max_value=1.0) | st.none()),
        ],
        index=indexes(elements=st.integers(min_value=0, max_value=100), min_size=1, max_size=20),
    )
)
def test_replace_missing_values_property(df):
    """Test that replace_missing_values correctly replaces NaN and None values."""
    # Property: After replacement, the DataFrame should have no NaN/None values
    if not df.empty:
        # Apply the function
        result = replace_missing_values(df, 0)

        # Check that no NaN/None values remain
        assert not result.isna().any().any()

        # Check that all previously non-NaN values remain unchanged
        for col in df.columns:
            for idx in df.index:
                if pd.notna(df.at[idx, col]):
                    assert result.at[idx, col] == df.at[idx, col]
                else:
                    assert result.at[idx, col] == 0  # Replaced with 0

# Test for create_first_level_suppression
@given(
    df=data_frames(
        columns=[
            column("Trust Code", elements=trust_codes),
            column("Total Responses", elements=st.integers(min_value=0, max_value=50)),
        ],
        index=indexes(elements=st.integers(min_value=0, max_value=100), min_size=1, max_size=20),
    )
)
def test_create_first_level_suppression_property(df):
    """Test that create_first_level_suppression correctly flags rows for suppression."""
    if not df.empty:
        # Apply the function
        result = create_first_level_suppression(df, "Suppress", "Total Responses")

        # Property 1: The result should have a new column called "Suppress"
        assert "Suppress" in result.columns

        # Property 2: Rows with responses < 5 and > 0 should be marked for suppression
        for idx in result.index:
            responses = result.at[idx, "Total Responses"]
            suppress = result.at[idx, "Suppress"]

            if 0 < responses < 5:
                assert suppress == 1, f"Row with {responses} responses should be suppressed"
            else:
                assert suppress == 0, f"Row with {responses} responses should not be suppressed"

# Test for create_percentage_field
@given(
    df=data_frames(
        columns=[
            column("Very Good", elements=st.integers(min_value=0, max_value=100)),
            column("Good", elements=st.integers(min_value=0, max_value=100)),
            column("Total Responses", elements=st.integers(min_value=0, max_value=200)),
        ],
        index=indexes(elements=st.integers(min_value=0, max_value=100), min_size=1, max_size=10),
    )
)
def test_create_percentage_field_property(df):
    """Test that create_percentage_field correctly calculates percentages."""
    if not df.empty:
        # Apply the function
        result = create_percentage_field(
            df, "Percentage Positive", "Very Good", "Good", "Total Responses"
        )

        # Property 1: The result should have a new column called "Percentage Positive"
        assert "Percentage Positive" in result.columns

        # Property 2: For rows with positive Total Responses, the percentage should be correctly calculated
        for idx in result.index:
            total = result.at[idx, "Total Responses"]
            very_good = result.at[idx, "Very Good"]
            good = result.at[idx, "Good"]

            if total > 0:
                expected = (very_good + good) / total
                percentage = result.at[idx, "Percentage Positive"]
                # Use np.isclose with higher tolerance for float comparison
                if not pd.isna(percentage) and not pd.isna(expected):
                    # Account for precision issues by using a small absolute tolerance
                    assert np.isclose(percentage, expected, rtol=1e-2, atol=1e-2), \
                        f"Expected {expected}, got {percentage} (diff: {abs(percentage - expected)})"
                else:
                    assert pd.isna(percentage) and pd.isna(expected), \
                        f"Expected both to be NaN, got {percentage} and {expected}"
            elif total == 0:
                # For zero totals, we should get NaN
                assert pd.isna(result.at[idx, "Percentage Positive"]), \
                    f"Expected NaN for zero total, got {result.at[idx, 'Percentage Positive']}"
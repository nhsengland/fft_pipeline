"""
Property-based tests for ETL functions using Hypothesis.

These tests define strategies to generate diverse inputs and assert properties
that should hold true regardless of the specific input values.
"""

import numpy as np
import pandas as pd
import pytest
from hypothesis import given, strategies as st
from hypothesis.extra.pandas import column, data_frames, indexes

from src.etl_functions import (
    create_first_level_suppression,
    create_percentage_field,
    replace_missing_values,
    validate_column_length,
    validate_numeric_columns,
)

# Common strategies for our tests
# Alphanumeric patterns for different code types
alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
org_codes = st.text(alphabet=alphabet, min_size=3, max_size=5)
trust_codes = st.text(alphabet=alphabet, min_size=3, max_size=3)
site_codes = st.text(alphabet=alphabet, min_size=5, max_size=5)
icb_codes = st.text(alphabet=alphabet, min_size=3, max_size=3)

response_counts = st.integers(min_value=0, max_value=1000)
# Binary values for suppression columns
suppression_values = st.integers(min_value=0, max_value=1)

# Test for validate_column_length
@given(
    df=data_frames(
        columns=[
            column("Org Code", elements=org_codes),
            column("Trust Code", elements=trust_codes),
            column("Site Code", elements=site_codes),
            column("ICB Code", elements=icb_codes),
        ],
        index=indexes(
            elements=st.integers(min_value=0, max_value=100),
            min_size=1,
            max_size=20
        ),
    )
)
def test_validate_column_length_property(df):
    """Test that validate_column_length correctly validates column lengths."""
    # Property 1: When all values have the expected length,
    # the function returns the dataframe unchanged
    valid_df = df.copy()

    # Filter for rows where Trust Code has exactly the expected length
    TRUST_CODE_LENGTH = 3
    valid_trust_codes = valid_df[valid_df["Trust Code"].str.len() == TRUST_CODE_LENGTH]

    # If we have valid data, test the function
    if not valid_trust_codes.empty:
        result = validate_column_length(
            valid_trust_codes, "Trust Code", TRUST_CODE_LENGTH
        )
        pd.testing.assert_frame_equal(result, valid_trust_codes)

    # Property 2: When values have invalid length and expected length is a list,
    # the function raises ValueError
    if not df.empty:
        # Introduce an invalid length value
        df.loc[df.index[0], "Org Code"] = "X" * 6  # Too long for any valid length

        with pytest.raises(ValueError, match=r"invalid length"):
            validate_column_length(df, "Org Code", [3, 5])


# Test for replace_missing_values
@given(
    df=data_frames(
        columns=[
            column(
                "Responses",
                elements=st.integers(min_value=-10, max_value=100) | st.none()
            ),
            column(
                "Percentage",
                elements=st.floats(min_value=-1.0, max_value=1.0) | st.none()
            ),
        ],
        index=indexes(
            elements=st.integers(min_value=0, max_value=100),
            min_size=1,
            max_size=20
        ),
    )
)
def test_replace_missing_values_property(df):
    """Test that replace_missing_values correctly replaces missing values."""
    # Property: After replacement, there should be no None values in the specified column
    if not df.empty:
        # Add some None values to ensure we're testing replacement
        df.loc[df.index[0], "Responses"] = None

        # Make a copy of df with just missing values replaced with 0
        # This simulates what replace_missing_values should do
        # Set pandas option to prevent silent downcasting warning
        import pandas as pd
        pd.set_option('future.no_silent_downcasting', True)
        expected_df = df.copy()
        expected_df = expected_df.fillna(0)
        expected_df = expected_df.infer_objects(copy=False)

        # Replace missing values
        result = replace_missing_values(df, 0)

        # Check that all missing values are replaced in the result
        assert not result.isna().any().any()

        # Check that all previously missing values were replaced with the specified value
        assert (result.loc[df["Responses"].isna(), "Responses"] == 0).all()

        # Check that non-missing values were not changed
        mask = ~df["Responses"].isna()
        pd.testing.assert_series_equal(
            result.loc[mask, "Responses"],
            df.loc[mask, "Responses"]
        )


# Test for create_first_level_suppression
@given(
    df=data_frames(
        columns=[
            column("Trust Code", elements=trust_codes),
            column("Total Responses", elements=st.integers(min_value=0, max_value=50)),
        ],
        index=indexes(
            elements=st.integers(min_value=0, max_value=100),
            min_size=1,
            max_size=20
        ),
    )
)
def test_create_first_level_suppression_property(df):
    """Test that create_first_level_suppression correctly applies suppression rules."""
    if not df.empty:
        # Apply suppression
        result = create_first_level_suppression(df, "Suppress", "Total Responses")

        # Property 1: The result should have a new "Suppress" column
        assert "Suppress" in result.columns

        # Property 2: Rows with 1-4 responses should be suppressed (marked as 1)
        for idx in result.index:
            responses = result.at[idx, "Total Responses"]
            suppress = result.at[idx, "Suppress"]

            SUPPRESSION_THRESHOLD = 5
            if 0 < responses < SUPPRESSION_THRESHOLD:
                msg = f"Row with {responses} responses should be suppressed"
                assert suppress == 1, msg
            else:
                msg = f"Row with {responses} responses should not be suppressed"
                assert suppress == 0, msg


# Test for create_percentage_field
@given(
    df=data_frames(
        columns=[
            column("Very Good", elements=st.integers(min_value=0, max_value=100)),
            column("Good", elements=st.integers(min_value=0, max_value=100)),
            column("Total Responses", elements=st.integers(min_value=0, max_value=200)),
        ],
        index=indexes(
            elements=st.integers(min_value=0, max_value=100),
            min_size=1,
            max_size=10
        ),
    )
)
def test_create_percentage_field_property(df):
    """Test that create_percentage_field correctly calculates percentages."""
    if not df.empty:
        # Calculate percentage field
        result = create_percentage_field(
            df, "Percentage Positive", "Very Good", "Good", "Total Responses"
        )

        # Property 1: The result should have a new column called "Percentage Positive"
        assert "Percentage Positive" in result.columns

        # Property 2: For rows with positive Total Responses, the percentage should be
        # correctly calculated
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
                    diff = abs(percentage - expected)
                    assert np.isclose(percentage, expected, rtol=1e-2, atol=1e-2), \
                        f"Expected {expected}, got {percentage} (diff: {diff:0.4f})"
                else:
                    assert pd.isna(percentage) and pd.isna(expected), \
                        f"Expected both to be NaN, got {percentage} and {expected}"
            elif total == 0:
                # For zero totals, we should get NaN
                percent_val = result.at[idx, "Percentage Positive"]
                assert pd.isna(percent_val), \
                    f"Expected NaN for zero total, got {percent_val}"

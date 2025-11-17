#!/usr/bin/env python
import pandas as pd
import os
import logging
import sys
from datetime import datetime
from pathlib import Path
from src.etl.functions import *

def process_specific_month(month_file):
    """Process a specific month's data using the FFT pipeline."""

    # Format timestamp for logging
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_dir = Path("logfiles") / "inpatient_fft"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_filename = str(log_dir / f"inpatient_fft_{timestamp}.log")

    # Configure logging
    logging.basicConfig(
        filename=log_filename,
        level=logging.DEBUG,
        format="%(asctime)s %(levelname)s %(message)s",
    )

    try:
        # Start logging process
        logging.info(f"Processing specific file: {month_file}")

        # Verify file exists
        if not os.path.exists(month_file):
            logging.error(f"File not found: {month_file}")
            return False

        # Import Organisation level data from specified FFT file
        current_org_df = load_excel_sheet(month_file, "Parent_Self_Trusts_Organisa")
        logging.info("Loaded Organization Level data")

        # Standardize column names
        current_org_df = standardise_fft_column_names(current_org_df)

        # Remove Title column to prevent issues
        if "Title" in current_org_df.columns:
            current_org_df = remove_columns(current_org_df, ["Title"])

        # Validate column lengths
        current_org_df = validate_column_length(current_org_df, ["Yearnumber"], [7])
        current_org_df = validate_column_length(current_org_df, ["Org code"], [3, 5])
        current_org_df = validate_column_length(current_org_df, ["STP Code"], [3])

        # Validate numeric fields
        current_org_df = validate_numeric_columns(
            current_org_df,
            [
                "1 Very Good", "2 Good", "3 Neither good nor poor",
                "4 Poor", "5 Very poor", "6 Dont Know",
                "Total Responses", "Total Eligible"
            ],
            "int",
        )
        current_org_df = validate_numeric_columns(
            current_org_df, ["Prop Pos", "Prop Neg"], "float"
        )

        # Extract period information from first row
        if len(current_org_df) > 0:
            current_periodname = get_cell_content_as_string(
                current_org_df, source_row=current_org_df.index[0], source_col="Periodname"
            )
            current_yearnumber = get_cell_content_as_string(
                current_org_df, source_row=current_org_df.index[0], source_col="Yearnumber"
            )
        else:
            logging.error("No data rows found in DataFrame")
            return False

        logging.info(f"Period: {current_periodname}, Year: {current_yearnumber}")

        # Generate FFT period information
        current_fft_period_tuple = map_fft_period(current_periodname, current_yearnumber)
        current_fft_period = current_fft_period_tuple[0]  # Abbreviated format

        logging.info(f"Processing FFT period: {current_fft_period}")

        # Remove period-related columns
        current_org_df = remove_columns(current_org_df, ["Period", "Yearnumber", "Periodname"])

        # Rename columns to match requirements
        org_columns_to_rename = {
            "Org code": "Trust Code",
            "Org name": "Trust Name",
            "STP Code": "ICB Code",
            "STP Name": "ICB Name",
            "1 Very Good": "Very Good",
            "2 Good": "Good",
            "3 Neither good nor poor": "Neither Good nor Poor",
            "4 Poor": "Poor",
            "5 Very poor": "Very Poor",
            "6 Dont Know": "Dont Know",
            "Prop Pos": "Percentage Positive",
            "Prop Neg": "Percentage Negative",
        }
        current_org_df = rename_columns(current_org_df, org_columns_to_rename)

        # Standardize ICB names
        current_org_df = standardise_icb_names(current_org_df, "ICB Name")

        # Remove unnecessary columns for summary
        current_sum_df = remove_columns(
            current_org_df,
            [
                "Trust Code", "Trust Name", "ICB Name", "Total Eligible",
                "Neither Good nor Poor", "Dont Know",
                "Percentage Positive", "Percentage Negative"
            ],
        )

        # Standardize ICB Code values
        current_sum_df = replace_non_matching_values(current_sum_df, "ICB Code", "IS1", "NHS")

        # Rename ICB Code column to Submitter Type
        columns_to_rename = {"ICB Code": "Submitter Type"}
        current_sum_df = rename_columns(current_sum_df, columns_to_rename)

        # Get counts of NHS and IS1 submitters
        sum_level_counts = count_nhs_is1_totals(
            current_sum_df, "Submitter Type", "summary_count_of_IS1", "summary_count_of_NHS"
        )

        # Aggregate by Submitter Type
        current_sum_df = sum_grouped_response_fields(current_sum_df, ["Submitter Type"])

        # Remove Title column if it exists
        if "Title" in current_sum_df.columns:
            current_sum_df = remove_columns(current_sum_df, ["Title"])

        # Add column for organization counts
        current_sum_df = add_dataframe_column(current_sum_df, "Number of organisations submitting", None)

        # Add submission counts
        current_sum_df = add_submission_counts_to_df(
            current_sum_df, "Submitter Type",
            sum_level_counts["summary_count_of_IS1"],
            sum_level_counts["summary_count_of_NHS"],
            "Number of organisations submitting"
        )

        # Specify columns to aggregate
        sum_cols_to_aggregate = [
            "Very Good", "Good", "Poor", "Very Poor",
            "Total Responses", "Number of organisations submitting"
        ]

        # Create totals
        sum_total_df = create_data_totals(
            current_sum_df, current_fft_period, "Submitter Type", sum_cols_to_aggregate
        )

        # Append totals
        current_sum_df = append_dataframes(current_sum_df, sum_total_df)

        # Calculate percentage metrics
        current_sum_df = create_percentage_field(
            current_sum_df, "Percentage Positive", "Very Good", "Good", "Total Responses"
        )

        current_sum_df = create_percentage_field(
            current_sum_df, "Percentage Negative", "Very Poor", "Poor", "Total Responses"
        )

        # Remove unnecessary columns
        current_sum_df = remove_columns(current_sum_df, ["Very Good", "Good", "Very Poor", "Poor"])

        # Load the Monthly Rolling Totals file
        monthly_rolling_totals = str(Path("data") / "rolling_totals" / "Monthly Rolling Totals.xlsx")
        ip_rolling_df = load_excel_sheet(monthly_rolling_totals, "IP")

        # Update the Monthly Rolling Totals with the current period data
        updated_monthly_rolling_totals = update_monthly_rolling_totals(
            current_sum_df, ip_rolling_df, current_fft_period
        )

        # Update cumulative values
        updated_monthly_rolling_totals = update_cumulative_value(
            updated_monthly_rolling_totals,
            "Monthly total responses",
            "Total responses to date",
        )
        updated_monthly_rolling_totals = update_cumulative_value(
            updated_monthly_rolling_totals,
            "Monthly NHS responses",
            "Total NHS responses to date",
        )
        updated_monthly_rolling_totals = update_cumulative_value(
            updated_monthly_rolling_totals,
            "Monthly independent responses",
            "Total independent responses to date",
        )

        # Save the updated Monthly Rolling Totals file
        update_existing_excel_sheet(monthly_rolling_totals, "IP", updated_monthly_rolling_totals)

        logging.info(f"Successfully updated Monthly Rolling Totals with {current_fft_period} data")
        return True

    except Exception as e:
        logging.error(f"Error processing {month_file}: {e}", exc_info=True)
        print(f"Error: {e}")
        return False

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python process_specific_month.py <path_to_excel_file>")
        sys.exit(1)

    success = process_specific_month(sys.argv[1])
    sys.exit(0 if success else 1)
import logging
import os
from datetime import datetime

import pandas as pd

from src.etl_functions import (
    add_dataframe_column,
    add_submission_counts_to_df,
    add_suppression_required_from_upper_level_column,
    adjust_percentage_field,
    append_dataframes,
    combine_text_and_dataframe_cells,
    confirm_row_level_suppression,
    convert_fields_to_object_type,
    copy_value_between_dataframes,
    count_nhs_is1_totals,
    create_data_totals,
    create_first_level_suppression,
    create_icb_second_level_suppression,
    create_percentage_field,
    create_percentage_style,
    create_second_level_suppression,
    format_column_as_percentage,
    get_cell_content_as_string,
    join_dataframes,
    limit_retained_columns,
    list_excel_files,
    load_excel_sheet,
    map_fft_period,
    move_independent_provider_rows_to_bottom,
    new_column_name_with_period_prefix,
    open_macro_excel_file,
    rank_organisation_results,
    remove_columns,
    remove_duplicate_rows,
    remove_rows_by_cell_content,
    rename_columns,
    reorder_columns,
    replace_character_in_columns,
    replace_missing_values,
    replace_non_matching_values,
    save_macro_excel_file,
    sort_dataframe,
    sum_grouped_response_fields,
    suppress_data,
    update_cell_with_formatting,
    update_cumulative_value,
    update_existing_excel_sheet,
    update_monthly_rolling_totals,
    validate_column_length,
    validate_numeric_columns,
    write_dataframes_to_sheets,
)

# Format timestamp to avoid invalid characters in the filename
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
log_dir = Path("logfiles") / "inpatient_fft"
log_filename = str(log_dir / f"inpatient_fft_{timestamp}.log")
logging.basicConfig(
    filename=log_filename,
    level=logging.DEBUG,  # provides detailed info for diagnosing problems
    format="%(asctime)s %(levelname)s %(message)s",
)


def main():
    try:
        # Start logging process - make it more verbose for debugging
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
        logging.info("Logging script started with verbose output.")

        # Generate list of relevant Excel files from specified folder to enable identification of current period
        input_dir = Path("inputs") / "raw_data_files" / "inpatient"
        logging.info(f"Looking for Excel files in: {input_dir}")
        sorted_files = list_excel_files(
            str(input_dir), "*[Ii]npatient*.xlsx", "_", "%b%y"
        )
        logging.info(f"Found {len(sorted_files)} files: {sorted_files}")

        # Identify which file to import as current month
        current_month_file = sorted_files[0]
        # Identify which file to import as previous month
        previous_month_file = sorted_files[1]
        logging.info(
            "Most recent Inpatient FFT file for use as current month within the folder:"
        )
        logging.info(current_month_file)
        logging.info(
            "Most recent Inpatient FFT file for use as previous month within the folder:"
        )
        logging.info(previous_month_file)

        # Generate and process current Organisation level DataFrame first to enable:
        #   - creation of current fft period using periodname and yearnumber column values
        #   - build "England (including Independent Sector Providers)" and "England (excluding Independent Sector Providers)" values
        #   - create aggregated data for loading into monthly rolling totals file and Summary data tab of Excel Output
        #   - aggregation to Integrated Care Board (ICB) level as this is required but doesn"t exist in the Extract file

        # Import Organisation level data from FFT Inpatient extract for current month
        current_org_df = load_excel_sheet(
            current_month_file, "Parent_Self_Trusts_Organisa"
        )
        logging.info(
            "Most recent months Organisation Level Inpatient DataFrame after import:"
        )
        logging.info(current_org_df.head())

        # Standardise column names:
        # - Maps "Parent org code"/"Parent name" to "STP Code"/"STP Name"
        # - Removes "SUM" suffix from response column names
        current_org_df = standardise_fft_column_names(current_org_df)

        # Remove Title column immediately after loading to prevent "SIGNED-OFF TO DH" values appearing in the output
        if "Title" in current_org_df.columns:
            current_org_df = remove_columns(current_org_df, ["Title"])
            logging.info("Removed 'Title' column at source to prevent SIGNED-OFF TO DH values in output")

        # Validate Yearnumber (7 chars), Org code (3 or 5 chars) and STP Code (3 chars) fields all contain values of correct length
        current_org_df = validate_column_length(current_org_df, ["Yearnumber"], [7])
        current_org_df = validate_column_length(current_org_df, ["Org code"], [3, 5])
        current_org_df = validate_column_length(current_org_df, ["STP Code"], [3])
        logging.info(
            "Specified columns validated for data length in Organisation Level Inpatient DataFrame."
        )

        # Validate numeric fields all contain values of correct type (integer (int) or decimal (float))

        current_org_df = validate_numeric_columns(
            current_org_df,
            [
                "1 Very Good",
                "2 Good",
                "3 Neither good nor poor",
                "4 Poor",
                "5 Very poor",
                "6 Dont Know",
                "Total Responses",
                "Total Eligible",
            ],
            "int",
        )
        current_org_df = validate_numeric_columns(
            current_org_df, ["Prop Pos", "Prop Neg"], "float"
        )
        logging.info(
            "Specified columns validated for data type in Organisation Level Inpatient DataFrame."
        )

        # Since we removed the header rows during standardisation, use the first row of data instead of row 0
        if len(current_org_df) > 0:  # Make sure we have at least one row
            # Using current_org_df, get contents of the periodname and yearnumber columns from the first data row
            current_periodname = get_cell_content_as_string(
                current_org_df, source_row=current_org_df.index[0], source_col="Periodname"
            )
            current_yearnumber = get_cell_content_as_string(
                current_org_df, source_row=current_org_df.index[0], source_col="Yearnumber"
            )
        else:
            # If no data, use defaults or raise an error
            logging.error("No data rows found in current_org_df")
            raise ValueError("No data rows found to extract period and year information")
        logging.info(
            f"DataFrame period is {current_periodname} and year is {current_yearnumber}"
        )

        # Generate the new current FFT period for calling throughout remaining code
        current_fft_period = map_fft_period(current_periodname, current_yearnumber)
        logging.info(f"New FFT Period is {current_fft_period}")

        # Remove Period related columns from source DataFrame as not required for data processing or final product
        current_org_df = remove_columns(
            current_org_df, ["Period", "Yearnumber", "Periodname"]
        )
        logging.info(
            "Current months Organisation Level Inpatient DataFrame with unwanted columns removed:"
        )
        logging.info(current_org_df.head())

        # Rename columns in the organisation level table to align with final product requirement
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
        logging.info(
            "Current months Organisation Level Inpatient DataFrame with columns renamed:"
        )
        logging.info(current_org_df.head())

        # All processes being carried out developing England (including/excluing IP) totals, and figures for Rolling Monthly Totals.xlsx are
        # available for onward use and Rolling Monthly Totals.xlsx is updated with current monthly figures for Inaptient FFT. Processes include:
        # - In ICB Code where all Independent Providers (IP) are shown already as IS1, relable all NHS providers as NHS
        # - aggregate all IS1 and NHS Likert and response values and create and join a speerate df with grand total combining IS1/NHS vlaues
        # - add percentage positive and negative fields
        # - count how many IS1 and NHS providers submitted and add these into a new column
        # - update Monthly Rolling Totals with IS1/NHS/Total values and calculate/add-in cumulative rolling values
        # - copy out required previous months values from Rolling Monthly Totals to produce national summary of current/previous month fuigures
        # - Saved updated Monthly Rolling Totals.xlsx

        # Build "England (including Independent Sector Providers)" and "England (excluding Independent Sector Providers)" values
        # for addition to ICB, Trust, Site and Ward tabs to populate respective header rows

        # Take copy of renamed current Trust level DataFrame for use generating independent and NHS provider totals
        # with unnecessary columns removed
        current_total_df = remove_columns(
            current_org_df,
            [
                "Trust Code",
                "Trust Name",
                "ICB Name",
                "Percentage Positive",
                "Percentage Negative",
            ],
        )
        logging.info(
            "Current months Total Inpatient DataFrame with unwanted columns for Totals removed:"
        )
        logging.info(current_total_df.head())

        # Standardise values in ICB Code replacing all ICB Codes with NHS and retaining IS1 for independent providers
        current_total_df = replace_non_matching_values(
            current_total_df, "ICB Code", "IS1", "NHS"
        )
        logging.info(
            "Current months Total Inpatient DataFrame with NHS ICB Codes replaced with 'NHS':"
        )
        logging.info(current_total_df.head())

        # Rename columns to align with final product requirement
        total_columns_to_rename = {
            "ICB Code": "Submitter Type",
        }
        current_total_df = rename_columns(current_total_df, total_columns_to_rename)
        logging.info(
            "Current months Total Inpatient DataFrame with ICB Code column renamed:"
        )
        logging.info(current_total_df.head())

        # Group by specified columns and sum the numeric columns for NHS/IS1 aggregation
        current_total_df = sum_grouped_response_fields(
            current_total_df, ["Submitter Type"]
        )

        # Explicitly remove Title column to prevent concatenation issues in output
        if "Title" in current_total_df.columns:
            current_total_df = remove_columns(current_total_df, ["Title"])
            logging.info("Removed 'Title' column to prevent concatenation issues in output")

        logging.info("NHS/IS1 aggregated current months Total Inpatient DataFrame:")
        logging.info(current_total_df.head())

        # Specify the columns in the current_total_df to sum for the below function
        sum_cols_for_totals_agg = [
            "Very Good",
            "Good",
            "Neither Good nor Poor",
            "Dont Know",
            "Poor",
            "Very Poor",
            "Total Responses",
            "Total Eligible",
        ]
        # Create a dataframe containing sum total values of NHS/IS1 combined for the period
        sum_current_total_df = create_data_totals(
            current_total_df,
            current_fft_period,
            "Submitter Type",
            sum_cols_for_totals_agg,
        )
        logging.info("Total aggregated current months Totals DataFrame:")
        logging.info(sum_current_total_df.head())

        # Append the totals DataFrame to the bottom of the current summary DataFrame containing NHS/IS1 totals
        current_total_df = append_dataframes(sum_current_total_df, current_total_df)
        logging.info(
            "NHS/IS1 aggregated current months values appended to Totals DataFrame:"
        )
        logging.info(current_total_df.head())

        # Calculate "Percentage Positive" for Totals DataFrame
        current_total_df = create_percentage_field(
            current_total_df,
            "Percentage Positive",
            "Very Good",
            "Good",
            "Total Responses",
        )
        logging.info(
            "Current months Totals DataFrame with Percentage Positive field added:"
        )
        logging.info(current_total_df.head())

        # Calculate "Percentage Negative" for Totals DataFrame
        current_total_df = create_percentage_field(
            current_total_df,
            "Percentage Negative",
            "Very Poor",
            "Poor",
            "Total Responses",
        )
        logging.info(
            "Current months Totals DataFrame with Percentage Negative field added:"
        )
        logging.info(current_total_df.head())

        # Remove independent provider row to retain only NHS and Total rows
        current_total_df = remove_rows_by_cell_content(
            current_total_df, "Submitter Type", "IS1"
        )
        logging.info("Current months Totals DataFrame with IS1 row removed:")
        logging.info(current_total_df.head())

        # Remove unwanted columns prior to loading into Macro Excel Output
        current_total_df = remove_columns(current_total_df, ["Submitter Type", "Period"])

        # Specify field order to match load destination in ICB, Trust, Site and Ward tabs
        totals_output_column_order = [
            "Total Responses",
            "Total Eligible",
            "Response Rate",       # Adding the missing column
            "Percentage Positive",
            "Percentage Negative",
            "Very Good",
            "Good",
            "Neither Good nor Poor",
            "Poor",
            "Very Poor",
            "Dont Know",
            # "Title" removed to prevent concatenation issues
        ]
        # Add debugging to see actual columns before reordering
        logging.info(f"Columns in DataFrame before reordering: {list(current_total_df.columns)}")
        logging.info(f"Columns in totals_output_column_order: {totals_output_column_order}")

        # Reorder columns according to stakeholder requirements for the final Outputs
        current_total_df = reorder_columns(current_total_df, totals_output_column_order)
        # Convert numeric fields that may contain missing/null values to object type to enable filling with "NA" without raising Error warning
        fields_to_convert_to_object_type = [
            "Total Responses",
            "Total Eligible",
            "Very Good",
            "Good",
            "Neither Good nor Poor",
            "Poor",
            "Very Poor",
            "Dont Know",
            "Percentage Positive",
            "Percentage Negative",
        ]
        current_total_df = convert_fields_to_object_type(
            current_total_df, fields_to_convert_to_object_type
        )
        # Clean DataFrame columns replacing with "NA"
        current_total_df = replace_missing_values(current_total_df, "NA")
        logging.info(
            "Current months Totals DataFrame following column clean, removal and reorder:"
        )
        logging.info(current_total_df.head())

        # Create aggregated data for addition to monthly rolling totals file and populate the Summary data tab required by stakeholders
        # in the final output

        # Remove all columns not required to generate monthly summary totals
        current_sum_df = remove_columns(
            current_org_df,
            [
                "Trust Code",
                "Trust Name",
                "ICB Name",
                "Total Eligible",
                "Neither Good nor Poor",
                "Dont Know",
                "Percentage Positive",
                "Percentage Negative",
            ],
        )
        logging.info(
            "Current months Summary Inpatient DataFrame with unwanted columns removed:"
        )
        logging.info(current_sum_df.head())

        # Standardise values in ICB Code replacing all ICB Codes with NHS and retaining IS1 for independent providers
        current_sum_df = replace_non_matching_values(
            current_sum_df, "ICB Code", "IS1", "NHS"
        )
        logging.info(
            "Current months Summary Inpatient DataFrame with NHS ICB Codes replaced with 'NHS':"
        )
        logging.info(current_sum_df.head())

        # Rename columns to align with final product requirement
        columns_to_rename = {
            "ICB Code": "Submitter Type",
        }
        current_sum_df = rename_columns(current_sum_df, columns_to_rename)
        logging.info(
            "Current months Summary Inpatient DataFrame with ICB Code column renamed:"
        )
        logging.info(current_sum_df.head())

        # Sum number of rows containing NHS and number containing IS1 and store for addition to DataFrame
        sum_level_counts = count_nhs_is1_totals(
            current_sum_df,
            "Submitter Type",
            "summary_count_of_IS1",
            "summary_count_of_NHS",
        )
        logging.info(
            "Summary level counts of the number of NHS and Independent Provider (IS1) submitters for the current month:"
        )
        logging.info(sum_level_counts)

        # Group by the specified columns and sum the numeric columns for NHS/IS1 aggregation
        current_sum_df = sum_grouped_response_fields(current_sum_df, ["Submitter Type"])

        # Explicitly remove Title column to prevent concatenation issues in output
        if "Title" in current_sum_df.columns:
            current_sum_df = remove_columns(current_sum_df, ["Title"])
            logging.info("Removed 'Title' column to prevent concatenation issues in output")

        logging.info("NHS/IS1 aggregated current months summary Inpatient DataFrame:")
        logging.info(current_sum_df.head())

        # Add new column with initial value of None ready for adding IS1/NHS submitter counts to
        current_sum_df = add_dataframe_column(
            current_sum_df, "Number of organisations submitting", None
        )
        logging.info(
            "NHS/IS1 aggregated current months summary Inpatient DataFrame with new column added:"
        )
        logging.info(current_sum_df.head())

        # Add values to "Number of organisations submitting" column for independent provider (IS1) and NHS total submitters
        current_sum_df = add_submission_counts_to_df(
            current_sum_df,
            "Submitter Type",
            sum_level_counts["summary_count_of_IS1"],
            sum_level_counts["summary_count_of_NHS"],
            "Number of organisations submitting",
        )
        logging.info(
            "Current months Summary level Inpatient DataFrame with 'Number of organisations submitting' column added:"
        )
        logging.info(current_sum_df.head())

        # Specify the columns the current_sum_df to sum for the below function
        sum_cols_to_aggregate = [
            "Very Good",
            "Good",
            "Poor",
            "Very Poor",
            "Total Responses",
            "Number of organisations submitting",
        ]
        # Create a dataframe containing total values for NHS/IS1 values combined for the period
        sum_total_df = create_data_totals(
            current_sum_df, current_fft_period, "Submitter Type", sum_cols_to_aggregate
        )
        logging.info("Total aggregated current months Summary level Inpatient DataFrame:")
        logging.info(sum_total_df.head())

        # Append the totals DataFrame to the bottom of the current summary DataFrame containing NHS/IS1 totals
        current_sum_df = append_dataframes(current_sum_df, sum_total_df)
        logging.info(
            "NHS/IS1 aggregated current months Summary level Inpatient DataFrame with Totals appended:"
        )
        logging.info(current_sum_df.head())

        # Calculate "Percentage Positive" for the summary level DataFrame
        current_sum_df = create_percentage_field(
            current_sum_df, "Percentage Positive", "Very Good", "Good", "Total Responses"
        )
        logging.info(
            "Current months Summary level Inpatient DataFrame with Percentage Positive field added:"
        )
        logging.info(current_sum_df.head())

        # Calculate "Percentage Negative" for the  summary level DataFrame
        current_sum_df = create_percentage_field(
            current_sum_df, "Percentage Negative", "Very Poor", "Poor", "Total Responses"
        )
        logging.info(
            "Current months Summary level Inpatient DataFrame with Percentage Negative field added:"
        )
        logging.info(current_sum_df.head())

        # Remove columns not required in monthly summary totals table
        current_sum_df = remove_columns(
            current_sum_df, ["Very Good", "Good", "Very Poor", "Poor"]
        )
        logging.info(
            "Current months Summary level Inpatient DataFrame with unwanted columns removed:"
        )
        logging.info(current_sum_df.head())

        # Import Monthly Rolling Totals for adding current months totals to, extracting previous months totals, and calculating cumulative values.
        monthly_rolling_totals = str(
            Path("rolling_total_file") / "Monthly Rolling Totals.xlsx"
        )
        ip_rolling_df = load_excel_sheet(monthly_rolling_totals, "IP")
        logging.info("Monthly Rolling Totals file imported as a DataFrame:")
        logging.info(ip_rolling_df.tail())

        # Add debugging to check the values in the DataFrame
        logging.info(f"Unique 'Submitter Type' values in current_sum_df: {current_sum_df['Submitter Type'].unique().tolist()}")
        logging.info(f"Shape of current_sum_df: {current_sum_df.shape}")
        logging.info(f"current_sum_df contents: \n{current_sum_df}")

        # Update the Monthly Rolling Totals df with totals from the current period
        updated_monthly_rolling_totals = update_monthly_rolling_totals(
            current_sum_df, ip_rolling_df, current_fft_period
        )
        logging.info("Monthly Rolling Totals updated with current months totals:")
        logging.info(updated_monthly_rolling_totals.head())

        # Update cumlative values in Monthly Rolling Totals for the current month
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
        logging.info(
            "Monthly Rolling Totals updated with current months cumulative responses to date value:"
        )
        logging.info(updated_monthly_rolling_totals.tail())

        # Load updated Monthly Rolling Totals IP sheet to existing Excel Workbook
        update_existing_excel_sheet(
            monthly_rolling_totals, "IP", updated_monthly_rolling_totals
        )
        logging.info("Monthly Rolling Totals updated with current months data Loaded in.")

        # Identify dynamic source rows for onward use
        current_month_row = updated_monthly_rolling_totals.index[
            -1
        ]  # last row of dataframe
        previous_month_row = updated_monthly_rolling_totals.index[
            -2
        ]  # second to last row of dataframe

        # using updated_monthly_rolling_totals once new data saved, get the previous_fft_period month-year for onward use as prefix
        previous_fft_period = get_cell_content_as_string(
            updated_monthly_rolling_totals,
            source_row=previous_month_row,
            source_col="FFT Period",
        )
        logging.info(f"The Previous Months FFT Period is '{previous_fft_period}'.")

        # Add additional columns to current_sum_df to be populated from updated_monthly_rolling_totals
        current_sum_df = add_dataframe_column(
            current_sum_df,
            [
                "Total Responses to Date",
                "Previous Months Responses",
                "Previous Months Percentage Positive",
                "Previous Months Percentage Negative",
            ],
            [0, 0, 0.1, 0.1],
        )
        logging.info("Summary level table with additional columns added for populating:")
        logging.info(current_sum_df.head())

        # Get cumulative response values for Total, IS1 and NHS from Rolling Monthly Totals and add to current_sum_df
        current_sum_df = copy_value_between_dataframes(
            updated_monthly_rolling_totals,
            current_sum_df,
            "Total responses to date",
            current_month_row,
            "Total Responses to Date",
            2,
        )
        current_sum_df = copy_value_between_dataframes(
            updated_monthly_rolling_totals,
            current_sum_df,
            "Total NHS responses to date",
            current_month_row,
            "Total Responses to Date",
            1,
        )
        current_sum_df = copy_value_between_dataframes(
            updated_monthly_rolling_totals,
            current_sum_df,
            "Total independent responses to date",
            current_month_row,
            "Total Responses to Date",
            0,
        )
        # Get previous months responses for Total, IS1 and NHS from Rolling Monthly Totals and add to current_sum_df
        current_sum_df = copy_value_between_dataframes(
            updated_monthly_rolling_totals,
            current_sum_df,
            "Monthly total responses",
            previous_month_row,
            "Previous Months Responses",
            2,
        )
        current_sum_df = copy_value_between_dataframes(
            updated_monthly_rolling_totals,
            current_sum_df,
            "Monthly NHS responses",
            previous_month_row,
            "Previous Months Responses",
            1,
        )
        current_sum_df = copy_value_between_dataframes(
            updated_monthly_rolling_totals,
            current_sum_df,
            "Monthly independent responses",
            previous_month_row,
            "Previous Months Responses",
            0,
        )
        # Get previous months Percentage Positive values for Total, IS1 and NHS from Rolling Monthly Totals and add to current_sum_df
        current_sum_df = copy_value_between_dataframes(
            updated_monthly_rolling_totals,
            current_sum_df,
            "Monthly total percentage positive",
            previous_month_row,
            "Previous Months Percentage Positive",
            2,
        )
        current_sum_df = copy_value_between_dataframes(
            updated_monthly_rolling_totals,
            current_sum_df,
            "Monthly NHS percentage positive",
            previous_month_row,
            "Previous Months Percentage Positive",
            1,
        )
        current_sum_df = copy_value_between_dataframes(
            updated_monthly_rolling_totals,
            current_sum_df,
            "Monthly independent percentage positive",
            previous_month_row,
            "Previous Months Percentage Positive",
            0,
        )
        # Get previous months Percentage Negative values for Total, IS1 and NHS from Rolling Monthly Totals and add to current_sum_df
        current_sum_df = copy_value_between_dataframes(
            updated_monthly_rolling_totals,
            current_sum_df,
            "Monthly total percentage negative",
            previous_month_row,
            "Previous Months Percentage Negative",
            2,
        )
        current_sum_df = copy_value_between_dataframes(
            updated_monthly_rolling_totals,
            current_sum_df,
            "Monthly NHS percentage negative",
            previous_month_row,
            "Previous Months Percentage Negative",
            1,
        )
        current_sum_df = copy_value_between_dataframes(
            updated_monthly_rolling_totals,
            current_sum_df,
            "Monthly independent percentage negative",
            previous_month_row,
            "Previous Months Percentage Negative",
            0,
        )
        logging.info("Summary level table with additional columns populated:")
        logging.info(current_sum_df.head())

        # Remove the Period column as not required by stakeholders in final output
        current_sum_df = remove_columns(current_sum_df, ["Period"])
        logging.info("Summary level table with period column removed:")
        logging.info(current_sum_df.head())

        # Specify new column order based on stakeholder requirements for Summary level table in output
        new_column_order = [
            "Submitter Type",
            "Number of organisations submitting",
            "Total Responses to Date",
            "Total Responses",
            "Previous Months Responses",
            "Percentage Positive",
            "Previous Months Percentage Positive",
            "Percentage Negative",
            "Previous Months Percentage Negative",
        ]
        # Reorder columns to ensure dataframe is output with columns in order required by stakeholders
        current_sum_df = reorder_columns(current_sum_df, new_column_order)
        logging.info(
            "Summary table with column order changed to conform with expected output:"
        )
        logging.info(current_sum_df.head())

        # Create column names prefixed with either current or previous period to rename current_sum_df
        # columns to meet stakeholder requirements for final output.
        # Use dynamic source rows identified above.
        current_number_orgs_submitting = new_column_name_with_period_prefix(
            current_fft_period, "Number of organisations submitting"
        )
        current_total_responses = new_column_name_with_period_prefix(
            current_fft_period, "Responses"
        )
        previous_total_responses = new_column_name_with_period_prefix(
            previous_fft_period, "Responses"
        )
        current_percent_pos = new_column_name_with_period_prefix(
            current_fft_period, "Percentage Positive"
        )
        previous_percent_pos = new_column_name_with_period_prefix(
            previous_fft_period, "Percentage Positive"
        )
        current_percent_neg = new_column_name_with_period_prefix(
            current_fft_period, "Percentage Negative"
        )
        previous_percent_neg = new_column_name_with_period_prefix(
            previous_fft_period, "Percentage Negative"
        )

        # Rename columns to align with final product requirement
        columns_to_rename = {
            "Submitter Type": "FFT",
            "Number of organisations submitting": current_number_orgs_submitting,
            "Total Responses": current_total_responses,
            "Previous Months Responses": previous_total_responses,
            "Percentage Positive": current_percent_pos,
            "Previous Months Percentage Positive": previous_percent_pos,
            "Percentage Negative": current_percent_neg,
            "Previous Months Percentage Negative": previous_percent_neg,
        }
        current_sum_df = rename_columns(current_sum_df, columns_to_rename)
        logging.info("Current months Summary Inpatient DataFrame with columns renamed:")
        logging.info(current_sum_df.head())

        # Sort dataframe rows by FFT column so Total is top, NHS is middle and IS1 is bottom (descending order)
        current_sum_df = sort_dataframe(current_sum_df, "FFT", False)
        logging.info(
            "Current months Summary Inpatient DataFrame with columns sorted in descending order by the FFT column:"
        )
        logging.info(current_sum_df.head())

        # Processes carried out to generate ICB level data through aggregation of Organisation level DataFrame. Processes/Transformations include:
        # - create copy of organisation level DataFrame without Trust Code/Name and Percentage fields, and group by ICB (Code/Name) fields to
        # aggregate all Likert and Total Response fields
        # - add in and recalculate percentage fields following aggregation then reorder the DataFrame by Total Responses (lowest to highest)
        # - create a new field and implement first level suppression, to ensure no matter how infeasible, it can be applied at ICB level.
        # - create a new field and implement ICB second level suppression so where an ICB requires first level suppression a second ICB
        # (the one with next lowest level of Total Responses) will be suppressed as well
        # - create a new field confirming row level suppression, convert all fields from numeric to object type that would require suppression
        # being applied, and then apply suppression according to suppression rules – all Likert responses are suppressed with ‘*’ for any row
        # requiring suppression, but where it requires first level suppression, the Percentage fields are suppressed as well
        # - sort DataFrame by ICB Code (ascending) and then split the DataFrame (NHS and Independent Providers) and put all Independent Provider
        # rows below NHS rows.

        # Create ICB level DataFrame through aggregation of necessary fields from the Organisation level DataFrame
        # This can be used in the final output and to identify need for suppression to cascade down levels including to Trust level
        # Remove Period and Response rate columns from source DataFrame as not required for data processing or final product
        current_icb_df = remove_columns(
            current_org_df,
            ["Trust Code", "Trust Name", "Percentage Positive", "Percentage Negative"],
        )
        logging.info(
            "Current months ICB Level Inpatient DataFrame with unwanted columns removed:"
        )
        logging.info(current_icb_df.head())

        # Group by the specified columns and sum the numeric columns for ICB-level aggregation
        current_icb_df = sum_grouped_response_fields(
            current_icb_df, ["ICB Code", "ICB Name"]
        )

        # Explicitly remove Title column to prevent concatenation issues in output
        if "Title" in current_icb_df.columns:
            current_icb_df = remove_columns(current_icb_df, ["Title"])
            logging.info("Removed 'Title' column to prevent concatenation issues in output")

        logging.info("Aggregated current months ICB Level Inpatient DataFrame:")
        logging.info(current_icb_df.head())

        # Calculate "Percentage Positive" for the ICB-level DataFrame
        current_icb_df = create_percentage_field(
            current_icb_df, "Percentage Positive", "Very Good", "Good", "Total Responses"
        )
        logging.info(
            "Current months ICB Level Inpatient DataFrame with Percentage Positive field added:"
        )
        logging.info(current_icb_df.head())

        # Calculate "Percentage Positive" for the ICB-level DataFrame
        current_icb_df = create_percentage_field(
            current_icb_df, "Percentage Negative", "Very Poor", "Poor", "Total Responses"
        )
        logging.info(
            "Current months ICB Level Inpatient DataFrame with Percentage Negative field added:"
        )
        logging.info(current_icb_df.head())

        # Reorder current ICB level DataFrame to ensure Integrated Care Board (ICB) are ordered by total responses (lowest to highest)
        current_icb_df = sort_dataframe(current_icb_df, ["Total Responses"], True)
        logging.info("Current months ICB Level Inpatient DataFrame after sort:")
        logging.info(current_icb_df.head())

        # Add first level suppression field at ICB level. Any ICB reporting a non zero number of responses that is less than 5
        # will be flagged as requiring suppression. This is unlikely at ICB level.
        current_icb_df = create_first_level_suppression(
            current_icb_df, "first_level_suppression", "Total Responses"
        )
        logging.info(
            "Current months ICB Level Inpatient DataFrame with first_level_suppression column:"
        )
        logging.info(current_icb_df.head())

        # Add second level suppression field at ICB level. Where an ICB requires first level suppression a second ICB will also be suppressed.
        current_icb_df = create_icb_second_level_suppression(
            current_icb_df, "first_level_suppression", "second_level_suppression"
        )
        logging.info(
            "Current months ICB Level Inpatient DataFrame with second_level_suppression column:"
        )
        logging.info(current_icb_df.head())

        # Add field to show all ICB level rows that require suppression as a result of either first or second level suppression
        current_icb_df = confirm_row_level_suppression(
            current_icb_df,
            "icb_level_suppression_required",
            "first_level_suppression",
            "second_level_suppression",
        )
        logging.info(
            "Current months ICB Level Inpatient DataFrame with icb_level_suppression_required column:"
        )
        logging.info(current_icb_df.head())

        # Convert all fields that can contain suppression values from numeric (integer) to string to avoid a data type mismatch error when suppression is applied
        # List fields requiring conversion impacted by suppression process
        fields_for_converting = [
            "Very Good",
            "Good",
            "Neither Good nor Poor",
            "Poor",
            "Very Poor",
            "Dont Know",
            "Percentage Positive",
            "Percentage Negative",
        ]
        current_icb_df = convert_fields_to_object_type(
            current_icb_df, fields_for_converting
        )
        logging.info(
            "Current months ICB Level Inpatient DataFrame with columns impacted by suppression converted to string:"
        )
        logging.info(current_icb_df.head())

        # Suppress response breakdown fields and percentage fields according to overall and first level suppression conditions in ICB level DataFrame
        # Generate Organisation (Trust) level output using partially cleaned Organisation level DataFrame (current_org_df) and output from ICB transformation
        # Reorder current Organisation file to ensure all Organisations within the same Integrated Care Board (ICB)
        # are grouped together and ordered by total responses
        current_org_df = sort_dataframe(
            current_org_df, ["ICB Code", "Total Responses"], [True, True]
        )
        logging.info("Current months Organisation Level Inpatient DataFrame after sort:")
        logging.info(current_org_df.head())

        # adjust format of positive and negative percentage fields to ensure they are in the correct unit when loaded to output Excel file.
        current_org_df = adjust_percentage_field(current_org_df, "Percentage Positive")
        current_org_df = adjust_percentage_field(current_org_df, "Percentage Negative")

        # Add ranking field to show order of lowest to highest number of responses by Organisation within each ICB.
        # Any Organisation reporting no responses is rated 0, with lowest non zero value ranked 1,
        # with highest Organisation being ranked highest within the ICB
        current_org_df = rank_organisation_results(
            current_org_df, "ICB Code", "Total Responses", "Rank"
        )
        logging.info(
            "Current months Organisation Level Inpatient DataFrame with rank column:"
        )
        logging.info(current_org_df.head())

        # Add first level suppression field at Organisation level. Any Organisation reporting a non zero number of responses that is less than 5
        # will be flagged as requiring suppression. This is very uncommon at Organisation level unless for independent providers.
        current_org_df = create_first_level_suppression(
            current_org_df, "first_level_suppression", "Total Responses"
        )
        logging.info(
            "Current months Organisation Level Inpatient DataFrame with first_level_suppression column:"
        )
        logging.info(current_org_df.head())

        # Add second level suppression field at Organisation level. Where an Organisation within an ICB is the first to be suppressed,
        # the second organisation will also be suppressed.
        current_org_df = create_second_level_suppression(
            current_org_df, "first_level_suppression", "Rank", "second_level_suppression"
        )
        logging.info(
            "Current months Organisation Level Inpatient DataFrame with second_level_suppression column:"
        )
        logging.info(current_org_df.head())

        # Add field showing whether suppression is required at current level as a result of suppression at upper level
        current_org_df = add_suppression_required_from_upper_level_column(
            current_icb_df,
            current_org_df,
            "upper_level_suppression_required",
            "ICB Code",
            "icb_level_suppression_required",
        )
        logging.info(
            "Current months Organisation Level Inpatient DataFrame with upper_level_suppression_required column:"
        )
        logging.info(current_org_df.head())

        # Add field to show all Organisation level rows that require suppression as a result of either first, second or upper level suppression
        current_org_df = confirm_row_level_suppression(
            current_org_df,
            "trust_level_suppression_required",
            "first_level_suppression",
            "second_level_suppression",
            "upper_level_suppression_required",
        )
        logging.info(
            "Current months Organisation Level Inpatient DataFrame with trudt_level_suppression_required column:"
        )
        logging.info(current_org_df.head())

        # Convert all fields that can contain suppression values from numeric (integer) to string to avoid a data type mismatch error when suppression is applied
        current_org_df = convert_fields_to_object_type(
            current_org_df, fields_for_converting
        )
        logging.info(
            "Current months Organisation Level Inpatient DataFrame with columns impacted by suppression converted to string:"
        )
        logging.info(current_org_df.head())

        # Suppress response breakdown fields and percentage fields according to overall and first level suppression conditions in Organisation level DataFrame
        current_org_df = suppress_data(
            current_org_df, "trust_level_suppression_required", "first_level_suppression"
        )
        logging.info(
            "Current months Organisation Level Inpatient DataFrame with suppression applied:"
        )
        logging.info(current_org_df.head())

        # Reorder current Organisation level DataFrame by ICB Code and Total Responses
        current_org_df = sort_dataframe(
            current_org_df, ["ICB Code", "Total Responses"], [True, False]
        )
        logging.info("Current months Organisation Level Inpatient DataFrame after sort:")
        logging.info(current_org_df.head())

        # Move responses for Independent Service Providers to bottom of DataFrame aligned with Stakeholder requirements for output
        current_org_df = move_independent_provider_rows_to_bottom(current_org_df)
        logging.info(
            "Current months Organisation Level Inpatient DataFrame with IS1 orgs moved to bottom:"
        )
        logging.info(current_org_df.tail())

        # Retrieve Trust collection modes and join to current_org_df to meet Output requirement
        # Import Collection Mode data from FFT Inpatient extract for current month
        current_mode_df = load_excel_sheet(
            current_month_file, "Parent_Self_Trusts_Collecti"
        )

        # Standardise column names
        current_mode_df = standardise_fft_column_names(current_mode_df)

        # Remove Title column immediately after loading to prevent "SIGNED-OFF TO DH" values appearing in the output
        if "Title" in current_mode_df.columns:
            current_mode_df = remove_columns(current_mode_df, ["Title"])
            logging.info("Removed 'Title' column from Collection Mode data at source")

        logging.info(
            "Most recent months Collection Mode Inpatient DataFrame after import:"
        )
        logging.info(current_mode_df.head())

        # Take copy of mode df for generating summed Totals including/excluding independent providers, and remove unnecessary fields
        current_mode_totals_df = remove_columns(
            current_mode_df,
            ["Period", "Yearnumber", "Periodname", "Org code", "Org name", "STP Name"],
        )
        logging.info(
            "Current months Collection Mode Inpatient DataFrame with unwanted columns removed to generate Totals:"
        )
        logging.info(current_mode_totals_df.head())

        # Remove Period related columns from source DataFrame as not required for data processing or final product
        current_mode_df = remove_columns(
            current_mode_df,
            ["Period", "Yearnumber", "Periodname", "Org name", "STP Code", "STP Name"],
        )
        logging.info(
            "Current months Collection Mode Inpatient DataFrame with unwanted columns removed:"
        )
        logging.info(current_mode_df.head())

        # Rename columns to align with final product requirement
        mode_columns_to_rename = {
            "Org code": "Trust Code",
        }
        current_mode_df = rename_columns(current_mode_df, mode_columns_to_rename)
        logging.info(
            "Current months Collection Mode Inpatient DataFrame with columns renamed:"
        )
        logging.info(current_mode_df.head())

        current_org_df = join_dataframes(
            current_org_df, current_mode_df, "Trust Code", "left", "one_to_one"
        )

        # Explicitly remove Title column if it reappears after join
        if "Title" in current_org_df.columns:
            current_org_df = remove_columns(current_org_df, ["Title"])
            logging.info("Removed 'Title' column after join to prevent concatenation issues in output")

        logging.info(
            "Current months Trust Level Inpatient DataFrame with Collection Mode columns joined:"
        )
        logging.info(current_mode_df.head())

        # Generate Mode Totals to populate totals fields of the Trust tab
        # Standardise values in STP Code replacing all ICB Codes with NHS and retaining IS1 for independent providers
        current_mode_totals_df = replace_non_matching_values(
            current_mode_totals_df, "STP Code", "IS1", "NHS"
        )
        logging.info(
            "Current months Total mode Inpatient DataFrame with NHS STP(ICB) Codes replaced with 'NHS':"
        )
        logging.info(current_mode_totals_df.head())

        # Rename columns for understanding while processing
        total_mode_columns_to_rename = {
            "STP Code": "Submitter Type",
        }
        current_mode_totals_df = rename_columns(
            current_mode_totals_df, total_mode_columns_to_rename
        )
        logging.info(
            "Current months Total mode Inpatient DataFrame with STP(ICB) Code column renamed:"
        )
        logging.info(current_mode_totals_df.head())

        # Group by specified columns and sum the numeric columns for NHS/IS1 aggregation
        current_mode_totals_df = sum_grouped_response_fields(
            current_mode_totals_df, ["Submitter Type"]
        )

        # Explicitly remove Title column to prevent concatenation issues in output
        if "Title" in current_mode_totals_df.columns:
            current_mode_totals_df = remove_columns(current_mode_totals_df, ["Title"])
            logging.info("Removed 'Title' column to prevent concatenation issues in output")

        logging.info(
            "NHS/IS1 aggregated current months Total mode values Inpatient DataFrame:"
        )
        logging.info(current_mode_totals_df.head())

        # Specify the columns in the current_mode_totals_df to sum for the below function to generate Totals
        sum_mode_cols_for_totals_agg = [
            "Mode SMS",
            "Mode Electronic Discharge",
            "Mode Electronic Home",
            "Mode Paper Discharge",
            "Mode Paper Home",
            "Mode Telephone",
            "Mode Online",
            "Mode Other",
        ]
        # Create a dataframe containing sum total values of NHS/IS1 combined for the period
        sum_current_mode_total_df = create_data_totals(
            current_mode_totals_df,
            current_fft_period,
            "Submitter Type",
            sum_mode_cols_for_totals_agg,
        )
        logging.info("Total aggregated current months mode Totals DataFrame:")
        logging.info(sum_current_mode_total_df.head())

        # Remove extra columns ("Period") prior to merging with IS1/NHS Totals
        sum_current_mode_total_df = remove_columns(sum_current_mode_total_df, ["Period"])

        # Append the totals DataFrame to the bottom of the current summary DataFrame containing NHS/IS1 totals
        current_mode_totals_df = append_dataframes(
            sum_current_mode_total_df, current_mode_totals_df
        )
        logging.info(
            "NHS/IS1 aggregated current months mode values appended to mode Totals DataFrame:"
        )
        logging.info(current_mode_totals_df.head())

        # Remove independent provider row to retain only NHS and Total rows
        current_mode_totals_df = remove_rows_by_cell_content(
            current_mode_totals_df, "Submitter Type", "IS1"
        )
        logging.info("Current months mode Totals DataFrame with IS1 row removed:")
        logging.info(current_mode_totals_df.head())

        # Remove unwanted column ("Submitter Type") prior to loading into Macro Excel Output
        current_mode_totals_df = remove_columns(
            current_mode_totals_df, ["Submitter Type"]
        )

        # Convert numeric fields that may contain missing/null values to object type to enable filling with "NA" without raising Error warning
        mode_fields_to_convert_to_object_type = [
            "Mode SMS",
            "Mode Electronic Discharge",
            "Mode Electronic Home",
            "Mode Paper Discharge",
            "Mode Paper Home",
            "Mode Telephone",
            "Mode Online",
            "Mode Other",
        ]
        current_mode_totals_df = convert_fields_to_object_type(
            current_mode_totals_df, mode_fields_to_convert_to_object_type
        )

        # Clean DataFrame columns replacing with "NA"
        current_mode_totals_df = replace_missing_values(current_mode_totals_df, "NA")
        logging.info(
            "Current months mode Totals DataFrame following column clean and removal:"
        )
        logging.info(current_mode_totals_df.head())

        # Generate and process Site Level DataFrame with Upper Level Suppression input from Organisation Level process
        # Import Site level data from FFT Inpatient extract for current month
        current_site_df = load_excel_sheet(
            current_month_file, "Parent_Self_Trusts_Site_Lev"
        )

        # Standardise column names
        current_site_df = standardise_fft_column_names(current_site_df)

        # Remove Title column immediately after loading to prevent "SIGNED-OFF TO DH" values appearing in the output
        if "Title" in current_site_df.columns:
            current_site_df = remove_columns(current_site_df, ["Title"])
            logging.info("Removed 'Title' column from Site Level data at source")

        logging.info("Most recent months Site Level Inpatient DataFrame after import:")
        logging.info(current_site_df.head())

        # Remove Period related columns from source DataFrame as not required for data processing or final product
        current_site_df = remove_columns(
            current_site_df, ["Period", "Yearnumber", "Periodname"]
        )
        logging.info(
            "Current months Site Level Inpatient DataFrame with unwanted columns removed:"
        )
        logging.info(current_site_df.head())

        # Rename columns to align with final product requirement
        site_columns_to_rename = {
            "Org code": "Trust Code",
            "Org name": "Trust Name",
            "STP Code": "ICB Code",
            "STP Name": "ICB Name",
            "Site Name MAX": "Site Name",
            "1 Very Good": "Very Good",
            "2 Good": "Good",
            "3 Neither good nor poor": "Neither Good nor Poor",
            "4 Poor": "Poor",
            "5 Very poor": "Very Poor",
            "6 Dont Know": "Dont Know",
            "Prop Pos": "Percentage Positive",
            "Prop Neg": "Percentage Negative",
        }
        current_site_df = rename_columns(current_site_df, site_columns_to_rename)
        logging.info(
            "Current months Site Level Inpatient DataFrame with columns renamed:"
        )
        logging.info(current_site_df.head())

        # Reorder current Site Level file to ensure all Sites within an Organisations and within the same Integrated Care Board (ICB)
        # are grouped together and ordered by total responses lowest to highest
        current_site_df = sort_dataframe(
            current_site_df,
            ["ICB Code", "Trust Code", "Total Responses"],
            [True, True, True],
        )
        logging.info("Current months Site Level Inpatient DataFrame after sort:")
        logging.info(current_site_df.head())

        # adjust format of positive and negative percentage fields to ensure they are in the correct unit when loaded to output Excel file.
        current_site_df = adjust_percentage_field(current_site_df, "Percentage Positive")
        current_site_df = adjust_percentage_field(current_site_df, "Percentage Negative")

        # Add ranking field to show order of lowest to highest number of responses by Site within each Organisation (Trust).
        # Any Site reporting no responses is rated 0, with lowest non zero value ranked 1,
        # and Site with highest responses being ranked highest within the Organisation (Trust)
        current_site_df = rank_organisation_results(
            current_site_df, "Trust Code", "Total Responses", "Rank"
        )
        logging.info("Current months Site Level Inpatient DataFrame with rank column:")
        logging.info(current_site_df.head())

        # Add first level suppression field at Site level. Any Site reporting a non zero number of responses that is less than 5
        # will be flagged as requiring suppression.
        current_site_df = create_first_level_suppression(
            current_site_df, "first_level_suppression", "Total Responses"
        )
        logging.info(
            "Current months Site Level Inpatient DataFrame with first_level_suppression column:"
        )
        logging.info(current_site_df.head())

        # Add second level suppression field at Site level. Where a site within a Trust is the first to be suppressed,
        # a second site will also need to be suppressed.
        current_site_df = create_second_level_suppression(
            current_site_df, "first_level_suppression", "Rank", "second_level_suppression"
        )
        logging.info(
            "Current months Site Level Inpatient DataFrame with second_level_suppression column:"
        )
        logging.info(current_site_df.head())

        # Add field showing whether suppression is required at current level as a result of suppression at upper level
        current_site_df = add_suppression_required_from_upper_level_column(
            current_org_df,
            current_site_df,
            "upper_level_suppression_required",
            "Trust Code",
            "trust_level_suppression_required",
        )
        logging.info(
            "Current months Site Level Inpatient DataFrame with upper_level_suppression_required column:"
        )
        logging.info(current_site_df.head())

        # Add field to show all Site level rows that require suppression as a result of either first, second or upper level suppression
        current_site_df = confirm_row_level_suppression(
            current_site_df,
            "site_level_suppression_required",
            "first_level_suppression",
            "second_level_suppression",
            "upper_level_suppression_required",
        )
        logging.info(
            "Current months Site Level Inpatient DataFrame with site_level_suppression_required column:"
        )
        logging.info(current_site_df.head())

        # Convert all fields that can contain suppression values from numeric (integer) to string to avoid a data type mismatch error when suppression is applied
        current_site_df = convert_fields_to_object_type(
            current_site_df, fields_for_converting
        )
        logging.info(
            "Current months Site Level Inpatient DataFrame with columns impacted by suppression converted to string:"
        )
        logging.info(current_site_df.head())

        # Suppress response breakdown fields and percentage fields according to overall and first level suppression conditions in site level DataFrame
        current_site_df = suppress_data(
            current_site_df, "site_level_suppression_required", "first_level_suppression"
        )
        logging.info(
            "Current months Site Level Inpatient DataFrame with suppression applied:"
        )
        logging.info(current_site_df.head())

        # Reorder current Site level DataFrame by ICB Code, Turst Code and Total Responses (highest to lowest)
        current_site_df = sort_dataframe(
            current_site_df,
            ["ICB Code", "Trust Code", "Total Responses"],
            [True, True, False],
        )
        logging.info("Current months Site Level Inpatient DataFrame after sort:")
        logging.info(current_site_df.head())

        # Move responses for Independent Service Providers to bottom of DataFrame aligned with Stakeholder requirements for output
        current_site_df = move_independent_provider_rows_to_bottom(current_site_df)
        logging.info(
            "Current months Site Level Inpatient DataFrame with IS1 orgs moved to bottom:"
        )
        logging.info(current_site_df.tail())

        # Generate and process current Ward level DataFrame with Upper Level Suppression input from Site Level process
        # Import Ward level data from FFT Inpatient extract for current month
        current_ward_df = load_excel_sheet(
            current_month_file, "Parent_Self_Trusts_Ward_Lev"
        )

        # Standardise column names
        current_ward_df = standardise_fft_column_names(current_ward_df)

        # Remove Title column immediately after loading to prevent "SIGNED-OFF TO DH" values appearing in the output
        if "Title" in current_ward_df.columns:
            current_ward_df = remove_columns(current_ward_df, ["Title"])
            logging.info("Removed 'Title' column from Ward Level data at source")

        logging.info("Most recent months Ward Level Inpatient DataFrame after import:")
        logging.info(current_ward_df.head())

        # Remove Period related columns from source DataFrame as not required for data processing or final product
        current_ward_df = remove_columns(
            current_ward_df, ["Period", "Yearnumber", "Periodname"]
        )
        logging.info(
            "Current months Ward Level Inpatient DataFrame with unwanted columns removed:"
        )
        logging.info(current_ward_df.head())

        # Rename columns to align with final product requirement
        ward_columns_to_rename = {
            #        "FFT_Period": "Period",
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
            "Spec 1": "First Speciality",
            "Spec 2": "Second Speciality",
            "Prop Pos": "Percentage Positive",
            "Prop Neg": "Percentage Negative",
        }
        current_ward_df = rename_columns(current_ward_df, ward_columns_to_rename)
        logging.info(
            "Current months Ward Level Inpatient DataFrame with columns renamed:"
        )
        logging.info(current_ward_df.head())

        # Reorder current Ward Level file to ensure all Wards within a Site, within an Organisations and within the same Integrated Care Board (ICB)
        # are grouped together and ordered by total responses (lowest to highest)
        current_ward_df = sort_dataframe(
            current_ward_df,
            ["ICB Code", "Trust Code", "Site Code", "Total Responses"],
            [True, True, True, True],
        )
        logging.info("Current months Ward Level Inpatient DataFrame after sort:")
        logging.info(current_ward_df.head())

        # Add ranking field to show order of lowest to highest number of responses by Ward within each Organisation (Trust) Site.
        # Any Ward reporting no responses is rated 0, with lowest non zero value ranked 1,
        # and the Ward with highest responses being ranked highest within the Organisation (Trust) Site
        current_ward_df = rank_organisation_results(
            current_ward_df, "Trust Code", "Total Responses", "Rank"
        )
        logging.info("Current months Ward Level Inpatient DataFrame with rank column:")
        logging.info(current_ward_df.head())

        # Add first level suppression field at Ward level. Any Ward reporting a non zero number of responses that is less than 5
        # will be flagged as requiring suppression.
        current_ward_df = create_first_level_suppression(
            current_ward_df, "first_level_suppression", "Total Responses"
        )
        logging.info(
            "Current months Ward Level Inpatient DataFrame with first_level_suppression column:"
        )
        logging.info(current_ward_df.head())

        # Add second level suppression field at Ward level. Where a Ward within a Site is the first to be suppressed,
        # a second Ward will also need to be suppressed.
        current_ward_df = create_second_level_suppression(
            current_ward_df, "first_level_suppression", "Rank", "second_level_suppression"
        )
        logging.info(
            "Current months Ward Level Inpatient DataFrame with second_level_suppression column:"
        )
        logging.info(current_ward_df.head())

        # Add field showing whether suppression is required at current level as a result of suppression at upper level
        current_ward_df = add_suppression_required_from_upper_level_column(
            current_site_df,
            current_ward_df,
            "upper_level_suppression_required",
            "Site Code",
            "site_level_suppression_required",
        )
        logging.info(
            "Current months Ward Level Inpatient DataFrame with upper_level_suppression_required column:"
        )
        logging.info(current_ward_df.head())

        # Add field to show all Ward level rows that require suppression as a result of either first, second or upper level suppression
        current_ward_df = confirm_row_level_suppression(
            current_ward_df,
            "ward_level_suppression_required",
            "first_level_suppression",
            "second_level_suppression",
            "upper_level_suppression_required",
        )
        logging.info(
            "Current months Ward Level Inpatient DataFrame with ward_level_suppression_required column:"
        )
        logging.info(current_ward_df.head())

        # Convert all fields that can contain suppression values from numeric (integer) to string to avoid a data type mismatch error when suppression is applied
        current_ward_df = convert_fields_to_object_type(
            current_ward_df, fields_for_converting
        )
        logging.info(
            "Current months Ward Level Inpatient DataFrame with columns impacted by suppression converted to string:"
        )
        logging.info(current_ward_df.head())

        # Suppress response breakdown fields and percentage fields according to overall and first level suppression conditions in Ward level DataFrame
        current_ward_df = suppress_data(
            current_ward_df, "ward_level_suppression_required", "first_level_suppression"
        )
        logging.info(
            "Current months Ward Level Inpatient DataFrame with suppression applied:"
        )
        logging.info(current_ward_df.head())

        # Reorder current Ward level DataFrame by ICB Code, Trust Code, Site Code and Total Responses (Highest to Lowest)
        current_ward_df = sort_dataframe(
            current_ward_df,
            ["ICB Code", "Trust Code", "Site Code", "Total Responses"],
            [True, True, True, False],
        )
        logging.info("Current months Ward Level Inpatient DataFrame after sort:")
        logging.info(current_ward_df.head())

        # Move responses for Independent Service Providers to bottom of DataFrame aligned with Stakeholder requirements for output
        current_ward_df = move_independent_provider_rows_to_bottom(current_ward_df)
        logging.info(
            "Current months Ward Level Inpatient DataFrame with IS1 orgs moved to bottom:"
        )
        logging.info(current_ward_df.tail())

        # With Ward Name as a manual input field, replace special characters to avoid processing errors
        # List of characters to be replaced
        target_chars = [
            ">",
            "<",
            "//",
            "+",
            "*",
            "!",
            "£",
            "$",
            '"',
            "%",
            "^",
            "=",
            "#",
            "@",
        ]
        current_ward_df = replace_character_in_columns(
            current_ward_df, "Ward Name", target_chars, "-"
        )
        logging.info(
            "Current months Ward Level Inpatient DataFrame with special characters in Ward name replaced:"
        )
        logging.info(current_ward_df.head())

        # Processes carried out to finalise transformations and Load the transformed DataFrames into the Macro-enabled Excel template ready for publication.
        # Processes/Transformations include:
        # - Generate Macro Excel back sheet Dropdown Lists to ensure all tab filters in the output contain correct ICB/Trust/Site/Ward name details using filters
        # from ward level to create deduplicated lists for paired ICB Code/Name, paired Site/Ward Code/Name, Trust Code, Trust Name, Site Code, Site Name and Ward Name. Use ICB level output to create unpaired ICB Code and ICB Name lists.
        # - For all DataFrames (ICB, Trust/Organisation, Site, Ward) to ensure compliance with stakeholder requirements of final outputs the following are completed:
        # remove unwanted columns, reorder retained columns, convert numeric fields containing missing/null values to objects and fill with ‘NA’,
        # - Open the Macro-Enabled Excel template to workbook object and load in all DataFrames to the correct workbook sheets using a list of tuples
        # stating which DataFrame to paste in which sheet starting in which row and column. DataFrames include ICB, Trust/Organisation, Site, Ward, England level
        # Likert/Responses including/excluding Independent Providers, Collection Mode totals, all back sheet lists.
        # - update Period subheadings on the Summary sheet with correct current/previous period labels and formatting.
        # - create a percentage style and convert all percentage columns to correct format with 0 decimal places
        # - Update the Period to the current period as part of the ‘Note’ sheet title and Save the updated workbook.

        # Generate Macro Excel Backsheet Dropdown Lists to ensure all tab filters in the output contain correct ICB/Trust/Site/Ward name details
        # Use filtered copy of current_ward_df to create paired ICB Code/Name list with duplicate rows removed
        bs_icb_detail_df = remove_duplicate_rows(
            limit_retained_columns(current_ward_df, ["ICB Code", "ICB Name"])
        )
        logging.info("ICB detail fields for backsheet filtering:")
        logging.info(bs_icb_detail_df.head())

        # Use filtered copy of bs_icb_detail_df to create unpaired sorted ICB Code and ICB Name lists
        bs_icb_code_df = limit_retained_columns(bs_icb_detail_df, ["ICB Code"])
        bs_icb_code_df = sort_dataframe(bs_icb_code_df, "ICB Code", True)
        bs_icb_name_df = limit_retained_columns(bs_icb_detail_df, ["ICB Name"])
        bs_icb_name_df = sort_dataframe(bs_icb_name_df, "ICB Name", True)
        logging.info("ICB code field sorted for backsheet filtering:")
        logging.info(bs_icb_code_df.head())
        logging.info("ICB name field sorted for backsheet filtering:")
        logging.info(bs_icb_name_df.head())

        # Use filtered copy of current_ward_df to create full site_ward Code/Name list with duplicate rows removed
        bs_site_ward_detail_df = remove_duplicate_rows(
            limit_retained_columns(
                current_ward_df,
                [
                    "ICB Code",
                    "Trust Code",
                    "Trust Name",
                    "Site Code",
                    "Site Name",
                    "Ward Name",
                ],
            )
        )
        bs_site_ward_detail_df = sort_dataframe(
            bs_site_ward_detail_df,
            [
                "ICB Code",
                "Trust Code",
                "Trust Name",
                "Site Code",
                "Site Name",
                "Ward Name",
            ],
            [True, True, True, True, True, True],
        )
        bs_site_ward_detail_df = move_independent_provider_rows_to_bottom(
            bs_site_ward_detail_df
        )
        logging.info("site_ward detail fields for backsheet filtering:")
        logging.info(bs_site_ward_detail_df.head(25))
        logging.info(bs_site_ward_detail_df.tail(25))

        # Use filtered copy of current_ward_df to create Trust Code list with duplicate rows removed
        bs_trust_code_df = remove_duplicate_rows(
            limit_retained_columns(current_ward_df, ["Trust Code"])
        )
        bs_trust_code_df = sort_dataframe(bs_trust_code_df, "Trust Code", True)
        logging.info("Trust Code field for backsheet filtering:")
        logging.info(bs_trust_code_df.head())

        # Use filtered copy of current_ward_df to create Trust Name list with duplicate rows removed
        bs_trust_name_df = remove_duplicate_rows(
            limit_retained_columns(current_ward_df, ["Trust Name"])
        )
        bs_trust_name_df = sort_dataframe(bs_trust_name_df, "Trust Name", True)
        logging.info("Trust Name field for backsheet filtering:")
        logging.info(bs_trust_name_df.head())

        # Use filtered copy of current_ward_df to create Site Code list with duplicate rows removed
        bs_site_code_df = remove_duplicate_rows(
            limit_retained_columns(current_ward_df, ["Site Code"])
        )
        bs_site_code_df = sort_dataframe(bs_site_code_df, "Site Code", True)
        logging.info("Site Code field for backsheet filtering:")
        logging.info(bs_site_code_df.head())

        # Use filtered copy of current_ward_df to create Site Name list with duplicate rows removed
        bs_site_name_df = remove_duplicate_rows(
            limit_retained_columns(current_ward_df, ["Site Name"])
        )
        bs_site_name_df = sort_dataframe(bs_site_name_df, "Site Name", True)
        logging.info("Site Name field for backsheet filtering:")
        logging.info(bs_site_name_df.head())

        # Use filtered copy of current_ward_df to create Ward Name list with duplicate rows removed
        bs_ward_name_df = remove_duplicate_rows(
            limit_retained_columns(current_ward_df, ["Ward Name"])
        )
        bs_ward_name_df = sort_dataframe(bs_ward_name_df, "Ward Name", True)
        logging.info("Ward name field for backsheet filtering:")
        logging.info(bs_ward_name_df.head())

        # Remove unwanted columns and reorder retained columns for each DataFrame to ensure it conforms with stakeholder requirements of final outputs.
        # ICB Level - Remove columns not required in final ICB level output including suppression helper columns and period
        current_icb_df = remove_columns(
            current_icb_df,
            [
                "first_level_suppression",
                "second_level_suppression",
                "icb_level_suppression_required",
            ],
        )

        # Specify the correct column order for ICB level output
        icb_output_column_order = [
            "ICB Code",
            "ICB Name",
            "Total Responses",
            "Total Eligible",
            "Percentage Positive",
            "Percentage Negative",
            "Very Good",
            "Good",
            "Neither Good nor Poor",
            "Poor",
            "Very Poor",
            "Dont Know",
        ]
        # Reorder columns of the ICB Level DataFrame according to stakeholder requirements for the final Outputs
        current_icb_df = reorder_columns(current_icb_df, icb_output_column_order)

        # Convert numeric fields that may contain missing/null values to object type to enable filling with "NA" without raising Error warning
        fields_to_convert_to_object_type = [
            "Total Responses",
            "Total Eligible",
            "Very Good",
            "Good",
            "Neither Good nor Poor",
            "Poor",
            "Very Poor",
            "Dont Know",
            "Percentage Positive",
            "Percentage Negative",
        ]
        current_icb_df = convert_fields_to_object_type(
            current_icb_df, fields_to_convert_to_object_type
        )

        # Clean DataFrame columns replacing with "NA"
        current_icb_df = replace_missing_values(current_icb_df, "NA")
        logging.info(
            "Current months ICB Level Inpatient DataFrame following column clean, removal and reorder:"
        )
        logging.info(current_icb_df.head())

        # Org Level - Remove columns not required in final Org level output including suppression helper columns and period
        current_org_df = remove_columns(
            current_org_df,
            [
                "ICB Name",
                "Rank",
                "first_level_suppression",
                "second_level_suppression",
                "upper_level_suppression_required",
                "trust_level_suppression_required",
            ],
        )

        # Specify the correct column order for Org level output
        org_output_column_order = [
            "ICB Code",
            "Trust Code",
            "Trust Name",
            "Total Responses",
            "Total Eligible",
            "Percentage Positive",
            "Percentage Negative",
            "Very Good",
            "Good",
            "Neither Good nor Poor",
            "Poor",
            "Very Poor",
            "Dont Know",
            "Mode SMS",
            "Mode Electronic Discharge",
            "Mode Electronic Home",
            "Mode Paper Discharge",
            "Mode Paper Home",
            "Mode Telephone",
            "Mode Online",
            "Mode Other",
        ]
        # Reorder columns of the Org Level DataFrame according to stakeholder requirements for the final Outputs
        current_org_df = reorder_columns(current_org_df, org_output_column_order)

        # Explicitly remove Title column to prevent concatenation issues in output
        if "Title" in current_org_df.columns:
            current_org_df = remove_columns(current_org_df, ["Title"])
            logging.info("Removed 'Title' column from current_org_df to prevent concatenation issues in output")

        # Convert numeric fields that may contain missing/null values to object type to enable filling with "NA" without raising Error warning
        fields_to_convert_to_object_type = [
            "Total Responses",
            "Total Eligible",
            "Very Good",
            "Good",
            "Neither Good nor Poor",
            "Poor",
            "Very Poor",
            "Dont Know",
            "Percentage Positive",
            "Percentage Negative",
        ]
        current_org_df = convert_fields_to_object_type(
            current_org_df, fields_to_convert_to_object_type
        )

        # Clean DataFrame columns replacing with "NA"
        current_org_df = replace_missing_values(current_org_df, "NA")

        logging.info(
            "Current months Org Level Inpatient DataFrame following column clean, removal and reorder:"
        )
        logging.info(current_org_df.head())

        # Site Level - Remove columns not required in final Site level output including suppression helper columns and period
        current_site_df = remove_columns(
            current_site_df,
            [
                "ICB Name",
                "Rank",
                "first_level_suppression",
                "second_level_suppression",
                "upper_level_suppression_required",
                "site_level_suppression_required",
            ],
        )

        # Specify the correct column order for Site level output
        site_output_column_order = [
            "ICB Code",
            "Trust Code",
            "Trust Name",
            "Site Code",
            "Site Name",
            "Total Responses",
            "Total Eligible",
            "Percentage Positive",
            "Percentage Negative",
            "Very Good",
            "Good",
            "Neither Good nor Poor",
            "Poor",
            "Very Poor",
            "Dont Know",
        ]
        # Reorder columns of the Site Level DataFrame according to stakeholder requirements for the final Outputs
        current_site_df = reorder_columns(current_site_df, site_output_column_order)

        # Convert numeric fields that may contain missing/null values to object type to enable filling with "NA" without raising Error warning
        fields_to_convert_to_object_type = [
            "Total Responses",
            "Total Eligible",
            "Very Good",
            "Good",
            "Neither Good nor Poor",
            "Poor",
            "Very Poor",
            "Dont Know",
            "Percentage Positive",
            "Percentage Negative",
        ]
        current_site_df = convert_fields_to_object_type(
            current_site_df, fields_to_convert_to_object_type
        )

        # Clean DataFrame columns replacing with "NA"
        current_site_df = replace_missing_values(current_site_df, "NA")
        logging.info(
            "Current months Site Level Inpatient DataFrame following column clean, removal and reorder:"
        )
        logging.info(current_site_df.head())

        # Ward Level - Remove columns not required in final Ward level output including suppression helper columns and period
        current_ward_df = remove_columns(
            current_ward_df,
            [
                "ICB Name",
                "Response Rate",
                "Rank",
                "first_level_suppression",
                "second_level_suppression",
                "upper_level_suppression_required",
                "ward_level_suppression_required",
            ],
        )

        # Specify the correct column order for Ward level output
        ward_output_column_order = [
            "ICB Code",
            "Trust Code",
            "Trust Name",
            "Site Code",
            "Site Name",
            "Ward Name",
            "Total Responses",
            "Total Eligible",
            "Percentage Positive",
            "Percentage Negative",
            "Very Good",
            "Good",
            "Neither Good nor Poor",
            "Poor",
            "Very Poor",
            "Dont Know",
            "First Speciality",
            "Second Speciality",
        ]
        # Reorder columns of the Ward Level DataFrame according to stakeholder requirements for the final Outputs
        current_ward_df = reorder_columns(current_ward_df, ward_output_column_order)

        # Convert numeric fields that may contain missing/null values to object type to enable filling with "NA" without raising Error warning
        fields_to_convert_to_object_type = [
            "Total Responses",
            "Total Eligible",
            "Very Good",
            "Good",
            "Neither Good nor Poor",
            "Poor",
            "Very Poor",
            "Dont Know",
            "Percentage Positive",
            "Percentage Negative",
        ]
        current_ward_df = convert_fields_to_object_type(
            current_ward_df, fields_to_convert_to_object_type
        )

        # Clean DataFrame columns replacing with "NA"
        current_ward_df = replace_missing_values(current_ward_df, "NA")
        logging.info(
            "Current months Ward Level Inpatient DataFrame following column clean, removal and reorder:"
        )
        logging.info(current_ward_df.head())

        # Open the Macro-Enabled Excel Source template file to workbook object
        template_file_path = str(
            Path("inputs") / "template_files" / "FFT-inpatient-data-template.xlsm"
        )
        Inpatient_Excel_Workbook = open_macro_excel_file(template_file_path)
        logging.info(
            "Macro-Enabled Excel IP Template opened as workbook object for populating."
        )

        # Write the current DataFrames to the Workbook sheets
        # Specify list of tuples stating which DataFrame to paste in which sheet starting in which row and column)
        dfs_sheets_cells = [
            (current_sum_df, "Summary", 8, 2),
            (current_total_df, "ICB", 12, 3),
            (current_icb_df, "ICB", 15, 1),
            (current_total_df, "Trusts", 12, 4),
            (current_mode_totals_df, "Trusts", 12, 14),
            (current_org_df, "Trusts", 15, 1),
            (current_total_df, "Sites", 12, 6),
            (current_site_df, "Sites", 15, 1),
            (current_total_df, "Wards", 12, 7),
            (current_ward_df, "Wards", 15, 1),
            (bs_icb_detail_df, "BS", 2, 17),
            (bs_icb_code_df, "BS", 2, 19),
            (bs_icb_name_df, "BS", 2, 20),
            (bs_site_ward_detail_df, "BS", 2, 21),
            (bs_trust_code_df, "BS", 2, 31),
            (bs_trust_name_df, "BS", 2, 32),
            (bs_trust_code_df, "BS", 2, 34),
            (bs_trust_name_df, "BS", 2, 35),
            (bs_site_code_df, "BS", 2, 36),
            (bs_site_name_df, "BS", 2, 37),
            (bs_trust_code_df, "BS", 2, 39),
            (bs_trust_name_df, "BS", 2, 40),
            (bs_site_code_df, "BS", 2, 41),
            (bs_site_name_df, "BS", 2, 42),
            (bs_ward_name_df, "BS", 2, 43),
        ]
        write_dataframes_to_sheets(Inpatient_Excel_Workbook, dfs_sheets_cells)
        logging.info(
            "DataFrames written to specified location of specified Workbook sheets."
        )

        # Ensure the Period subheadings on the Summary sheet are dynamically updated with correct current/previous period labels.
        update_cell_with_formatting(
            Inpatient_Excel_Workbook,
            sheet_name="Summary",
            start_row=7,
            start_col=3,
            data=current_fft_period,
            font_size=10,
            bg_color="BFBFBF",
            bold=True,
            font_name="Verdana",
            align_horizontal="center",
            align_vertical="center",
        )
        update_cell_with_formatting(
            Inpatient_Excel_Workbook,
            sheet_name="Summary",
            start_row=7,
            start_col=4,
            data=current_fft_period,
            font_size=10,
            bg_color="BFBFBF",
            bold=True,
            font_name="Verdana",
            align_horizontal="center",
            align_vertical="center",
        )
        update_cell_with_formatting(
            Inpatient_Excel_Workbook,
            sheet_name="Summary",
            start_row=7,
            start_col=5,
            data=current_fft_period,
            font_size=10,
            bg_color="BFBFBF",
            bold=True,
            font_name="Verdana",
            align_horizontal="center",
            align_vertical="center",
        )
        update_cell_with_formatting(
            Inpatient_Excel_Workbook,
            sheet_name="Summary",
            start_row=7,
            start_col=6,
            data=previous_fft_period,
            font_size=10,
            bg_color="BFBFBF",
            bold=True,
            font_name="Verdana",
            align_horizontal="center",
            align_vertical="center",
        )
        update_cell_with_formatting(
            Inpatient_Excel_Workbook,
            sheet_name="Summary",
            start_row=7,
            start_col=7,
            data=current_fft_period,
            font_size=10,
            bg_color="BFBFBF",
            bold=True,
            font_name="Verdana",
            align_horizontal="center",
            align_vertical="center",
        )
        update_cell_with_formatting(
            Inpatient_Excel_Workbook,
            sheet_name="Summary",
            start_row=7,
            start_col=8,
            data=previous_fft_period,
            font_size=10,
            bg_color="BFBFBF",
            bold=True,
            font_name="Verdana",
            align_horizontal="center",
            align_vertical="center",
        )
        update_cell_with_formatting(
            Inpatient_Excel_Workbook,
            sheet_name="Summary",
            start_row=7,
            start_col=9,
            data=current_fft_period,
            font_size=10,
            bg_color="BFBFBF",
            bold=True,
            font_name="Verdana",
            align_horizontal="center",
            align_vertical="center",
        )
        update_cell_with_formatting(
            Inpatient_Excel_Workbook,
            sheet_name="Summary",
            start_row=7,
            start_col=10,
            data=previous_fft_period,
            font_size=10,
            bg_color="BFBFBF",
            bold=True,
            font_name="Verdana",
            align_horizontal="center",
            align_vertical="center",
        )
        logging.info(
            "Summary sheet subheadings updated to correctly reflect current/previous periods."
        )

        # Create percentage style. Must only be done once ahead of application.
        percentage_style = create_percentage_style(Inpatient_Excel_Workbook)

        # convert all percentage columns to correct format with 0 decimal places
        format_column_as_percentage(
            Inpatient_Excel_Workbook,
            sheet_name="ICB",
            start_row=15,
            start_cols=[5, 6],
            percentage_style=percentage_style,
        )
        format_column_as_percentage(
            Inpatient_Excel_Workbook,
            sheet_name="Trusts",
            start_row=15,
            start_cols=[6, 7],
            percentage_style=percentage_style,
        )
        format_column_as_percentage(
            Inpatient_Excel_Workbook,
            sheet_name="Sites",
            start_row=15,
            start_cols=[8, 9],
            percentage_style=percentage_style,
        )
        format_column_as_percentage(
            Inpatient_Excel_Workbook,
            sheet_name="Wards",
            start_row=15,
            start_cols=[9, 10],
            percentage_style=percentage_style,
        )
        logging.info(
            "ICB, Trust, Site and Ward tabs updated to ensure percentage columns are displayed correctly."
        )

        # Define dynamic "Note" sheet title. MAKE SURE THIS IS THE FINAL ACTION SO FILE SAVES AND OPENS ON "Note" TAB.
        notepage_label = combine_text_and_dataframe_cells(
            "Inpatient Friends and Family Test (FFT) Data -", current_fft_period
        )
        # Ensure the Workbook title on the "Note" page updates dynamically with current FFT period suffix and correct format
        update_cell_with_formatting(
            Inpatient_Excel_Workbook,
            sheet_name="Notes",
            start_row=2,
            start_col=1,
            data=notepage_label,
            font_size=18,
            bg_color="FFFFFF",
            font_name="Aptos Narrow",
            align_horizontal="left",
            align_vertical="center",
        )
        logging.info("Note sheet heading updated with current month as suffix.")

        # save dataframes using template of Macro-Enabled Excel file with openpyxl
        output_file_path = str(Path("outputs"))
        prefix_name = "FFT-inpatient-data"
        save_macro_excel_file(
            Inpatient_Excel_Workbook,
            template_file_path,
            output_file_path,
            new=True,
            prefix=prefix_name,
            fft_period_suffix=current_fft_period,
        )

        logging.info(
            "Current months Inpatient Excel output generated from processed DataFrames"
        )

    except Exception as e:
        logging.error("An error occurred in main():", exc_info=e)
        print(f"Error: {e}")
        return False

    logging.info("============================================================")
    logging.info("FFT Inpatient Pipeline completed successfully!")
    logging.info(f"Output file generated: {output_file_path}/FFT-inpatient-data-{current_fft_period}.xlsm")
    logging.info("============================================================")

    print("\n✅ FFT Inpatient Pipeline completed successfully!")
    print(f"Output file generated: {output_file_path}/FFT-inpatient-data-{current_fft_period}.xlsm")

    logging.info("Logging script finished.")
    return True


if __name__ == "__main__":
    success = main()
    exit_code = 0 if success else 1
    exit(exit_code)

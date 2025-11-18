# Required Changes to Match Ground Truth Output

Based on analysis of the current pipeline and the ground truth output, the following changes are needed:

## 1. Filename Format

The filename format needs to be changed from `FFT-inpatient-data-Aug-25.xlsm` to `friends-and-family-test-inpatient-data-august-2025.xlsm`, which requires:

- Changing the prefix from "FFT-inpatient-data" to "friends-and-family-test-inpatient-data"
- Converting the period format from abbreviated month + abbreviated year (e.g., "Aug-25") to lowercase full month name + full year (e.g., "august-2025")

This involves updating the `save_macro_excel_file` function and how `current_fft_period` is generated.

## 2. Date/Period Formatting

- Current: Using string format like "Aug-25" in Summary sheet and elsewhere
- Required: Using datetime objects (2025-08-01 00:00:00) for consistency with ground truth

The `map_fft_period` function needs to be updated to:
1. Return a full month name + full year format for filenames
2. Generate proper datetime objects for display in the Excel file

## 3. Organization Naming Conventions

- Current naming: Using full NHS organization names (e.g., "NHS LANCASHIRE AND SOUTH CUMBRIA INTEGRATED CARE BOARD")
- Required naming: Using shorter format (e.g., "LANCASHIRE AND SOUTH CUMBRIA ICB")

This requires updating the naming conventions in the code or potentially a mapping table to convert between formats.

## 4. Data Processing Logic

The pipeline needs to ensure it processes the same raw data file as used for the ground truth, with identical aggregation logic:

- Ensure all sheets have the same data as the ground truth
- Fix column types to match ground truth (particularly in the Wards sheet)
- Fix data ordering to match ground truth
- Match all summary calculations and totals exactly

## 5. Column Types in Wards Sheet

There are data type inconsistencies in the Wards sheet that need to be addressed, ensuring that columns have consistent types (mostly numeric) as in the ground truth.

## Implementation Approach

1. Update the `map_fft_period` function to use full month names and full years
2. Modify `save_macro_excel_file` to use the new filename format
3. Update column type handling in `convert_fields_to_object_type` function
4. Fix any ordering issues in the data processing pipeline
5. Ensure "Don't Know" column (with apostrophe) is handled correctly across all sheets

# Changelog

The official changelog following semantic versioning conventions, with concise bullet points organised by type (Breaking Changes, New Features, Fixed).

All notable changes to this project will be documented in this file.

Instructions on how to update this Changelog are available in the `Updating the Changelog` section of the [`CONTRIBUTING.md`](./CONTRIBUTING.md).  This project follows [semantic versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased v1.0.0]

### Breaking Changes

- 20251106 - Full pipeline tests and functions built and implemented for Inpatient FFT
- 20251106 - Reorganized directory structure for better organization:
  - Moved data files to data/ directory
  - Created src/fft_pipeline/ for the automated pipeline
  - Created src/etl/ for ETL functions
  - Added utility scripts (process_specific_month.py, fix_rolling_totals_order.py)

### New Features

- 20251105 - Added automated pipeline (src.fft_pipeline) for one-command processing
- 20251105 - Added script to fix Monthly Rolling Totals ordering
- 20251105 - Added scripts for processing specific months
- 20251105 - Changed output filename format to allow side-by-side comparison with ground truth
- 20251031 - Added enhanced error handling throughout the ETL functions

### Fixed

- 20251110 - Added doctests to critical ETL functions and simplified test suite
- 20251110 - Improved test maintainability by combining documentation and testing
- 20251110 - Added test runner documentation to README
- 20251105 - Fixed incorrect previous period selection in Summary output
- 20251105 - Fixed output filename format to match ground truth requirements
- 20251105 - Fixed date/period formatting to use proper datetime objects
- 20251105 - Fixed organization naming conventions
- 20251031 - Fixed `update_monthly_rolling_totals` function to handle missing IS1 submitters
- 20251031 - Fixed `copy_value_between_dataframes` function to check row existence
- 20251031 - Fixed `reorder_columns` function to handle missing columns
- 20251031 - Fixed `adjust_percentage_field` function to handle non-numeric data
- 20251031 - Fixed `sort_dataframe` function to filter non-existent columns
- 20251031 - Fixed `add_suppression_required_from_upper_level_column` function to handle missing columns
- 20251031 - Fixed `replace_character_in_columns` function to check column existence
- 20251031 - Fixed `limit_retained_columns` function to return only available columns
- 20251031 - Fixed `sum_grouped_response_fields` function to prevent string concatenation
- 20251031 - Fixed `write_dataframes_to_sheets` function to handle unexpected columns
- 20251031 - Enhanced data cleaning for 'SIGNED-OFF TO DH' values

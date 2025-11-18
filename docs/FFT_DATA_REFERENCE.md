# Friends and Family Test (FFT) Data Reference

This document provides a detailed reference for understanding the NHS Friends and Family Test data processing pipeline.

## Data Structure Overview

The FFT data is structured hierarchically across multiple organizational levels:

1. **Ward level** - Most granular level capturing patient responses
2. **Site level** - Aggregation of wards within a physical location
3. **Trust/Organization level** - Aggregation of sites within an NHS trust
4. **ICB level** - Integrated Care Board, regional aggregation of trusts

## Response Categories

Patient responses are collected using a standard set of categories:

| Category ID | Description |
|------------|------------|
| 1 | Very Good |
| 2 | Good |
| 3 | Neither Good nor Poor |
| 4 | Poor |
| 5 | Very Poor |
| 6 | Don't Know (sometimes recorded as "Dont Know" without apostrophe) |

## Raw Data Format

Raw FFT data is provided in Excel format with multiple sheets:

- **Parent & Self Trusts - Ward Lev**: Individual ward-level responses
- **Parent & Self Trusts - Site Lev**: Site-level aggregated responses
- **Parent & Self Trusts - Organisa**: Trust/organization-level aggregated responses
- **Parent & Self Trusts - Collecti**: Collection mode data
- **SiteSpec1** and **SiteSpec2**: Specialty data

Key fields in the raw data include:

- `Yearnumber`: Fiscal year (e.g., "2025_26")
- `Periodname`: Month (e.g., "AUGUST")
- `Org code`/`Org name`: Trust/organization identifiers
- `Parent org code`/`Parent name`: ICB identifiers (contain "INTEGRATED CARE BOARD")
- Response fields (e.g., `1 Very Good SUM`)
- `Total Responses`: Sum of all responses
- `Total Eligible`: Total potential responses
- `Response Rate`: Calculated as `Total Responses / Total Eligible`

## Data Processing Pipeline

The FFT pipeline processes raw data through several stages:

1. **Data Loading**:
   - Raw Excel files are loaded from the input directory
   - Files are ordered by modification date to identify current and previous periods

2. **Data Standardization**:
   - Column names are standardized (e.g., "Parent org code" → "STP Code")
   - Organization names are standardized (e.g., removing "NHS" prefix, changing "INTEGRATED CARE BOARD" to "ICB")
   - Data validation is performed on column lengths and numeric types

3. **Aggregation**:
   - Data is aggregated at different levels using the `sum_grouped_response_fields` function
   - Numeric columns are summed, while non-numeric columns use the 'first' value
   - ICB-level data is created by aggregating trust-level data

4. **Metric Calculation**:
   - Percentage metrics are calculated using the `create_percentage_field` function
   - Positive percentage: `(Very Good + Good) / Total Responses`
   - Negative percentage: `(Very Poor + Poor) / Total Responses`
   - Results are stored as decimal values (0 to 1)

5. **Data Suppression**:
   - Multiple levels of suppression are applied for privacy protection
   - First-level suppression: Applied when response count is less than 5
   - Second-level suppression: Applied to prevent deduction of suppressed values
   - Suppressed fields are replaced with "*"

6. **Historical Data Tracking**:
   - Monthly Rolling Totals file maintains historical records
   - New data is appended or updated monthly
   - Cumulative statistics are maintained over time
   - Previous period data is extracted from this file for comparison

## Key Files and Outputs

1. **Input Data**:
   - Raw data files located in `data/inputs/raw_data/inpatient/`
   - Format: `FFT_Inpatients_V1 [Month]-[YY].xlsx`
   - Templates located in `data/inputs/templates/`

2. **Historical Data**:
   - Monthly Rolling Totals file: `data/rolling_totals/Monthly Rolling Totals.xlsx`
   - Contains historical data and cumulative statistics
   - Used for period-to-period comparisons
   - Chronological ordering is critical for correct period comparison

3. **Output Data**:
   - Processed output: `data/outputs/friends-and-family-test-inpatient-data-[month]-[year]_.xlsm`
   - Underscore suffix before extension for side-by-side comparison with ground truth
   - Uses a macro-enabled Excel template
   - Contains multiple sheets for different aggregation levels

4. **Ground Truth Data**:
   - Reference files: `data/outputs/ground_truth/friends-and-family-test-inpatient-data-[month]-[year].xlsm`
   - Used to validate output correctness
   - Compare output with ground truth files manually

## Output Structure

The final output file includes:

1. **Summary**: Top-level statistics with current and previous period comparisons
2. **ICB**: ICB-level aggregated data
3. **Trusts**: Trust/organization-level data with collection modes
4. **Sites**: Site-level data for each trust
5. **Wards**: Individual ward-level data
6. **BS**: Background data for Excel dropdown filters
7. **Notes**: Documentation and metadata

## Important Notes

1. **Previous Period Selection**:
   The pipeline selects the previous period for comparison from the Monthly Rolling Totals file, not the raw data files. This requires proper chronological ordering of entries in the Monthly Rolling Totals file.

2. **Data Consistency**:
   The pipeline assumes consistent data structure across periods. Changes in data format can cause processing errors.

3. **File Naming Conventions**:
   Output files follow a specific naming convention: `friends-and-family-test-inpatient-data-[month]-[year]_.xlsm` with an underscore before the extension to allow side-by-side comparison with ground truth files.

4. **Data Suppression Logic**:
   Suppression is designed to protect patient privacy while maintaining data utility. It works at multiple levels to prevent re-identification.

5. **Independent Sector Providers**:
   Data is separated for NHS and Independent Sector Providers (IS1), with separate aggregations and totals.

6. **Validation Process**:
   Output files should be validated against ground truth using the validation script to ensure correctness.

7. **Automated Pipeline**:
   The `src.fft_pipeline` module automates the entire process, including file discovery, Monthly Rolling Totals updates, and output generation.

---

*This document is intended as a technical reference for understanding the FFT data pipeline and may not reflect official NHS documentation.*
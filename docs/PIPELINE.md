# FFT Automated Pipeline

This document describes the automated FFT (Friends and Family Test) Pipeline that streamlines the process of updating the Monthly Rolling Totals file and generating output files.

## Overview

The FFT Pipeline automates the following steps with a single command:

1. **File Discovery**: Automatically finds FFT inpatient data files in the inputs directory
2. **Monthly Rolling Totals Update**: Updates the Monthly Rolling Totals file with new data chronologically
3. **Correct Previous Period Selection**: Ensures entries are in chronological order for proper period comparison
4. **Output Generation**: Runs the inpatient FFT process to generate the final output file

## Features

- **Fully Automated Workflow**: Process all steps with a single command
- **Smart Processing**: Only processes new files by default
- **Directory Management**: Creates all necessary directories automatically
- **Resilient Error Handling**: Detailed logging and clear error messages
- **Fixes Period Comparison Issue**: Ensures Aug-25 is compared with Jul-25

## Usage

### One-Command Process

To run the full pipeline with a single command:

```bash
# Run with default settings
uv run python -m src.fft_pipeline

# Force processing of all files even if already processed
uv run python -m src.fft_pipeline --force

# Specify a custom directory for raw data
uv run python -m src.fft_pipeline --raw-data-dir /path/to/raw/data
```

### Standard Workflow

1. Copy new raw data files to `data/inputs/raw_data/inpatient/`
2. Run `uv run python -m src.fft_pipeline`
3. Find the output file in `data/outputs/`

### How It Works

1. **File Discovery**:
   - The pipeline scans the `data/inputs/raw_data/inpatient/` directory for FFT inpatient data files
   - It identifies the period (e.g., "Jun-25") from each file name

2. **Processing Decision**:
   - The pipeline checks which periods are already in Monthly Rolling Totals
   - By default, it only processes files for periods not already in Monthly Rolling Totals
   - Use `--force` to process all files regardless

3. **Monthly Rolling Totals Update**:
   - For each file that needs processing, calls `process_specific_month.py`
   - After all files are processed, calls `fix_rolling_totals_order.py` to ensure chronological order

4. **Output Generation**:
   - Runs `inpatient_fft.py` to generate the final output file
   - The output file will use the correct previous period for comparisons
   - Output filename has underscore suffix for side-by-side comparison with ground truth

## Command-Line Options

```
usage: uv run python -m src.fft_pipeline [-h] [--raw-data-dir RAW_DATA_DIR] [--force] [--log-level {DEBUG,INFO,WARNING,ERROR,CRITICAL}]

Run the full FFT pipeline process

options:
  -h, --help            show this help message and exit
  --raw-data-dir RAW_DATA_DIR
                        Directory containing raw data files (default: data/inputs/raw_data/inpatient)
  --force               Force update all files even if they've been processed before (default: False)
  --log-level {DEBUG,INFO,WARNING,ERROR,CRITICAL}
                        Set the logging level (default: INFO)
```

## Validation

After running the pipeline, manually compare the output with the ground truth file:

```bash
# Output location
data/outputs/FFT-inpatient-data-*.xlsm

# Ground truth location
data/outputs/ground_truth/friends-and-family-test-inpatient-data-*.xlsm
```

Open both files to visually inspect structure and data accuracy.

## Logs

The pipeline logs detailed information to:
- Console output
- Log files in `logfiles/full_pipeline/`
- Individual process logs in `logfiles/inpatient_fft/`

## Files and Components

The FFT Pipeline consists of the following components:

1. **src/fft_pipeline/**: Main package
   - **__main__.py**: CLI entry point
   - **pipeline.py**: Main orchestration logic
   - **discovery.py**: File discovery module
   - **rolling_totals.py**: Monthly Rolling Totals handling

2. **Support Scripts**:
   - **process_specific_month.py**: Updates Monthly Rolling Totals with data from a specific file
   - **fix_rolling_totals_order.py**: Ensures chronological ordering of entries
   - **inpatient_fft.py**: Generates the final output file

## Workflow Diagram

```
[Raw Data Files] → [File Discovery] → [Process Each File] → [Reorder Entries] → [Generate Output] → [Validate]
(data/inputs/)      (discovery.py)    (process_month.py)    (fix_rolling       (inpatient_fft.py)  (validate_
                                                             _totals.py)                           fft_output.py)
```

## Troubleshooting

- **No files found**: Ensure FFT files are in `data/inputs/raw_data/inpatient/` and follow the naming pattern "FFT_Inpatients_V1 MMM-YY.xlsx"
- **Processing errors**: Check the logs in `logfiles/full_pipeline/` for detailed error messages
- **Previous period incorrect**: Check the Monthly Rolling Totals order and fix it if needed:
  ```bash
  # Check the current order
  uv run python -c "import pandas as pd; print(pd.read_excel('data/rolling_totals/Monthly Rolling Totals.xlsx', sheet_name='IP')[['FFT Period']])"

  # Fix ordering if needed
  uv run python -m src.scripts.fix_rolling_totals

  # Run pipeline with force flag
  uv run python -m src.fft_pipeline --force
  ```
- **Validation failures**: If the validation shows mismatches between output and ground truth, check the specific differences and review the corresponding processing steps

## Development

To modify the pipeline:

1. Edit files in `src/fft_pipeline/`
2. Test the full pipeline with `uv run python -m src.fft_pipeline`
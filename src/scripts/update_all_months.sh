#!/bin/bash
# Script to process all three months in order to update Monthly Rolling Totals

# Make a backup of the current Monthly Rolling Totals file
echo "Creating backup of Monthly Rolling Totals file..."
cp data/rolling_totals/Monthly\ Rolling\ Totals.xlsx data/rolling_totals/Monthly\ Rolling\ Totals.backup.xlsx

# Process each month in chronological order
echo "Processing June 2025..."
uv run python -m src.scripts.process_month "data/inputs/raw_data/inpatient/FFT_Inpatients_V1 Jun-25.xlsx"

echo "Processing July 2025..."
uv run python -m src.scripts.process_month "data/inputs/raw_data/inpatient/FFT_Inpatients_V1 Jul-25.xlsx"

echo "Processing August 2025..."
uv run python -m src.scripts.process_month "data/inputs/raw_data/inpatient/FFT_Inpatients_V1 Aug-25.xlsx"

echo "All months processed. Monthly Rolling Totals file has been updated."
echo "Now you can run the regular inpatient_fft.py script again to generate the output file."
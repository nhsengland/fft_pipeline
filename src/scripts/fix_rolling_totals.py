#!/usr/bin/env python
import pandas as pd
from pathlib import Path
import logging
from datetime import datetime
import sys

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def fix_rolling_totals_order():
    """
    Fix the order of entries in the Monthly Rolling Totals file to ensure chronological order
    with Jul-25 as the second-to-last entry and Aug-25 as the last entry.
    """
    # Path to the Monthly Rolling Totals file
    monthly_rolling_totals_path = Path("data") / "rolling_totals" / "Monthly Rolling Totals.xlsx"

    # Make a backup of the file
    backup_path = Path("data") / "rolling_totals" / f"Monthly Rolling Totals.backup.{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    logging.info(f"Creating backup at {backup_path}")

    # Copy the file
    import shutil
    shutil.copy2(monthly_rolling_totals_path, backup_path)

    # Load the IP sheet from the Monthly Rolling Totals file
    logging.info("Loading Monthly Rolling Totals file")
    df = pd.read_excel(monthly_rolling_totals_path, sheet_name='IP')

    # Display current order
    logging.info("Current order of entries:")
    for idx, row in df.iterrows():
        logging.info(f"{idx}: {row['FFT Period']}")

    # Create month order mapping to use for sorting
    month_order = {
        'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
        'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12
    }

    # Function to extract month and year from FFT Period for sorting
    def get_sort_key(period):
        if not isinstance(period, str):
            return (0, 0)  # Default value for non-string periods

        parts = period.split('-')
        if len(parts) != 2:
            return (0, 0)  # Default value for invalid format

        month = parts[0][:3]  # Extract first three characters (month abbreviation)
        year = int(parts[1])  # Extract year (last two digits)

        return (year, month_order.get(month, 0))

    # Sort the DataFrame by year and month
    df['sort_key_year'] = df['FFT Period'].apply(lambda x: get_sort_key(x)[0])
    df['sort_key_month'] = df['FFT Period'].apply(lambda x: get_sort_key(x)[1])
    df_sorted = df.sort_values(by=['sort_key_year', 'sort_key_month'])

    # Remove sorting columns
    df_sorted = df_sorted.drop(columns=['sort_key_year', 'sort_key_month'])

    # Reset the index
    df_sorted = df_sorted.reset_index(drop=True)

    # Display new order
    logging.info("New order of entries:")
    for idx, row in df_sorted.iterrows():
        logging.info(f"{idx}: {row['FFT Period']}")

    # Save the sorted DataFrame back to the Excel file
    logging.info("Saving sorted data back to Monthly Rolling Totals file")

    # Use openpyxl engine to preserve other sheets and formatting
    with pd.ExcelWriter(monthly_rolling_totals_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_sorted.to_excel(writer, sheet_name='IP', index=False)

    logging.info("Monthly Rolling Totals file has been updated with correctly ordered entries")
    return True

if __name__ == "__main__":
    success = fix_rolling_totals_order()
    sys.exit(0 if success else 1)
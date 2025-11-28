"""CLI entry point for FFT pipeline."""

import argparse
import logging
import sys

from src.fft.config import PROCESSING_LEVELS, TEMPLATE_CONFIG, SERVICE_TYPES
from src.fft.loaders import find_latest_files, load_raw_data
from src.fft.processors import (
    extract_fft_period,
    standardise_column_names,
    remove_unwanted_columns,
    aggregate_to_icb,
    aggregate_to_trust,
    aggregate_to_site,
    aggregate_to_national,
)
from src.fft.suppression import (
    add_rank_column,
    apply_first_level_suppression,
    apply_second_level_suppression,
    apply_cascade_suppression,
    suppress_values,
)
from src.fft.writers import (
    load_template,
    write_dataframe_to_sheet,
    write_bs_lookup_data,
    write_england_totals,
    update_period_labels,
    format_percentage_columns,
    save_output,
)

# Configure logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


def process_level(df, service_type, level, parent_df=None):
    """Process a single geographic level with suppression.

    Args:
        df: DataFrame for this level
        service_type: 'inpatient', 'ae', or 'ambulance'
        level: 'ward', 'site', 'organisation', or 'icb'
        parent_df: Parent level DataFrame for cascade suppression (optional)

    Returns:
        Processed DataFrame with suppression applied
    """
    logger.info(f"Processing {level} level...")

    # Standardise and clean only for raw data levels (not aggregated ICB)
    if level != "icb":
        df = standardise_column_names(df, service_type, level)
        df = remove_unwanted_columns(df, service_type, level)

    # Determine grouping column for ranking/suppression
    group_by_map = {
        "ward": "Site_Code",
        "site": "Trust_Code",
        "organisation": "ICB_Code",
        "icb": None,
    }
    group_by_col = group_by_map.get(level)

    # Add ranking
    df = add_rank_column(df, group_by_col)

    # Apply first-level suppression
    df = apply_first_level_suppression(df)

    # Apply second-level suppression
    df = apply_second_level_suppression(df, group_by_col)

    # Apply cascade suppression from parent if available
    if parent_df is not None:
        parent_code_col_map = {
            "organisation": "ICB_Code",
            "site": "Trust_Code",
            "ward": "Site_Code",
        }
        parent_code_col = parent_code_col_map.get(level)
        if parent_code_col:
            df = apply_cascade_suppression(
                parent_df, df, parent_code_col, parent_code_col, "Suppression_Required"
            )

    # Confirm overall suppression required
    suppression_cols = [col for col in df.columns if "Suppression" in col]
    df["Suppression_Required"] = df[suppression_cols].max(axis=1)

    # Apply suppression (replace values with *)
    df = suppress_values(df)

    return df


def run_pipeline(service_type: str) -> None:
    """Run the full FFT pipeline for a service type.

    Args:
        service_type: 'inpatient', 'ae', or 'ambulance'
    """
    logger.info(f"Starting FFT pipeline for {service_type}")

    processing_config = PROCESSING_LEVELS[service_type]
    levels = processing_config["levels"]
    sheet_mapping = processing_config["sheet_mapping"]

    # Step 1: Load latest files
    logger.info("Loading latest raw data files...")
    files = find_latest_files(service_type, n=2)
    if not files:
        raise FileNotFoundError(f"No raw data files found for {service_type}")
    logger.info(f"Found {len(files)} files: {[f.name for f in files]}")

    # Step 2: Load raw data (current month)
    logger.info("Loading raw data from Excel...")
    raw_data = load_raw_data(files[0])
    logger.info(f"Loaded {len(raw_data)} sheets: {list(raw_data.keys())}")

    # Step 3: Extract FFT period from first available sheet
    first_sheet = list(raw_data.values())[0]
    fft_period = extract_fft_period(first_sheet)
    logger.info(f"FFT Period: {fft_period}")

    # Step 4: Process each level
    processed_data = {}
    parent_df = None

    for level in levels:
        sheet_name = sheet_mapping.get(level)
        if sheet_name not in raw_data:
            raise KeyError(f"Sheet '{sheet_name}' not found in raw data")

        df = raw_data[sheet_name].copy()
        processed_df = process_level(df, service_type, level, parent_df)
        processed_data[level] = processed_df
        parent_df = processed_df  # Pass to next level for cascade

    # Step 5: Aggregate to ICB level
    logger.info("Aggregating to ICB level...")
    org_df = processed_data.get("organisation", processed_data[levels[-1]])
    icb_df = aggregate_to_icb(org_df)
    icb_df = process_level(icb_df, service_type, "icb")
    processed_data["icb"] = icb_df

    # Step 6: Aggregate to national level
    logger.info("Aggregating to national level...")
    national_df, org_counts = aggregate_to_national(org_df)
    logger.info(f"Organisation counts: {org_counts}")

    # DEBUG: Check what we're generating
    print("\n=== DEBUG: National DF ===")
    print(national_df)
    print("\n=== DEBUG: ICB DF ===")
    print(icb_df.head())
    print(f"ICB columns: {list(icb_df.columns)}")

    # Step 7: Load template
    logger.info("Loading template...")
    wb = load_template(service_type)

    # Step 8: Write data to sheets
    template_config = TEMPLATE_CONFIG[service_type]
    data_start_row = template_config["data_start_row"]

    level_to_sheet = {
        "icb": "ICB",
        "organisation": "Trusts",
        "site": "Sites",
        "ward": "Wards",
    }

    for level, df in processed_data.items():
        sheet_name = level_to_sheet.get(level)
        if sheet_name and sheet_name in wb.sheetnames:
            logger.info(f"Writing {level} data to {sheet_name}...")
            write_dataframe_to_sheet(wb, df, sheet_name, data_start_row)

    # Step 9: Write England totals
    logger.info("Writing England totals...")
    write_england_totals(wb, service_type, national_df, org_counts)

    # Step 10: Write BS lookup data (from ward level or lowest available)
    logger.info("Writing BS lookup data...")
    lowest_level = levels[0]  # ward for inpatient, site for ae, etc.
    write_bs_lookup_data(wb, processed_data[lowest_level], service_type)

    # Step 11: Update period labels
    logger.info("Updating period labels...")
    update_period_labels(wb, service_type, fft_period)

    # Step 12: Format percentage columns
    logger.info("Formatting percentage columns...")
    format_percentage_columns(wb, service_type)

    # Step 13: Save output
    logger.info("Saving output...")
    output_path = save_output(wb, service_type, fft_period)
    logger.info(f"✓ Output saved to: {output_path}")


def main():
    """Main entry point for FFT pipeline."""

    parser = argparse.ArgumentParser(
        description="FFT Pipeline - Process NHS Friends and Family Test data"
    )

    # Build mutually exclusive group from config
    group = parser.add_mutually_exclusive_group(required=True)
    for flag, service_type in SERVICE_TYPES.items():
        group.add_argument(
            f"--{flag}", action="store_true", help=f"Process {service_type.title()} data"
        )

    args = parser.parse_args()

    # Determine service type from args
    service_type = None
    for flag, stype in SERVICE_TYPES.items():
        if getattr(args, flag, False):
            service_type = stype
            break

    try:
        run_pipeline(service_type)
        logger.info("✓ Pipeline completed successfully")
    except Exception as e:
        logger.error(f"Pipeline failed: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()

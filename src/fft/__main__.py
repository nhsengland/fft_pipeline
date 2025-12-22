"""CLI entry point for FFT pipeline."""

import argparse
import logging
import sys
from pathlib import Path

from fft.config import (
    OUTPUT_COLUMNS,
    PROCESSING_LEVELS,
    SERVICE_TYPES,
    TEMPLATE_CONFIG,
)
from fft.loaders import find_latest_files, load_raw_data
from fft.processors import (
    aggregate_to_icb,
    aggregate_to_national,
    clean_icb_name,
    extract_fft_period,
    merge_collection_modes,
    remove_unwanted_columns,
    standardise_column_names,
)
from fft.suppression import (
    add_rank_column,
    apply_cascade_suppression,
    apply_first_level_suppression,
    apply_second_level_suppression,
    suppress_values,
)
from fft.writers import (
    format_percentage_columns,
    load_template,
    populate_summary_sheet,
    save_output,
    update_period_labels,
    write_bs_lookup_data,
    write_dataframe_to_sheet,
    write_england_totals,
)

# Configure logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)



# %%
def process_single_file(
    service_type: str, file_path: Path, processing_config: dict
) -> None:
    """Process a single raw data file and generate output report."""
    levels = processing_config["levels"]
    sheet_mapping = processing_config["sheet_mapping"]

    # Step 2: Load raw data
    logger.info("Loading raw data from Excel...")
    raw_data = load_raw_data(file_path)

    # After loading raw_data (Step 2)
    logger.debug(f"Sheets loaded: {list(raw_data.keys())}")
    for sheet_name, df in raw_data.items():
        logger.debug(f"  {sheet_name}: {len(df)} rows")

    # Step 3: Extract FFT period
    first_sheet = list(raw_data.values())[0]
    fft_period = extract_fft_period(first_sheet)
    logger.info(f"FFT Period: {fft_period}")

    # Step 4: Standardise and clean each level (NO suppression yet)
    cleaned_data = {}
    for level in levels:
        sheet_name = sheet_mapping.get(level)
        if sheet_name not in raw_data:
            raise KeyError(f"Sheet '{sheet_name}' not found in raw data")

        logger.info(f"Cleaning {level} level...")
        df = raw_data[sheet_name].copy()
        df = standardise_column_names(df, service_type, level)
        df = remove_unwanted_columns(df, service_type, level)

        if level == "organisation" and "collection_mode" in sheet_mapping:
            coll_sheet = sheet_mapping["collection_mode"]
            if coll_sheet in raw_data:
                logger.info("Merging collection mode data...")
                coll_df = raw_data[coll_sheet].copy()
                coll_df = coll_df.rename(columns={"Org code": "Trust_Code"})

                df = merge_collection_modes(df, coll_df)
        cleaned_data[level] = df

    # Check if we have any data to process
    if all(df.empty for df in cleaned_data.values()):
        logger.warning(f"No data found in {file_path.name} - skipping")
        return

    # Step 4.5: Mark independent sector providers across all levels
    logger.info("Marking independent sector providers...")

    for level in ["organisation", "site", "ward"]:
        if level not in cleaned_data:
            continue
        df = cleaned_data[level]
        df["ICB_Code"] = df.apply(
            lambda row: "IS1"
            if not (
                "NHS" in str(row["Trust_Name"]).upper()
                and "TRUST" in str(row["Trust_Name"]).upper()
            )
            else row["ICB_Code"],
            axis=1,
        )
        df["ICB_Name"] = df.apply(
            lambda row: "INDEPENDENT SECTOR PROVIDERS"
            if row["ICB_Code"] == "IS1"
            else row["ICB_Name"],
            axis=1,
        )
        cleaned_data[level] = df

    # Step 4.6: Clean ICB names
    logger.info("Cleaning ICB names...")
    for level in cleaned_data:
        if "ICB_Name" in cleaned_data[level].columns:
            cleaned_data[level]["ICB_Name"] = cleaned_data[level]["ICB_Name"].apply(
                clean_icb_name
            )

    # Step 5: Aggregate to ICB level
    logger.info("Aggregating to ICB level...")
    org_df = cleaned_data["organisation"]
    icb_df = aggregate_to_icb(org_df)

    # Step 6: Aggregate to national level
    logger.info("Aggregating to national level...")
    national_df, org_counts = aggregate_to_national(org_df)
    logger.info(f"Organisation counts: {org_counts}")

    # Step 7: Apply suppression to each level (top-down cascade)
    logger.info("Applying suppression...")

    # ICB level suppression (no parent)
    icb_df = add_rank_column(icb_df, group_by_col=None)
    icb_df = apply_first_level_suppression(icb_df)
    icb_df = apply_second_level_suppression(icb_df, group_by_col=None)
    icb_df["Suppression_Required"] = (
        icb_df["First_Level_Suppression"] | icb_df["Second_Level_Suppression"]
    )
    icb_suppressed = suppress_values(icb_df.copy())

    # Organisation level suppression (cascade from ICB)
    org_df = add_rank_column(org_df, group_by_col="ICB_Code")
    org_df = apply_first_level_suppression(org_df)
    org_df = apply_second_level_suppression(org_df, group_by_col="ICB_Code")
    org_df = apply_cascade_suppression(
        icb_df, org_df, "ICB_Code", "ICB_Code", "Suppression_Required"
    )
    org_df["Suppression_Required"] = org_df[
        ["First_Level_Suppression", "Second_Level_Suppression", "Cascade_Suppression"]
    ].max(axis=1)
    org_suppressed = suppress_values(org_df.copy())

    # Site level suppression (cascade from Organisation)
    if "site" in cleaned_data:
        site_df = cleaned_data["site"]
        site_df = add_rank_column(site_df, group_by_col="Trust_Code")
        site_df = apply_first_level_suppression(site_df)
        site_df = apply_second_level_suppression(site_df, group_by_col="Trust_Code")
        site_df = apply_cascade_suppression(
            org_df, site_df, "Trust_Code", "Trust_Code", "Suppression_Required"
        )
        site_df["Suppression_Required"] = site_df[
            ["First_Level_Suppression", "Second_Level_Suppression", "Cascade_Suppression"]
        ].max(axis=1)
        site_suppressed = suppress_values(site_df.copy())

    # Ward level suppression (cascade from Site)
    if "ward" in cleaned_data:
        ward_df = cleaned_data["ward"]
        ward_df = add_rank_column(ward_df, group_by_col="Site_Code")
        ward_df = apply_first_level_suppression(ward_df)
        ward_df = apply_second_level_suppression(ward_df, group_by_col="Site_Code")
        ward_df = apply_cascade_suppression(
            site_df, ward_df, "Site_Code", "Site_Code", "Suppression_Required"
        )
        ward_df["Suppression_Required"] = ward_df[
            ["First_Level_Suppression", "Second_Level_Suppression", "Cascade_Suppression"]
        ].max(axis=1)
        ward_suppressed = suppress_values(ward_df.copy())

    # Step 8: Load template workbook
    logger.info("Loading template...")
    wb = load_template(service_type)

    # Step 8.5: Sort DataFrames (NHS entries alphabetically, IS1 at end)
    logger.info("Sorting data...")

    def sort_with_is1_last(df, sort_cols):
        """Sort DataFrame with IS1 entries appearing last."""
        df = df.copy()
        df["_is_is1"] = df["ICB_Code"] == "IS1"
        df = df.sort_values(["_is_is1"] + sort_cols)
        df = df.drop(columns=["_is_is1"])
        return df

    icb_suppressed = sort_with_is1_last(icb_suppressed, ["ICB_Code"])
    org_suppressed = sort_with_is1_last(org_suppressed, ["ICB_Code", "Trust_Code"])
    if "site" in cleaned_data:
        site_suppressed = sort_with_is1_last(
            site_suppressed, ["ICB_Code", "Trust_Code", "Site_Code"]
        )
    if "ward" in cleaned_data:
        ward_suppressed = sort_with_is1_last(
            ward_suppressed, ["ICB_Code", "Trust_Code", "Site_Code", "Ward_Name"]
        )

    # Step 9: Write data to sheets (use SUPPRESSED versions)
    suppressed_data = {
        "icb": icb_suppressed,
        "organisation": org_suppressed,
    }
    if "site" in cleaned_data:
        suppressed_data["site"] = site_suppressed
    if "ward" in cleaned_data:
        suppressed_data["ward"] = ward_suppressed

    template_config = TEMPLATE_CONFIG[service_type]
    data_start_row = template_config["data_start_row"]

    for level, df in suppressed_data.items():
        sheet_config = template_config["sheets"].get(level)
        if not sheet_config:
            continue

        sheet_name = sheet_config["sheet_name"]
        if sheet_name in wb.sheetnames:
            logger.info(f"Writing {level} data to {sheet_name}...")

            # Filter to only output columns
            output_cols = OUTPUT_COLUMNS[service_type].get(sheet_name, [])
            available_cols = [col for col in output_cols if col in df.columns]
            output_df = df[available_cols]

            write_dataframe_to_sheet(wb, output_df, sheet_name, data_start_row)

    # Step 10: Write England totals
    logger.info("Writing England totals...")
    write_england_totals(wb, service_type, national_df, org_counts)

    # Step 11: Write BS lookup data (use unsuppressed ward data for lookups)
    logger.info("Writing BS lookup data...")
    if "ward" in cleaned_data:
        write_bs_lookup_data(wb, ward_df, service_type)
    elif "site" in cleaned_data:
        write_bs_lookup_data(wb, site_df, service_type)
    else:
        write_bs_lookup_data(wb, org_df, service_type)

    # Step 11.5: Populate Summary sheet
    logger.info("Populating Summary sheet...")
    populate_summary_sheet(wb, service_type, fft_period)

    # Step 12: Update period labels
    logger.info("Updating period labels...")
    update_period_labels(wb, service_type, fft_period)

    # Step 13: Format percentage columns
    logger.info("Formatting percentage columns...")
    format_percentage_columns(wb, service_type)

    # Step 14: Save output
    logger.info("Saving output...")
    output_path = save_output(wb, service_type, fft_period)
    logger.info(f"✓ Output saved to: {output_path}")


# %%
def run_pipeline(service_type: str, month: str = None) -> None:
    """Run the full FFT pipeline for a service type."""
    logger.info(f"Starting FFT pipeline for {service_type}")

    processing_config = PROCESSING_LEVELS[service_type]

    # Step 1: Find all raw data files
    logger.info("Finding raw data files...")
    files = find_latest_files(service_type, n=100)  # Get all available files
    if not files:
        raise FileNotFoundError(f"No raw data files found for {service_type}")
    logger.info(f"Found {len(files)} files to process")

    # Filter to specific month if requested
    if month:
        files = [f for f in files if month in f.name]
        if not files:
            raise FileNotFoundError(f"No file found for month: {month}")

    # Process each file
    for file_path in files:
        logger.info("")
        logger.info("=" * 50)
        logger.info(f"Processing: {file_path.name}")
        logger.info("=" * 50)

        try:
            process_single_file(service_type, file_path, processing_config)
        except Exception as e:
            logger.error(f"Failed to process {file_path.name}: {e}", exc_info=True)
            continue  # Continue with next file

    logger.info("")
    logger.info(f"✓ Pipeline completed - processed {len(files)} files")


# %%
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

    parser.add_argument(
        "--month",
        type=str,
        default=None,
        help="Process specific month only (e.g., Aug-25)",
    )

    args = parser.parse_args()

    # Determine service type from args
    service_type = None
    for flag, stype in SERVICE_TYPES.items():
        if getattr(args, flag, False):
            service_type = stype
            break

    try:
        run_pipeline(service_type, month=args.month)
        logger.info("✓ Pipeline completed successfully")
    except Exception as e:
        logger.error(f"Pipeline failed: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()

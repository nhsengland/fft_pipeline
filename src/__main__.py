"""CLI entry point for FFT pipeline."""

import argparse
import logging
import sys

from src.fft.config import (
    PROCESSING_LEVELS,
    TEMPLATE_CONFIG,
    SERVICE_TYPES,
    OUTPUT_COLUMNS,
)
from src.fft.loaders import find_latest_files, load_raw_data
from src.fft.processors import (
    extract_fft_period,
    standardise_column_names,
    remove_unwanted_columns,
    aggregate_to_icb,
    aggregate_to_trust,
    aggregate_to_site,
    aggregate_to_national,
    merge_collection_modes,
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


# TODO: Can I remove this function?
# %%
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
    # df = suppress_values(df)

    return df


# %%
def run_pipeline(service_type: str) -> None:
    """Run the full FFT pipeline for a service type."""
    logger.info(f"Starting FFT pipeline for {service_type}")

    processing_config = PROCESSING_LEVELS[service_type]
    levels = processing_config["levels"]
    sheet_mapping = processing_config["sheet_mapping"]

    # Step 1: Find latest files
    logger.info("Loading latest raw data files...")
    files = find_latest_files(service_type, n=2)
    if not files:
        raise FileNotFoundError(f"No raw data files found for {service_type}")
    logger.info(f"Found {len(files)} files: {[f.name for f in files]}")

    # Step 2: Load raw data (current month)
    logger.info("Loading raw data from Excel...")
    raw_data = load_raw_data(files[0])
    logger.info(f"Loaded {len(raw_data)} sheets: {list(raw_data.keys())}")

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

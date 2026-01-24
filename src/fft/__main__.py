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
    VALIDATION_CONFIG,
    VALIDATION_KEY_COLUMNS,
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
from fft.validation import (
    compare_data_by_key,
    compare_data_range,
    extract_service_type,
    find_matching_ground_truth,
    print_comparison_report,
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
def process_single_file(  # noqa: PLR0912,PLR0915 # Justified: Sequential ETL pipeline with 14 steps
    service_type: str, file_path: Path, processing_config: dict
) -> Path | None:
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

    # Step 5: Mark independent sector providers across all levels
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

    # Step 6: Clean ICB names
    logger.info("Cleaning ICB names...")
    for level, df in cleaned_data.items():
        if "ICB_Name" in df.columns:
            cleaned_data[level]["ICB_Name"] = df["ICB_Name"].apply(clean_icb_name)

    # Step 7: Aggregate to ICB level
    logger.info("Aggregating to ICB level...")
    org_df = cleaned_data["organisation"]
    icb_df = aggregate_to_icb(org_df)

    # Step 8: Aggregate to national level
    logger.info("Aggregating to national level...")
    national_df, org_counts = aggregate_to_national(org_df)
    logger.info(f"Organisation counts: {org_counts}")

    # Step 9: Apply suppression to each level (top-down cascade)
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

    # Ward level suppression (includes second-level and cascade from Site)
    if "ward" in cleaned_data:
        ward_df = cleaned_data["ward"]
        # Apply ranking with VBA-compliant tie-breaking
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

    # Step 10: Load template workbook
    logger.info("Loading template...")
    wb = load_template(service_type)

    # Step 11: Sort DataFrames (NHS entries alphabetically, IS1 at end)
    logger.info("Sorting data...")

    def sort_with_is1_last(df, sort_cols):
        """Sort DataFrame with IS1 entries appearing last."""
        df = df.copy()
        df["_is_is1"] = df["ICB_Code"] == "IS1"
        df = df.sort_values(["_is_is1"] + sort_cols)
        df = df.drop(columns=["_is_is1"])
        return df

    icb_suppressed = sort_with_is1_last(icb_suppressed, ["ICB_Code"])
    # Apply VBA-aligned sorting: ICB_Code, Trust_Name
    org_suppressed = sort_with_is1_last(org_suppressed, ["ICB_Code", "Trust_Name"])
    if "site" in cleaned_data:
        # Apply VBA-aligned sorting: ICB_Code, Trust_Name, Site_Name
        site_suppressed = sort_with_is1_last(
            site_suppressed, ["ICB_Code", "Trust_Name", "Site_Name"]
        )
    if "ward" in cleaned_data:
        # Apply VBA-aligned sorting: ICB_Code, Trust_Name, Site_Name, Ward_Name
        # VBA sorts by Trust Name (column C), not Trust Code (column B)
        ward_suppressed = sort_with_is1_last(
            ward_suppressed, ["ICB_Code", "Trust_Name", "Site_Name", "Ward_Name"]
        )

    # Step 12: Write data to sheets (use SUPPRESSED versions)
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

    # Step 13: Write England totals
    logger.info("Writing England totals...")
    write_england_totals(wb, service_type, national_df, org_counts)

    # Step 14: Write BS lookup data (use unsuppressed ward data for lookups)
    logger.info("Writing BS lookup data...")
    if "ward" in cleaned_data:
        write_bs_lookup_data(wb, ward_df, service_type)
    elif "site" in cleaned_data:
        write_bs_lookup_data(wb, site_df, service_type)
    else:
        write_bs_lookup_data(wb, org_df, service_type)

    # Step 15: Populate Summary sheet
    logger.info("Populating Summary sheet...")
    populate_summary_sheet(wb, service_type, fft_period)

    # Step 16: Update period labels
    logger.info("Updating period labels...")
    update_period_labels(wb, service_type, fft_period)

    # Step 17: Format percentage columns
    logger.info("Formatting percentage columns...")
    format_percentage_columns(wb, service_type)

    # Step 18: Save output
    logger.info("Saving output...")
    output_path = save_output(wb, service_type, fft_period)
    logger.info(f"✓ Output saved to: {output_path}")

    return output_path


def validate_existing_outputs(month_filter: str | None = None) -> None:
    """Validate existing output files against ground truth."""
    outputs_dir = Path("data/outputs")

    if not outputs_dir.exists():
        logger.error(f"Outputs directory not found: {outputs_dir}")
        sys.exit(1)

    # Find output files
    output_files = list(outputs_dir.glob("*.xl*"))
    output_files = [f for f in output_files if f.is_file()]

    if month_filter:
        output_files = [f for f in output_files if month_filter in f.name]

    if not output_files:
        msg = f"No output files found in {outputs_dir}"
        if month_filter:
            msg += f" for month: {month_filter}"
        logger.error(msg)
        sys.exit(1)

    logger.info(f"Found {len(output_files)} output file(s) to validate")

    validated_count = 0
    for output_path in output_files:
        # Determine service type
        service_type = extract_service_type(output_path.name)

        if not service_type:
            logger.warning(f"Cannot determine service type for: {output_path.name}")
            continue

        logger.info(f"Validating {output_path.name} (service: {service_type})")

        try:
            _validate_output(output_path, service_type)
            validated_count += 1
        except Exception as e:
            logger.error(f"Validation failed for {output_path.name}: {e}")

    if validated_count == 0:
        logger.error("No files could be validated")
        sys.exit(1)

    logger.info(f"✓ Validation completed for {validated_count} file(s)")


def _validate_output(output_path: Path, service_type: str) -> None:
    """Validate generated output against ground truth files.

    Compares pipeline output from data/outputs/ against
    corresponding ground truth from data/outputs/ground_truth/.
    """
    # Find best matching ground truth file using smart pattern matching
    ground_truth_dir = output_path.parent / "ground_truth"
    ground_truth_path = find_matching_ground_truth(output_path, ground_truth_dir)

    if ground_truth_path is None:
        logger.warning(f"No matching ground truth file found in: {ground_truth_dir}")
        logger.warning(
            "Skipping validation - place matching ground truth file in "
            "data/outputs/ground_truth/ (matches by month and service type)"
        )
        return

    logger.info(f"Found matching ground truth: {ground_truth_path.name}")
    logger.info(f"Comparing against: {ground_truth_path}")

    # Compare all sheets, focusing on data areas (skip template control rows)
    max_differences_to_show = 25  # Show enough detail but keep readable

    try:
        results = []
        sheets_to_check = VALIDATION_CONFIG.get(service_type, [])

        for sheet_name in sheets_to_check:
            # Use key-based comparison for sheets with configured key columns
            if sheet_name in VALIDATION_KEY_COLUMNS:
                result = compare_data_by_key(
                    expected_path=ground_truth_path,
                    actual_path=output_path,
                    sheet_name=sheet_name,
                    key_column=VALIDATION_KEY_COLUMNS[sheet_name],
                    start_row=15,
                    data_only=True,
                )
            else:
                # Use row-based comparison for other sheets
                result = compare_data_range(
                    expected_path=ground_truth_path,
                    actual_path=output_path,
                    sheet_name=sheet_name,
                    start_row=15,
                    data_only=True,
                )
            results.append(result)

        # Generate comprehensive validation report
        print("\n" + "=" * 60)
        print(f"VALIDATION REPORT: {output_path.name}")
        print("=" * 60)
        print("Comparing: Pipeline Output vs VBA Ground Truth")
        print_comparison_report(results, max_diffs_per_sheet=max_differences_to_show)

        # Summary assessment
        total_sheets = len(
            [
                r
                for r in results
                if not (r["missing_in_actual"] or r["missing_in_expected"])
            ]
        )
        identical_sheets = len([r for r in results if r["identical"]])

        if total_sheets > 0 and identical_sheets == total_sheets:
            logger.info("✓ Validation PASSED - All sheets identical to ground truth")
        elif total_sheets == 0:
            logger.warning("⚠ Validation INCOMPLETE - No comparable sheets found")
        else:
            msg = f"✗ Validation FAILED - {identical_sheets}/{total_sheets} sheets match"
            logger.warning(msg)

    except Exception as e:
        logger.error(f"Validation failed with error: {e}")
        raise


# %%
def run_pipeline(service_type: str, month: str | None = None) -> None:
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
    """Process FFT pipeline data from command line arguments."""
    parser = argparse.ArgumentParser(
        description="FFT Pipeline - Process NHS Friends and Family Test data"
    )

    # Build mutually exclusive group from config
    group = parser.add_mutually_exclusive_group(required=True)
    for flag, service_type in SERVICE_TYPES.items():
        group.add_argument(
            f"--{flag}", action="store_true", help=f"Process {service_type.title()} data"
        )
    group.add_argument(
        "--validate",
        action="store_true",
        help="Validate existing output against ground truth",
    )

    parser.add_argument(
        "--month",
        type=str,
        default=None,
        help="Process specific month only (e.g., Aug-25)",
    )

    args = parser.parse_args()

    if args.validate:
        # Validation-only mode
        try:
            validate_existing_outputs(args.month)
        except Exception as e:
            logger.error(f"Validation failed: {e}", exc_info=True)
            sys.exit(1)
    else:
        # Pipeline mode - determine service type from args
        service_type: str | None = None
        for flag, stype in SERVICE_TYPES.items():
            if getattr(args, flag, False):
                service_type = stype
                break

        if service_type is None:
            parser.error("No service type specified")
            sys.exit(1)

        assert service_type is not None

        try:
            run_pipeline(service_type, month=args.month)
            logger.info("✓ Pipeline completed successfully")
        except Exception as e:
            logger.error(f"Pipeline failed: {e}", exc_info=True)
            sys.exit(1)


if __name__ == "__main__":
    main()

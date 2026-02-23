#!/usr/bin/env python3
"""VBA Extraction Tool.

Extracts VBA macros from Excel suppression workbooks and saves them as
individual .bas files.
Useful for analyzing and comparing VBA code in FFT suppression workbooks.
"""

import argparse
from pathlib import Path

from oletools.olevba import VBA_Parser


def extract_vba_macros(input_file, output_dir, verbose=False):
    """Extract VBA macros from an Excel file and save them as individual .bas files.

    Args:
        input_file (str): Path to the Excel file containing VBA macros
        output_dir (str): Directory where extracted VBA files will be saved
        verbose (bool): Whether to print detailed output

    Returns:
        dict: Dictionary mapping filenames to VBA code

    """
    try:
        # Convert to Path objects
        input_path = Path(input_file)
        output_path = Path(output_dir)

        if not input_path.exists():
            print(f"Error: Input file '{input_file}' not found.")
            return {}

        # Create output directory if it doesn't exist
        output_path.mkdir(parents=True, exist_ok=True)

        # Extract VBA macros
        vba_parser = VBA_Parser(str(input_path))

        if not vba_parser.detect_vba_macros():
            print(f"No VBA macros found in '{input_file}'.")
            return {}

        if verbose:
            print(f"VBA macros detected in '{input_file}'. Extracting...")

        # Extract and save all macros
        macros = {}
        for filename, stream_path, vba_filename, vba_code in vba_parser.extract_macros():
            if vba_code.strip():  # Only save non-empty macros
                # Clean up the filename
                clean_name = vba_filename.replace("/", "_").replace("\\", "_")
                if clean_name.endswith(".cls"):
                    clean_name = clean_name[:-4] + ".bas"  # Convert .cls to .bas
                elif not clean_name.endswith(".bas"):
                    clean_name += ".bas"

                # Save to file
                output_file = output_path / clean_name
                with open(output_file, "w", encoding="utf-8") as f:
                    f.write(vba_code)

                if verbose:
                    print(f"  Saved: {clean_name}")

                macros[clean_name] = vba_code

        if verbose:
            print(f"\nSuccessfully extracted {len(macros)} VBA modules to '{output_dir}'")
        else:
            print(f"Extracted {len(macros)} VBA modules to '{output_dir}'")

        return macros

    except Exception as e:
        print(f"Error extracting VBA macros: {e}")
        return {}


def main():
    """Extract VBA macros from Excel suppression workbooks."""
    parser = argparse.ArgumentParser(
        description="Extract VBA macros from Excel suppression workbooks",
        epilog=(
            "Example: python extract_vba.py "
            "data/inputs/suppression_files/IP_Suppression_V3.5.xlsm "
            "--output vba_extracted --verbose"
        ),
    )

    parser.add_argument("input_file", help="Path to the Excel file containing VBA macros")

    parser.add_argument(
        "--output",
        "-o",
        default="vba_extracted",
        help="Output directory for extracted VBA files (default: vba_extracted)",
    )

    parser.add_argument(
        "--verbose",
        "-v",
        action="store_true",
        help="Print detailed output during extraction",
    )

    parser.add_argument(
        "--compare",
        "-c",
        help="Compare extracted VBA with existing .bas files in the specified directory",
    )

    args = parser.parse_args()

    # Extract VBA macros
    macros = extract_vba_macros(args.input_file, args.output, args.verbose)

    # If compare flag is set, compare with existing files
    if args.compare and macros:
        compare_dir = Path(args.compare)
        if compare_dir.exists() and compare_dir.is_dir():
            print(f"\nComparing extracted VBA with files in '{args.compare}'...")

            # Find matching files
            matches = 0
            differences = 0

            for extracted_file, extracted_code in macros.items():
                compare_file = compare_dir / extracted_file
                if compare_file.exists():
                    with open(compare_file, encoding="utf-8") as f:
                        compare_code = f.read()

                    # Remove VB_Name attributes for comparison (oletools adds these)
                    extracted_clean = "\n".join(
                        line
                        for line in extracted_code.split("\n")
                        if not line.startswith("Attribute VB_Name")
                    )
                    compare_clean = "\n".join(
                        line
                        for line in compare_code.split("\n")
                        if not line.startswith("Attribute VB_Name")
                    )

                    if extracted_clean.strip() == compare_clean.strip():
                        if args.verbose:
                            print(f"  ✅ {extracted_file}: Identical")
                        matches += 1
                    else:
                        if args.verbose:
                            print(f"  ❌ {extracted_file}: Different")
                        differences += 1

            print("\nComparison results:")
            print(f"  Identical files: {matches}")
            print(f"  Different files: {differences}")

            if differences == 0:
                print("  🎉 All files match! The existing .bas files are accurate.")
            else:
                print("  ⚠️  Some files differ. Check the differences above.")
        else:
            print(f"Error: Comparison directory '{args.compare}' does not exist.")


if __name__ == "__main__":
    main()

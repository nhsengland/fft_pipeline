"""Setup script to create FFT pipeline directory structure."""

import argparse
from pathlib import Path


def create_data_structure(base: Path = None):
    """Create the data directory structure only."""
    if base is None:
        base = Path(".")

    # Data directories only
    data_dirs = [
        "data/inputs/raw",
        "data/inputs/templates",
        "data/inputs/collections_overview",
        "data/inputs/suppression_files",
        "data/outputs",
    ]

    for d in data_dirs:
        (base / d).mkdir(parents=True, exist_ok=True)

    print("✓ Data directory structure created")


def create_package_structure(base: Path = None):
    """Create the src/fft package structure only."""
    if base is None:
        base = Path(".")

    # Package directories
    package_dirs = [
        "src/fft",
        "src/fft/app",
    ]

    for d in package_dirs:
        (base / d).mkdir(parents=True, exist_ok=True)

    # Module files with docstrings
    modules = {
        # Core FFT package
        "src/fft/__init__.py": '"""FFT Pipeline package."""\n',
        "src/fft/__main__.py": '"""CLI entry point for FFT pipeline."""\n',
        "src/fft/config.py": '"""Configuration for FFT pipeline paths, mappings, and constants."""\n',
        "src/fft/loaders.py": '"""Data loading functions."""\n',
        "src/fft/processors.py": '"""Data transformation functions."""\n',
        "src/fft/suppression.py": '"""Suppression logic for data privacy protection."""\n',
        "src/fft/writers.py": '"""Excel output functions."""\n',
        "src/fft/utils.py": '"""Helper utilities."""\n',
        # FastHTML web app subpackage
        "src/fft/app/__init__.py": '"""FastHTML web interface for FFT Pipeline."""\n',
        "src/fft/app/__main__.py": '"""FastHTML app entry point."""\n',
        "src/fft/app/server.py": '"""FastHTML web interface implementation."""\n',
    }

    for file_path, content in modules.items():
        (base / file_path).write_text(content)

    print("✓ Package structure created")


def create_full_structure(base: Path = None):
    """Create both package and data structures."""
    if base is None:
        base = Path(".")

    create_package_structure(base)
    create_data_structure(base)
    print("✓ Full project structure created")


def main():
    """Main CLI interface for setup script."""
    parser = argparse.ArgumentParser(
        description="Create FFT pipeline directory structure",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python setup_structure.py                 # Create full structure (default)
  python setup_structure.py --data-only     # Create only data directories
  python setup_structure.py --package-only  # Create only src/fft package structure
  python setup_structure.py --all           # Create full structure (explicit)
        """
    )

    group = parser.add_mutually_exclusive_group()
    group.add_argument(
        "--data-only",
        action="store_true",
        help="Create only data/ directory structure"
    )
    group.add_argument(
        "--package-only",
        action="store_true",
        help="Create only src/fft/ package structure"
    )
    group.add_argument(
        "--all",
        action="store_true",
        help="Create full structure (package + data)"
    )

    args = parser.parse_args()

    if args.data_only:
        create_data_structure()
    elif args.package_only:
        create_package_structure()
    else:
        # Default: create full structure (--all or no args)
        create_full_structure()


if __name__ == "__main__":
    main()

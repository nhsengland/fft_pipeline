"""Setup script to create FFT pipeline directory structure."""

from pathlib import Path


def create_structure():
    """Create the FFT pipeline package structure."""

    base = Path(".")

    # Directories
    dirs = [
        "src/fft",
        "data/inputs/raw",
        "data/inputs/rolling_totals",
        "data/inputs/templates",
        "data/outputs",
    ]

    for d in dirs:
        (base / d).mkdir(parents=True, exist_ok=True)

    # Module files with docstrings
    modules = {
        "src/fft/__init__.py": '"""FFT Pipeline package."""\n',
        "src/fft/config.py": '"""Configuration for FFT pipeline paths, mappings, and constants."""\n',
        "src/fft/loaders.py": '"""Data loading functions."""\n',
        "src/fft/processors.py": '"""Data transformation functions."""\n',
        "src/fft/suppression.py": '"""Suppression logic for data privacy protection."""\n',
        "src/fft/writers.py": '"""Excel output functions."""\n',
        "src/fft/utils.py": '"""Helper utilities."""\n',
        "src/__main__.py": '"""CLI entry point for FFT pipeline."""\n',
    }

    for file_path, content in modules.items():
        (base / file_path).write_text(content)

    print("✓ Structure created successfully")


if __name__ == "__main__":
    create_structure()

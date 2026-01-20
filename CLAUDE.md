# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

FFT Pipeline is an automated ETL system for processing NHS Friends and Family Test (FFT) data (~2 million responses monthly) into publishable reports with privacy-first suppression. The system processes healthcare data at multiple geographic levels (Ward → Site → Trust → ICB → National) and implements cascading suppression rules to prevent patient identification.

## Key Development Commands

### Environment Setup
```bash
# Create virtual environment and install dependencies
uv venv
uv sync
```

### Running the Pipeline
```bash
# CLI - Process latest 2 months of inpatient data
uv run python -m fft --ip

# CLI - Process specific month
uv run python -m fft --ip --month Aug-25

# CLI - Process A&E data (coming soon)
uv run python -m fft --ae

# CLI - Process ambulance data (coming soon)
uv run python -m fft --amb

# CLI - Validate outputs against ground truth
uv run python -m fft --validate

# Web interface - Opens browser on port 5001
uv run python -m fft.app
```

### Testing and Code Quality
```bash
# Run all doctests (main testing approach)
uv run python -m doctest $(find src/fft/ -name "*.py" -not -name "__main__.py")

# Format code (line length: 90)
uv run ruff format

# Lint code
uv run ruff check

# Type check
uv run ty
```

### Utilities
```bash
# Populate template BS sheets from suppression files
uv run python src/fft/utils.py
```

## Architecture Overview

### Core Modules Structure
- `src/fft/__main__.py` (514 lines) - CLI entry point and 20-step processing orchestration
- `src/fft/config.py` (596 lines) - Centralized configuration hub with paths, mappings, and constants
- `src/fft/processors.py` (869 lines) - Data transformation pipeline (standardize, aggregate, calculate)
- `src/fft/writers.py` (811 lines) - Excel output generation with VBA macro preservation
- `src/fft/suppression.py` (525 lines) - Privacy protection via three-level cascading suppression
- `src/fft/validation.py` (705 lines) - Ground truth comparison and output verification
- `src/fft/loaders.py` (166 lines) - Data loading from raw Excel files
- `src/fft/app/server.py` (1,171 lines) - FastHTML web interface with HTMX real-time updates

### Data Processing Flow
1. **Load** raw Excel files from `data/inputs/raw/` (FFT_IP_V1*.xlsx format)
2. **Transform** through processors.py: standardize columns → remove unwanted → aggregate by level
3. **Suppress** via cascading rules: first-level (1-4 responses) → second-level → cascade to children
4. **Write** to Excel templates from `data/inputs/templates/` preserving VBA macros
5. **Output** to `data/outputs/` as publishable .xlsm files

### Geographic Hierarchy Processing
- **Inpatient**: Ward → Site → Organisation → ICB → National (4 sheets)
- **A&E**: Site → Organisation → ICB → National (3 sheets) - coming soon
- **Ambulance**: Organisation → ICB → National (2 sheets) - coming soon

### Configuration System
All processing behavior is centralized in `config.py`:
- File paths and patterns
- Column mappings per service type/level
- Processing levels and sheet configurations
- Suppression thresholds (currently 4)
- Template and validation settings

## Privacy and Suppression Rules

This system implements NHS England's cascading suppression approach:

1. **First-level**: Any organization with 1-4 responses gets all Likert values replaced with `*`
2. **Second-level**: The next-lowest responding organization also gets suppressed (prevents reverse calculation)
3. **Cascade**: If parent level is suppressed, the two lowest-responding children are also suppressed

Example: If an ICB has a Trust with 2 responses, the ICB is marked for suppression, causing the two lowest-responding child Trusts to be suppressed even if they individually have >4 responses.

## Key Implementation Patterns

### Data Handling
- All data processing uses pandas DataFrames
- Column standardization happens at each geographic level
- Excel files loaded via openpyxl to preserve VBA macros
- Template-based output generation maintains existing formatting

### Testing Approach
- Primary testing via inline doctests (not pytest)
- Ground truth validation compares outputs against reference files
- Validation handles both key-based and range-based comparisons

### Error Handling
- Comprehensive logging at INFO level for pipeline steps
- Detailed validation reports with diff summaries
- File not found errors provide helpful suggestions

## File Locations and Patterns

### Input Data Structure
```
data/inputs/raw/                    # Raw Excel files (excluded from git)
├── FFT_IP_V1 Aug-25.xlsx          # Inpatient raw data
├── FFT_AE_V1 Aug-25.xlsx          # A&E raw data
└── FFT_Amb_V1 Aug-25.xlsx         # Ambulance raw data

data/inputs/templates/              # Excel templates (tracked in git)
├── FFT_IP_template.xlsm
├── FFT_AE_template.xlsm
└── FFT_Amb_template.xlsm

data/inputs/suppression_files/      # VBA reference files (excluded from git)
data/inputs/collections_overview/  # Collection metadata (excluded from git)
```

### Output Structure
```
data/outputs/
├── FFT-inpatient-data-Aug-25.xlsm     # Processed outputs
├── FFT-ae-data-Aug-25.xlsm
└── ground_truth/                       # Validation references (excluded from git)
```

## Development Environment

- **Python**: 3.13+ required
- **Dependency manager**: `uv` (not pip/poetry)
- **Code style**: Ruff with Black compatibility, 90 character line length
- **Testing**: Doctests only (no pytest framework)
- **Type checking**: `ty` for static analysis
- **Virtual environment**: `.venv/` in project root

## Important Development Notes

### Data Sensitivity
- This processes NHS patient feedback data - handle with care
- All patient data files are excluded from git via `.gitignore`
- Only templates and configuration files are tracked
- GDPR compliance is essential

### Service Type Status
- **Inpatient**: Fully implemented and tested
- **A&E**: In development (branch 36, 38, 40-42 active)
- **Ambulance**: Future implementation

### Branch Strategy
- Main development branch: `main`
- Feature branches follow numeric naming: `37-suppression-rule-inconsistency-...`
- Current active development on suppression rule fixes and validation improvements

### Configuration Changes
When modifying processing logic, update `config.py` rather than hardcoding values. The configuration system controls:
- Column name mappings per service type
- Processing levels and sheet names
- File patterns and paths
- Suppression rules and thresholds

### Excel Template System
- Templates must be .xlsm files to preserve VBA macros
- BS sheets contain lookup data populated from suppression reference files
- Output formatting and formulas are preserved from templates
- Never modify templates directly - use the BS population utility

This codebase represents a production healthcare data processing system with comprehensive privacy controls and dual CLI/web interfaces for operational flexibility.
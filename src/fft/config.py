"""Configuration for FFT pipeline."""

from pathlib import Path

BASE_DIR = Path(__file__).parent.parent.parent
DATA_DIR = BASE_DIR / "data"
INPUTS_DIR = DATA_DIR / "inputs"
RAW_DIR = INPUTS_DIR / "raw"
ROLLING_TOTALS_DIR = INPUTS_DIR / "rolling_totals"
TEMPLATES_DIR = INPUTS_DIR / "templates"
OUTPUTS_DIR = DATA_DIR / "outputs"

SUPPRESSION_THRESHOLD = 5

# File patterns for each service type
FILE_PATTERNS = {
    "inpatient": "FFT_Inpatients_V1*.xlsx",
    "ae": "FFT_AE_V1*.xlsx",
    "ambulance": "FFT_Ambulance_V1*.xlsx",
}


# Column mappings for each service type and level
COLUMN_MAPS = {
    "inpatient": {
        "organisation": {
            "Parent org code": "ICB_Code",
            "Parent name": "ICB_Name",
            # ... more mappings
        },
        # "site": {...},
        # "ward": {...},
    },
    # "ae": {
    #     "organisation": {...},
    #        "site": {...}},
}


# Month abbreviations for FFT period formatting
MONTH_ABBREV = {
    "JANUARY": "Jan",
    "FEBRUARY": "Feb",
    "MARCH": "Mar",
    "APRIL": "Apr",
    "MAY": "May",
    "JUNE": "Jun",
    "JULY": "Jul",
    "AUGUST": "Aug",
    "SEPTEMBER": "Sep",
    "OCTOBER": "Oct",
    "NOVEMBER": "Nov",
    "DECEMBER": "Dec",
}


# Columns to remove from raw data at each level
COLUMNS_TO_REMOVE = {
    "inpatient": {
        "organisation": ["Yearnumber", "Periodname", "Title", "Response Rate"],
        "site": ["Yearnumber", "Periodname", "Title", "Response Rate"],
        "ward": ["Yearnumber", "Periodname", "Title", "Response Rate"],
    },
    # Add ae, ambulance later
}

# Validation rules for data quality checks
VALIDATION_RULES = {
    "inpatient": {
        "column_lengths": {
            "Yearnumber": [7],
            "Org code": [3, 5],
            "Parent org code": [3],
        },
        "numeric_columns": {
            "int": [
                "1 Very Good",
                "2 Good",
                "3 Neither good nor poor",
                "4 Poor",
                "5 Very poor",
                "6 Dont Know",
                "Total Responses",
                "Total Eligible",
            ],
            "float": ["Prop_Pos", "Prop_Neg"],
        },
    },
    # Add ae, ambulance later
}

"""Configuration for FFT pipeline paths, mappings, and constants."""

from pathlib import Path

# =============================================================================
# PATHS
# =============================================================================

BASE_DIR = Path(__file__).parent.parent.parent
DATA_DIR = BASE_DIR / "data"
INPUTS_DIR = DATA_DIR / "inputs"
RAW_DIR = INPUTS_DIR / "raw"
ROLLING_TOTALS_DIR = INPUTS_DIR / "rolling_totals"
TEMPLATES_DIR = INPUTS_DIR / "templates"
OUTPUTS_DIR = DATA_DIR / "outputs"

# =============================================================================
# FILE PATTERNS
# =============================================================================

FILE_PATTERNS = {
    "inpatient": "FFT_Inpatients_V1*.xlsx",
    "ae": "FFT_AE_V1*.xlsx",
    "ambulance": "FFT_Ambulance_V1*.xlsx",
}

# =============================================================================
# REUSABLE COLUMN GROUPS
# =============================================================================

# Geographic identifiers
ICB_COLS = ["ICB Code", "ICB Name"]
TRUST_COLS = ["Trust Code", "Trust Name"]
SITE_COLS = ["Site Code", "Site Name"]
WARD_COLS = ["Ward Name"]

# Core data columns
TOTALS_COLS = ["Total Responses", "Total Eligible"]
PERCENTAGE_COLS = ["Percentage Positive", "Percentage Negative"]
LIKERT_COLS = [
    "Very Good",
    "Good",
    "Neither Good nor Poor",
    "Poor",
    "Very Poor",
    "Don't Know",
]

# Collection mode columns
MODE_COLS = [
    "Mode SMS",
    "Mode Electronic Discharge",
    "Mode Electronic Home",
    "Mode Paper Discharge",
    "Mode Paper Home",
    "Mode Telephone",
    "Mode Online",
    "Mode Other",
]

# Specialty columns (ward level only)
SPECIALTY_COLS = ["First Speciality", "Second Speciality"]

# =============================================================================
# MONTH ABBREVIATIONS
# =============================================================================

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

# =============================================================================
# COLUMN MAPPINGS (raw data → standardised names)
# =============================================================================

# Column mappings for each service type and level
COLUMN_MAPS = {
    "inpatient": {
        "ward": {
            "Parent org code": "ICB_Code",
            "Parent name": "ICB_Name",
            "Org code": "Trust_Code",
            "Org name": "Trust_Name",
            "Site code": "Site_Code",
            "Site name": "Site_Name",
            "Ward name": "Ward_Name",
            "1 Very Good SUM": "Very Good",
            "2 Good SUM": "Good",
            "3 Neither Good nor Poor SUM": "Neither Good nor Poor",
            "4 Poor SUM": "Poor",
            "5 Very Poor SUM": "Very Poor",
            "6 Dont Know SUM": "Don't Know",
            "Total Eligible SUM": "Total Eligible",
            "Spec 1": "First Speciality",
            "Spec 2": "Second Speciality",
            "Prop_Pos": "Percentage_Positive",
            "Prop_Neg": "Percentage_Negative",
        },
        "site": {
            "Parent org code": "ICB_Code",
            "Parent name": "ICB_Name",
            "Org code": "Trust_Code",
            "Org name": "Trust_Name",
            "Site Code": "Site_Code",
            "Site Name MAX": "Site_Name",
            "1 Very Good SUM": "Very Good",
            "2 Good SUM": "Good",
            "3 Neither Good nor Poor SUM": "Neither Good nor Poor",
            "4 Poor SUM": "Poor",
            "5 Very Poor SUM": "Very Poor",
            "6 Dont Know SUM": "Don't Know",
            "Total Eligible SUM": "Total Eligible",
            "Prop_Pos": "Percentage_Positive",
            "Prop_Neg": "Percentage_Negative",
        },
        "organisation": {
            "Parent org code": "ICB_Code",
            "Parent name": "ICB_Name",
            "Org code": "Trust_Code",
            "Org name": "Trust_Name",
            "1 Very Good SUM": "Very Good",
            "2 Good SUM": "Good",
            "3 Neither Good nor Poor SUM": "Neither Good nor Poor",
            "4 Poor SUM": "Poor",
            "5 Very Poor SUM": "Very Poor",
            "6 Dont Know SUM": "Don't Know",
            "Total Eligible SUM": "Total Eligible",
            "Prop_Pos": "Percentage_Positive",
            "Prop_Neg": "Percentage_Negative",
        },
    },
    # Add ae, ambulance later
}

# =============================================================================
# COLUMNS TO REMOVE
# =============================================================================

COLUMNS_TO_REMOVE = {
    "inpatient": {
        "organisation": ["Yearnumber", "Periodname", "Title", "Response Rate"],
        "site": ["Yearnumber", "Periodname", "Title", "Response Rate"],
        "ward": ["Yearnumber", "Periodname", "Title", "Response Rate"],
    },
    # Add ae, ambulance later
}


# =============================================================================
# AGGREGATION COLUMNS
# =============================================================================


AGGREGATION_COLUMNS = {
    "likert_responses": [
        "Very Good",
        "Good",
        "Neither Good nor Poor",
        "Poor",
        "Very Poor",
        "Don't Know",
    ],
    "totals": ["Total Responses", "Total Eligible"],
    "collection_modes": [
        "Mode SMS",
        "Mode Electronic Discharge",
        "Mode Electronic Home",
        "Mode Paper Discharge",
        "Mode Paper Home",
        "Mode Telephone",
        "Mode Online",
        "Mode Other",
    ],
}

# =============================================================================
# OUTPUT COLUMNS
# =============================================================================

# Output columns per sheet (in order)
OUTPUT_COLUMNS = {
    "inpatient": {
        "ICB": [
            "ICB_Code",
            "ICB_Name",
            "Total Responses",
            "Total Eligible",
            "Percentage_Positive",
            "Percentage_Negative",
            "Very Good",
            "Good",
            "Neither Good nor Poor",
            "Poor",
            "Very Poor",
            "Don't Know",
        ],
        "Trusts": [
            "ICB_Code",
            "ICB_Name",
            "Trust_Code",
            "Trust_Name",
            "Total Responses",
            "Total Eligible",
            "Percentage_Positive",
            "Percentage_Negative",
            "Very Good",
            "Good",
            "Neither Good nor Poor",
            "Poor",
            "Very Poor",
            "Don't Know",
            "Mode SMS",
            "Mode Electronic Discharge",
            "Mode Electronic Home",
            "Mode Paper Discharge",
            "Mode Paper Home",
            "Mode Telephone",
            "Mode Online",
            "Mode Other",
        ],
        "Sites": [
            "ICB_Code",
            "ICB_Name",
            "Trust_Code",
            "Trust_Name",
            "Site_Code",
            "Site_Name",
            "Total Responses",
            "Total Eligible",
            "Percentage_Positive",
            "Percentage_Negative",
            "Very Good",
            "Good",
            "Neither Good nor Poor",
            "Poor",
            "Very Poor",
            "Don't Know",
        ],
        "Wards": [
            "ICB_Code",
            "ICB_Name",
            "Trust_Code",
            "Trust_Name",
            "Site_Code",
            "Site_Name",
            "Ward_Name",
            "Total Responses",
            "Total Eligible",
            "Percentage_Positive",
            "Percentage_Negative",
            "Very Good",
            "Good",
            "Neither Good nor Poor",
            "Poor",
            "Very Poor",
            "Don't Know",
            "First Speciality",
            "Second Speciality",
        ],
    }
}

# =============================================================================
# SUPPRESSION
# =============================================================================

SUPPRESSION_THRESHOLD = 5  # Responses < 5 get suppressed

# =============================================================================
# TEMPLATE CONFIGURATION
# =============================================================================

TEMPLATE_CONFIG = {
    "inpatient": {
        "template_file": "FFT_IP_template.xlsm",
        "output_prefix": "FFT-inpatient-data",
        "data_start_row": 15,
        "england_rows": {"including_is": 12, "excluding_is": 13, "selection": 14},
        "sheets": {
            "icb": {
                "sheet_name": "ICB",
                "name_column": "ICB_Name",
                "columns": [*ICB_COLS, *TOTALS_COLS, *PERCENTAGE_COLS, *LIKERT_COLS],
            },
            "organisation": {
                "sheet_name": "Trusts",
                "name_column": "Trust_Name",
                "columns": [
                    *ICB_COLS,
                    *TRUST_COLS,
                    *TOTALS_COLS,
                    *PERCENTAGE_COLS,
                    *LIKERT_COLS,
                    *MODE_COLS,
                ],
            },
            "site": {
                "sheet_name": "Sites",
                "name_column": "Site_Name",
                "columns": [
                    *ICB_COLS,
                    *TRUST_COLS,
                    *SITE_COLS,
                    *TOTALS_COLS,
                    *PERCENTAGE_COLS,
                    *LIKERT_COLS,
                ],
            },
            "ward": {
                "sheet_name": "Wards",
                "name_column": "Ward_Name",
                "columns": [
                    *ICB_COLS,
                    *TRUST_COLS,
                    *SITE_COLS,
                    *WARD_COLS,
                    *TOTALS_COLS,
                    *PERCENTAGE_COLS,
                    *LIKERT_COLS,
                    *SPECIALTY_COLS,
                ],
            },
        },
    },
    # Add ae, ambulance later using same composable pattern
}

# =============================================================================
# BS SHEET CONFIGURATION
# =============================================================================

# BS Sheet column positions (1-indexed)
BS_SHEET_CONFIG = {
    "inpatient": {
        "reference_list_start_col": 21,  # Column U
        "reference_list_start_row": 2,
        "reference_columns": [
            "ICB_Code",
            "ICB_Name",
            "Trust_Code",
            "Trust_Name",
            "Site_Code",
            "Site_Name",
        ],
        "linked_lists": {
            "trusts": {"start_col": 31, "columns": ["Trust_Code", "Trust_Name"]},
            "sites": {
                "start_col": 34,
                "columns": ["Trust_Code", "Trust_Name", "Site_Code", "Site_Name"],
            },
            "wards": {
                "start_col": 39,
                "columns": [
                    "Trust_Code",
                    "Trust_Name",
                    "Site_Code",
                    "Site_Name",
                    "Ward_Name",
                ],
            },
        },
    }
}

# =============================================================================
# PERIOD LABEL CONFIGURATION
# =============================================================================

# Period label configuration (cells that need FFT period updated)
PERIOD_LABEL_CONFIG = {
    "inpatient": {
        "notes_title": {
            "sheet": "Notes",
            "cell": "A2",
            "template": "Inpatient Friends and Family Test (FFT) Data - {period}",
        }
    },
    "ae": {
        "notes_title": {
            "sheet": "Notes",
            "cell": "A2",
            "template": "A&E Friends and Family Test (FFT) Data - {period}",
        }
    },
    "ambulance": {
        "notes_title": {
            "sheet": "Notes",
            "cell": "A2",
            "template": "Ambulance Friends and Family Test (FFT) Data - {period}",
        }
    },
}

# =============================================================================
# PERCENTAGE COLUMN POSITIONS
# =============================================================================

# Percentage column positions per sheet (1-indexed)
PERCENTAGE_COLUMN_CONFIG = {
    "inpatient": {
        "ICB": [5, 6],  # Columns E, F (Percentage Positive, Percentage Negative)
        "Trusts": [6, 7],  # Columns F, G
        "Sites": [8, 9],  # Columns H, I
        "Wards": [9, 10],  # Columns I, J
    }
}

# =============================================================================
# PROCESSING LEVELS PER SERVICE TYPE
# =============================================================================

# Processing levels per service type (in order of processing)
PROCESSING_LEVELS = {
    "inpatient": {
        "levels": ["ward", "site", "organisation"],
        "sheet_mapping": {
            "ward": "Parent & Self Trusts - Ward Lev",
            "site": "Parent & Self Trusts - Site Lev",
            "organisation": "Parent & Self Trusts - Organisa",
            "collection_mode": "Parent & Self Trusts - Collecti",
        },
    },
    "ae": {
        "levels": ["site", "organisation"],
        "sheet_mapping": {
            "site": "Parent & Self - Site Level",
            "organisation": "Parent & Self - Organisation Le",
            "collection_mode": "Parent & Self - Collection Mode",
        },
    },
    "ambulance": {
        "levels": ["organisation"],
        "sheet_mapping": {
            "organisation": "Organisation_Level_PTS",
            "collection_mode": "Collection Mode",
        },
    },
}

# =============================================================================
# CLI SERVICE TYPE MAPPINGS
# =============================================================================

SERVICE_TYPES = {
    "ip": "inpatient",
    "ae": "ae",
    "ambulance": "ambulance",
    # Add new service types here:
    # 'op': 'outpatient',
    # 'maternity': 'maternity',
}

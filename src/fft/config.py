"""Configuration for FFT pipeline paths, mappings, and constants."""

from pathlib import Path
from typing import TypedDict

# =============================================================================
# PATHS
# =============================================================================

BASE_DIR = Path(__file__).parent.parent.parent
DATA_DIR = BASE_DIR / "data"
INPUTS_DIR = DATA_DIR / "inputs"
RAW_DIR = INPUTS_DIR / "raw"
SUPPRESSION_FILES_DIR = INPUTS_DIR / "suppression_files"
TEMPLATES_DIR = INPUTS_DIR / "templates"
OUTPUTS_DIR = DATA_DIR / "outputs"
COLLECTIONS_OVERVIEW_DIR = INPUTS_DIR / "collections_overview"
COLLECTIONS_OVERVIEW_FILE = "_FFT_CollectionOverview V1 5.xlsm"

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
SPECIALITY_COLS = ["First Speciality", "Second Speciality"]

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
            # "Prop_Neg": "Percentage_Negative",  # Removed - input data incorrect
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
            # "Prop_Neg": "Percentage_Negative",  # Removed - input data incorrect
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
            # "Prop_Neg": "Percentage_Negative",  # Removed - input data incorrect
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


class EnglandRowsConfig(TypedDict):
    """Type definition for england_rows configuration."""

    including_is: int
    excluding_is: int
    selection: int


class SheetConfig(TypedDict):
    """Type definition for individual sheet configuration."""

    sheet_name: str
    name_column: str
    england_label_column: str
    columns: list[str]


class TemplateServiceConfig(TypedDict):
    """Type definition for template service configuration."""

    template_file: str
    output_prefix: str
    data_start_row: int
    england_rows: EnglandRowsConfig
    sheets: dict[str, SheetConfig]


TEMPLATE_CONFIG: dict[str, TemplateServiceConfig] = {
    "inpatient": {
        "template_file": "FFT_IP_template.xlsm",
        "output_prefix": "FFT-inpatient-data",
        "data_start_row": 15,
        "england_rows": {"including_is": 12, "excluding_is": 13, "selection": 14},
        "sheets": {
            "icb": {
                "sheet_name": "ICB",
                "name_column": "ICB_Name",
                "england_label_column": "ICB_Name",
                "columns": [*ICB_COLS, *TOTALS_COLS, *PERCENTAGE_COLS, *LIKERT_COLS],
            },
            "organisation": {
                "sheet_name": "Trusts",
                "name_column": "Trust_Name",
                "england_label_column": "Trust_Name",
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
                "england_label_column": "Site_Name",
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
                "england_label_column": "Ward_Name",
                "columns": [
                    *ICB_COLS,
                    *TRUST_COLS,
                    *SITE_COLS,
                    *WARD_COLS,
                    *TOTALS_COLS,
                    *PERCENTAGE_COLS,
                    *LIKERT_COLS,
                    *SPECIALITY_COLS,
                ],
            },
        },
    },
    # Add ae, ambulance later using same composable pattern
}

# =============================================================================
# VALIDATION CONFIGURATION
# =============================================================================

# Sheets to validate for each service type (derived from template config)
VALIDATION_CONFIG: dict[str, list[str]] = {
    service_type: [
        sheet_config["sheet_name"] for sheet_config in template_config["sheets"].values()
    ]
    for service_type, template_config in TEMPLATE_CONFIG.items()
}

# Key columns for record matching during validation (Excel column letters)
# Single column (str) or composite key (list of str) for unique identification
VALIDATION_KEY_COLUMNS: dict[str, str | list[str]] = {
    "ICB": "B",  # ICB_Code
    "Trusts": "B",  # Trust_Code
    "Sites": "D",  # Site_Code
    "Wards": ["B", "D", "F"],  # Trust_Code + Site_Code + Ward_Name (composite key)
}

# Tolerance for floating point comparisons during validation
VALIDATION_TOLERANCE: float = 1e-5


# =============================================================================
# BS SHEET CONFIGURATION
# =============================================================================

# BS Sheet column positions (1-indexed)


class LinkedListConfig(TypedDict):
    """Type definition for linked list configuration."""

    start_col: int
    pairs: list[list[str]]


class BSSheetServiceConfig(TypedDict):
    """Type definition for BS sheet service configuration."""

    reference_list_start_col: int
    reference_list_start_row: int
    reference_columns: list[str]
    linked_lists: dict[str, LinkedListConfig]


BS_SHEET_CONFIG: dict[str, BSSheetServiceConfig] = {
    "inpatient": {
        "reference_list_start_col": 21,  # Column U
        "reference_list_start_row": 2,
        "reference_columns": [
            "ICB_Code",
            "Trust_Code",
            "Trust_Name",
            "Site_Code",
            "Site_Name",
            "Ward_Name",
        ],
        "linked_lists": {
            "trusts": {
                "start_col": 31,  # AE
                "pairs": [["Trust_Code", "Trust_Name"]],
            },
            "sites": {
                "start_col": 34,  # AH
                "pairs": [["Trust_Code", "Trust_Name"], ["Site_Code", "Site_Name"]],
            },
            "wards": {
                "start_col": 39,  # AM
                "pairs": [
                    ["Trust_Code", "Trust_Name"],
                    ["Site_Code", "Site_Name"],
                    ["Ward_Name"],
                ],
            },
        },
    }
}

# =============================================================================
# PERIOD LABEL CONFIGURATION
# =============================================================================

# Period label configuration (cells that need FFT period updated)


class PeriodLabelCellConfig(TypedDict):
    """Type definition for period label cell configuration."""

    sheet: str
    cell: str
    template: str


PERIOD_LABEL_CONFIG: dict[str, dict[str, PeriodLabelCellConfig]] = {
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
PERCENTAGE_COLUMN_CONFIG: dict[str, dict[str, list[int]]] = {
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
    "amb": "ambulance",
    # Add new service types here:
    # "op": "outpatient",
    # "mat": "maternity",
}

# =============================================================================
# COLLECTIONS OVERVIEW CONFIGURATION
# =============================================================================

# Time series column prefixes for each service type
TIME_SERIES_PREFIXES = {
    "inpatient": "Inpatient",
    "ae": "A&E",
    "ambulance": "Ambulance",
    "outpatient": "Outpatient",
    "maternity": "Q1",  # Maternity uses Q1-Q4 format
    "community": "CH",
    "mental_health": "MH",
    "gp": "GP",
    "dental": "Dental",
    "post_covid": "Lcov Q1",
}


# Summary data column suffixes (appended to service prefix)
SUMMARY_COLUMNS = {
    "orgs_submitting": {
        "total": " Submitted",
        "nhs": " NHS Submitted",
        "is": " IS Submitted",
    },
    "responses": {
        "total": " Responses",
        "nhs": " NHS Responses",
        "is": " IS Responses",
    },
    "positive": {
        "likely": " Likely",
        "extremely_likely": " Extremely Likely",
        "nhs_likely": " NHS Likely",
        "nhs_extremely_likely": " NHS Extremely Likely",
        "is_likely": " IS Likely",
        "is_extremely_likely": " IS Extremely Likely",
    },
    "negative": {
        "unlikely": " Unlikely",
        "extremely_unlikely": " Extremely Unlikely",
        "nhs_unlikely": " NHS Unlikely",
        "nhs_extremely_unlikely": " NHS Extremely Unlikely",
        "is_unlikely": " IS Unlikely",
        "is_extremely_unlikely": " IS Extremely Unlikely",
    },
}

# =============================================================================
# VALIDATION CONFIGURATION
# =============================================================================

# Tolerance for floating point comparisons during validation
VALIDATION_TOLERANCE: float = 1e-8

# Provider type constants
IS1_CODE = "IS1"
IS1_NAME = "INDEPENDENT SECTOR PROVIDERS"
NHS_PROVIDER_KEYWORDS = ["NHS", "TRUST"]

# Sheet and data markers
SUPPRESSION_MARKER = "*"

# England rows column skip counts for Mode of Collection fix
ENGLAND_ROWS_SKIP_COLUMNS = {
    "ICB": 2,
    "Trusts": 3,
    "Sites": 4,
    "Wards": 5,
}

"""Configuration for FFT pipeline paths, mappings, and constants."""

from pathlib import Path
from typing import TypedDict

BASE_DIR = Path(__file__).parent.parent.parent
DATA_DIR = BASE_DIR / "data"
INPUTS_DIR = DATA_DIR / "inputs"
RAW_DIR = INPUTS_DIR / "raw"
SUPPRESSION_FILES_DIR = INPUTS_DIR / "suppression_files"
TEMPLATES_DIR = INPUTS_DIR / "templates"
OUTPUTS_DIR = DATA_DIR / "outputs"
COLLECTIONS_OVERVIEW_DIR = INPUTS_DIR / "collections_overview"
COLLECTIONS_OVERVIEW_FILE = "_FFT_CollectionOverview V1 5.xlsm"

FILE_PATTERNS = {
    "inpatient": "FFT_Inpatients_V1*.xlsx",
    "ae": "FFT_A&E_V1*.xlsx",
    "ambulance": "FFT_Ambulance_V1*.xlsx",
}

ICB_COLS = ["ICB Code", "ICB Name"]
TRUST_COLS = ["Trust Code", "Trust Name"]
SITE_COLS = ["Site Code", "Site Name"]
WARD_COLS = ["Ward Name"]
SPECIALITY_COLS = ["First Speciality", "Second Speciality"]
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
_STD_PCT_COLS = ["Percentage_Positive", "Percentage_Negative"]

ALL_MODES = [
    "Mode SMS",
    "Mode Electronic Discharge",
    "Mode Electronic Home",
    "Mode Paper Discharge",
    "Mode Paper Home",
    "Mode Telephone",
    "Mode Online",
    "Mode Other",
]
_AE_EXCLUDED_MODES = {"Mode Electronic Home"}
_AMBULANCE_EXCLUDED_MODES = {"Mode Electronic Home"}
MODE_COLS = {
    "inpatient": ALL_MODES,
    "ae": [m for m in ALL_MODES if m not in _AE_EXCLUDED_MODES],
    "ambulance": [m for m in ALL_MODES if m not in _AMBULANCE_EXCLUDED_MODES],
}
COMMON_MODES = [m for m in ALL_MODES if all(m in modes for modes in MODE_COLS.values())]
SERVICE_SPECIFIC_MODES = {
    s: [m for m in modes if m not in COMMON_MODES] for s, modes in MODE_COLS.items()
}

AGGREGATION_COLUMNS = {
    "likert_responses": LIKERT_COLS,
    "totals": TOTALS_COLS,
    "collection_modes": ALL_MODES,
}
COUNT_COLUMNS = {"common": LIKERT_COLS + TOTALS_COLS}


def get_count_columns_for_service(service_type):
    """Get count columns for a specific service type."""
    return (
        COUNT_COLUMNS["common"]
        + COMMON_MODES
        + SERVICE_SPECIFIC_MODES.get(service_type, [])
    )


_BASE_ID = {
    "Parent org code": "ICB_Code",
    "Parent name": "ICB_Name",
    "Org code": "Trust_Code",
    "Org name": "Trust_Name",
}
_AMB_ORG_ID = {
    "Parent org code": "ICB_Code",
    "Parent name": "ICB_Name",
    "Org code": "Trust_Code",
    "Org name": "Trust_Name",  # Standard mapping
    "Organisation Name": "Trust_Name",  # Map ambulance template's "Organisation Name" to standard "Trust_Name"
}
_IP_DATA = {
    "1 Very Good SUM": "Very Good",
    "2 Good SUM": "Good",
    "3 Neither Good nor Poor SUM": "Neither Good nor Poor",
    "4 Poor SUM": "Poor",
    "5 Very Poor SUM": "Very Poor",
    "6 Dont Know SUM": "Don't Know",
    "Total Eligible SUM": "Total Eligible",
    "Prop_Pos": "Percentage_Positive",
}
_AE_DATA = {
    "1 Very Good": "Very Good",
    "2 Good": "Good",
    "3 Neither good nor poor": "Neither Good nor Poor",
    "4 Poor": "Poor",
    "5 Very poor": "Very Poor",
    "6 Dont Know": "Don't Know",
    "Total Eligible": "Total Eligible",
    "Prop_Pos": "Percentage_Positive",
}
_AMB_DATA = {
    "1 Very Good": "Very Good",
    "2 Good": "Good",
    "3 Neither good nor poor": "Neither Good nor Poor",  # Match raw data case
    "4 Poor": "Poor",
    "5 Very poor": "Very Poor",  # Match raw data case
    "6 Dont Know": "Don't Know",  # Match raw data (no apostrophe in source)
    "Total Responses": "Total Responses",
    "Total Eligible": "Total Eligible",
    # Percentage columns don't exist in raw data - they're calculated automatically
}
_SITE_ID = {"Site Code": "Site_Code", "Site Name MAX": "Site_Name"}
_WARD_ID = {
    "Site code": "Site_Code",
    "Site name": "Site_Name",
    "Ward name": "Ward_Name",
    "Spec 1": "First Speciality",
    "Spec 2": "Second Speciality",
}

COLUMN_MAPS = {
    "inpatient": {
        "ward": {**_BASE_ID, **_WARD_ID, **_IP_DATA},
        "site": {**_BASE_ID, **_SITE_ID, **_IP_DATA},
        "organisation": {**_BASE_ID, **_IP_DATA},
    },
    "ae": {
        "site": {**_BASE_ID, **_SITE_ID, **_AE_DATA},
        "organisation": {**_BASE_ID, **_AE_DATA},
    },
    "ambulance": {
        "organisation": {**_AMB_ORG_ID, **_AMB_DATA},
    },
}

_COLS_TO_REMOVE = ["Yearnumber", "Periodname", "Title", "Response Rate"]
COLUMNS_TO_REMOVE = {
    "inpatient": {level: _COLS_TO_REMOVE for level in ("organisation", "site", "ward")},
    "ae": {level: _COLS_TO_REMOVE for level in ("organisation", "site")},
    "ambulance": {"organisation": _COLS_TO_REMOVE},
}

_OUT_ICB = ["ICB_Code", "ICB_Name"]
_OUT_TRUST = ["ICB_Code", "Trust_Code", "Trust_Name"]
_OUT_SITE = _OUT_TRUST + ["Site_Code", "Site_Name"]
_OUT_WARD = _OUT_SITE + ["Ward_Name"]
_OUT_DATA = TOTALS_COLS + _STD_PCT_COLS + LIKERT_COLS
_AMB_OUT_DATA = TOTALS_COLS + _STD_PCT_COLS + LIKERT_COLS  # Ambulance has same order as other services
_AMB_OUT_ORG = ["ICB_Code", "Trust_Code", "Trust_Name"]  # Ambulance org columns: ICB Code, Org Code, Organisation Name

OUTPUT_COLUMNS = {
    "inpatient": {
        "ICB": _OUT_ICB + _OUT_DATA,
        "Trusts": _OUT_TRUST + _OUT_DATA + MODE_COLS["inpatient"],
        "Sites": _OUT_SITE + _OUT_DATA,
        "Wards": _OUT_WARD + _OUT_DATA + SPECIALITY_COLS,
    },
    "ae": {
        "ICB": _OUT_ICB + _OUT_DATA,
        "Trusts": _OUT_TRUST + _OUT_DATA + MODE_COLS["ae"],
        "Sites": _OUT_SITE + _OUT_DATA,
    },
    "ambulance": {
        "PTS ICB": _OUT_ICB + _AMB_OUT_DATA,
        "PTS Org": _AMB_OUT_ORG + _AMB_OUT_DATA,
        "Mode Org": _AMB_OUT_ORG + ["Total Responses"] + MODE_COLS["ambulance"],
    },
}

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

SUPPRESSION_THRESHOLD = 5
SUPPRESSION_MARKER = "*"
VALIDATION_TOLERANCE = 1e-4
IS1_CODE = "IS1"
IS1_NAME = "INDEPENDENT SECTOR PROVIDERS"
NHS_PROVIDER_KEYWORDS = ["NHS", "TRUST"]


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
                    *MODE_COLS["inpatient"],
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
    "ae": {
        "template_file": "FFT_AE_template.xlsm",
        "output_prefix": "FFT-ae-data",
        "data_start_row": 7,
        "england_rows": {"including_is": 5, "excluding_is": 5, "selection": 6},
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
                    *MODE_COLS["ae"],
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
        },
    },
    "ambulance": {
        "template_file": "FFT_Amb_template.xlsm",
        "output_prefix": "FFT-ambulance-data",
        "data_start_row": 13,
        "england_rows": {"including_is": 10, "excluding_is": 11, "selection": 12},
        "sheets": {
            "icb": {
                "sheet_name": "PTS ICB",
                "name_column": "ICB_Name",
                "england_label_column": "ICB_Name",
                "columns": [*ICB_COLS, *TOTALS_COLS, *_STD_PCT_COLS, *LIKERT_COLS],
            },
            "organisation": {
                "sheet_name": "PTS Org",
                "name_column": "Trust_Name",
                "england_label_column": "Trust_Name",
                "columns": [
                    "ICB Code",  # Only ICB Code, no ICB Name in ambulance PTS Org template
                    *TRUST_COLS,
                    *TOTALS_COLS,
                    *_STD_PCT_COLS,
                    *LIKERT_COLS,
                    *MODE_COLS["ambulance"],
                ],
            },
            "mode_org": {
                "sheet_name": "Mode Org",
                "name_column": "Trust_Name",
                "england_label_column": "Trust_Name",
                "columns": [
                    "ICB Code",  # Only ICB Code, no ICB Name in ambulance Mode Org template
                    *TRUST_COLS,
                    "Total Responses",
                    *MODE_COLS["ambulance"],
                ],
            },
        },
    },
}


class LinkedListConfig(TypedDict):
    """Type definition for linked list configuration."""

    start_col: int
    pairs: list[list[str]]


class BSSheetServiceConfig(TypedDict, total=False):
    """Type definition for BS sheet service configuration."""

    reference_list_start_col: int
    reference_list_start_row: int
    reference_columns: list[str]
    region_reference: LinkedListConfig  # Optional for services that need region dropdown
    linked_lists: dict[str, LinkedListConfig]


BS_SHEET_CONFIG: dict[str, BSSheetServiceConfig] = {
    "inpatient": {
        "reference_list_start_col": 21,
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
            "trusts": {"start_col": 31, "pairs": [["Trust_Code", "Trust_Name"]]},
            "sites": {
                "start_col": 34,
                "pairs": [["Trust_Code", "Trust_Name"], ["Site_Code", "Site_Name"]],
            },
            "wards": {
                "start_col": 39,
                "pairs": [
                    ["Trust_Code", "Trust_Name"],
                    ["Site_Code", "Site_Name"],
                    ["Ward_Name"],
                ],
            },
        },
    },
    "ae": {
        "reference_list_start_col": 21,
        "reference_list_start_row": 2,
        "reference_columns": [
            "ICB_Code",
            "Trust_Code",
            "Trust_Name",
            "Site_Code",
            "Site_Name",
        ],
        "region_reference": {
            "start_col": 15,  # Column O
            "start_row": 2,
            "pairs": [["ICB_Code", "ICB_Name"]],
        },
        "linked_lists": {
            "regions": {
                "start_col": 18,
                "pairs": [["ICB_Code", "ICB_Name"]],
            },  # Column R-S
            "trusts": {
                "start_col": 27,
                "pairs": [["Trust_Code", "Trust_Name"]],
            },  # Column AA-AB
            "sites": {
                "start_col": 30,  # Column AD-AG
                "pairs": [["Trust_Code", "Trust_Name"], ["Site_Code", "Site_Name"]],
            },
        },
    },
    "ambulance": {
        "reference_list_start_col": 21,
        "reference_list_start_row": 2,
        "reference_columns": [
            "ICB_Code",
            "Trust_Code",
            "Trust_Name",
        ],
        "region_reference": {
            "start_col": 11,  # Column K (based on BS sheet analysis)
            "start_row": 2,
            "pairs": [["ICB_Code", "ICB_Name"]],
        },
        "linked_lists": {
            "regions": {
                "start_col": 14,  # Column N
                "pairs": [["ICB_Code", "ICB_Name"]],
            },
            "organisations": {"start_col": 31, "pairs": [["Trust_Code", "Trust_Name"]]},
        },
    },
}


class PeriodLabelCellConfig(TypedDict):
    """Type definition for period label cell configuration."""

    sheet: str
    cell: str
    template: str


def _period_cfg(label, cell="A2"):
    return {
        "notes_title": {
            "sheet": "Notes",
            "cell": cell,
            "template": f"{label} Friends and Family Test (FFT) Data - {{period}}",
        }
    }


PERIOD_LABEL_CONFIG: dict[str, dict[str, PeriodLabelCellConfig]] = {
    "inpatient": _period_cfg("Inpatient"),
    "ae": _period_cfg("A&E", "B2"),
    "ambulance": _period_cfg("Ambulance"),
}

_pct_base = {"ICB": [5, 6], "Trusts": [6, 7], "Sites": [8, 9]}
PERCENTAGE_COLUMN_CONFIG: dict[str, dict[str, list[int]]] = {
    "inpatient": {**_pct_base, "Wards": [9, 10]},
    "ae": _pct_base,
    "ambulance": {"PTS ICB": [5, 6], "PTS Org": [6, 7]},
}

# Percentage formatting configuration
PERCENTAGE_NUMBER_FORMAT = "0%"

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

SERVICE_TYPES = {"ip": "inpatient", "ae": "ae", "amb": "ambulance"}

TIME_SERIES_PREFIXES = {
    "inpatient": "Inpatient",
    "ae": "A&E",
    "ambulance": "Ambulance",
    "outpatient": "Outpatient",
    "maternity": "Q1",
    "community": "CH",
    "mental_health": "MH",
    "gp": "GP",
    "dental": "Dental",
    "post_covid": "Lcov Q1",
}

SUMMARY_COLUMNS = {
    "inpatient": {
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
    },
    "ae": {
        "orgs_submitting": {
            "total": " Submitted",
            "acute": " Acute Submitted",
            "wics": " WiCs & MIUs Submitted",
        },
        "responses": {
            "total": " Responses",
            "acute": " Acute Responses",
            "wics": " WiCs & MIUs Responses",
        },
        "positive": {
            "likely": " Likely",
            "extremely_likely": " Extremely Likely",
            "acute_likely": " Acute Likely",
            "acute_extremely_likely": " Acute Extremely Likely",
            "wics_likely": " WiCs & MIUs Likely",
            "wics_extremely_likely": " WiCs & MIUs Extremely Likely",
        },
        "negative": {
            "unlikely": " Unlikely",
            "extremely_unlikely": " Extremely Unlikely",
            "acute_unlikely": " Acute Unlikely",
            "acute_extremely_unlikely": " Acute Extremely Unlikely",
            "wics_unlikely": " WiCs & MIUs Unlikely",
            "wics_extremely_unlikely": " WiCs & MIUs Extremely Unlikely",
        },
    },
    "ambulance": {
        "orgs_submitting": {
            "total": " PTS Submitted",
            "nhs": " PTS NHS Submitted",
            "is": " PTS IS Submitted",
        },
        "responses": {
            "total": " PTS Responses",
            "nhs": " PTS NHS Responses",
            "is": " PTS IS Responses",
        },
        "positive": {
            "likely": " PTS Likely",
            "extremely_likely": " PTS Extremely Likely",
            "nhs_likely": " PTS NHS Likely",
            "nhs_extremely_likely": " PTS NHS Extremely Likely",
            "is_likely": " PTS IS Likely",
            "is_extremely_likely": " PTS IS Extremely Likely",
        },
        "negative": {
            "unlikely": " PTS Unlikely",
            "extremely_unlikely": " PTS Extremely Unlikely",
            "nhs_unlikely": " PTS NHS Unlikely",
            "nhs_extremely_unlikely": " PTS NHS Extremely Unlikely",
            "is_unlikely": " PTS IS Unlikely",
            "is_extremely_unlikely": " PTS IS Extremely Unlikely",
        },
    },
}

ENGLAND_ROWS_SKIP_COLUMNS = {"ICB": 2, "Trusts": 3, "Sites": 5, "Wards": 6, "PTS ICB": 2, "PTS Org": 3, "Mode Org": 3}

HEADER_ROW_RANGES_BY_SERVICE = {
    "inpatient": {s: [10, 14] for s in ("ICB", "Trusts", "Sites", "Wards")},
    "ae": {s: [3, 6] for s in ("ICB", "Trusts", "Sites")},
    "ambulance": {s: [1, 4] for s in ("PTS ICB", "PTS Org", "Mode Org")},
}

HEADER_ROWS_BY_SERVICE = {"inpatient": 11, "ae": 4, "ambulance": 3}

VALIDATION_CONFIG: dict[str, list[str]] = {
    "inpatient": ["ICB", "Trusts", "Sites", "Wards", "Summary"],
    "ae": ["ICB", "Trusts", "Sites", "Summary"],
    "ambulance": ["PTS ICB", "PTS Org", "Mode Org", "Summary"],
}

VALIDATION_KEY_COLUMNS: dict[str, str | list[str]] = {
    "ICB": "B",
    "Trusts": "B",
    "Sites": "D",
    "Wards": ["B", "D", "F"],
    "PTS ICB": "A",  # Use ICB Code (column A) instead of ICB Name (column B)
    "PTS Org": "B",
    "Mode Org": "B",
}

ENGLAND_TOTALS_DATA_SOURCE: dict[str, str] = {
    "Wards": "ward",
    "Sites": "site",
    "Trusts": "organisation",
    "ICB": "organisation",
    "PTS ICB": "organisation",
    "PTS Org": "organisation",
    "Mode Org": "organisation",
}

STANDARD_ENGLAND_DATA_COLUMNS = TOTALS_COLS + _STD_PCT_COLS + LIKERT_COLS

ENGLAND_ROWS_DATA_COLUMNS: dict[str, dict[str, list[str]]] = {
    "inpatient": {
        "ICB": STANDARD_ENGLAND_DATA_COLUMNS,
        "Trusts": STANDARD_ENGLAND_DATA_COLUMNS + MODE_COLS["inpatient"],
        "Sites": STANDARD_ENGLAND_DATA_COLUMNS,
        "Wards": STANDARD_ENGLAND_DATA_COLUMNS,
    },
    "ae": {
        "ICB": STANDARD_ENGLAND_DATA_COLUMNS,
        "Trusts": STANDARD_ENGLAND_DATA_COLUMNS + MODE_COLS["ae"],
        "Sites": STANDARD_ENGLAND_DATA_COLUMNS,
    },
    "ambulance": {
        "PTS ICB": STANDARD_ENGLAND_DATA_COLUMNS,
        "PTS Org": STANDARD_ENGLAND_DATA_COLUMNS,
        "Mode Org": MODE_COLS["ambulance"],
    },
}

# Summary sheet configuration for different service types
SUMMARY_SHEET_CONFIG = {
    "inpatient": {
        "rows": {"total": 8, "nhs": 9, "is": 10},
        "cols": {
            "orgs_submitting": 3,  # C
            "responses_to_date": 4,  # D
            "responses_current": 5,  # E
            "responses_previous": 6,  # F
            "pct_positive_current": 7,  # G
            "pct_positive_previous": 8,  # H
            "pct_negative_current": 9,  # I
            "pct_negative_previous": 10,  # J
        },
        "period_row": 7,
    },
    "ae": {
        "rows": {"total": 5, "acute": 6, "wics": 7},
        "cols": {
            "orgs_submitting": 3,  # C
            "responses_to_date": 4,  # D
            "responses_current": 5,  # E
            "responses_previous": 6,  # F
            "pct_positive_current": 7,  # G
            "pct_positive_previous": 8,  # H
            "pct_negative_current": 9,  # I
            "pct_negative_previous": 10,  # J
        },
        "period_row": 4,
    },
    "ambulance": {
        "rows": {"total": 8, "nhs": 9, "is": 10},
        "cols": {
            "orgs_submitting": 4,  # D
            "responses_to_date": 5,  # E
            "responses_current": 6,  # F
            "responses_previous": 7,  # G
            "pct_positive_current": 8,  # H
            "pct_positive_previous": 9,  # I
            "pct_negative_current": 10,  # J
            "pct_negative_previous": 11,  # K
        },
        "period_row": 7,
    },
}

_hdr_data = TOTALS_COLS + PERCENTAGE_COLS + ["Breakdown of Responses"]


def _expected_headers(svc):
    """Build expected header rows for a service type."""
    modes = MODE_COLS[svc]
    return {
        "ICB": {1: ICB_COLS + _hdr_data, 2: [""] * 6 + LIKERT_COLS, 3: [""] * 6 + modes},
        "Trusts": {
            1: ["ICB Code"] + TRUST_COLS + _hdr_data,
            2: [""] * 6 + LIKERT_COLS,
            3: [""] * 6 + modes,
        },
        "Sites": {
            1: ["ICB Code"] + TRUST_COLS + SITE_COLS + _hdr_data,
            2: [""] * 8 + LIKERT_COLS,
            3: [""] * 8 + modes,
        },
    }


EXPECTED_HEADERS: dict[str, dict[str, dict[int, list[str]]]] = {
    svc: _expected_headers(svc) for svc in ("inpatient", "ae")
}

_CRITICAL_COLS = ["A3", "B3", "C3", "D3", "E3", "F3", "G3"]
CRITICAL_HEADER_CELLS = {s: _CRITICAL_COLS for s in ("ICB", "Trusts", "Sites")}

_EXCLUDED_SHEETS = ["Notes", "BS"]
HEADER_VALIDATION_EXCLUDED_SHEETS = {
    s: _EXCLUDED_SHEETS for s in ("inpatient", "ae", "ambulance")
}

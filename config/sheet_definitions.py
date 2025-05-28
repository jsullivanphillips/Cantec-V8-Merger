from handlers import (
    sheet_20_1,
    # Add other sheet handler modules as needed
)

# List of sheet names that should exist in a valid V8 Excel file
REQUIRED_SHEETS = [
    "APPENDIX C-C1 FAS VER template",
    "ULC Coverpage",
    "32.5 Response Times",
    "32.6 Large Scale Network System",
    "32.11",
    "32.12",
    "32.13",
    "ULC Cover Page",
    "Deficiency Summary",
    "EXT only",
    "ELU only",
    "HOSES only",
    "20.1 | Report",
    "20.2 | Deficiencies",
    "20.3 | Recommendations",
    "21 | Documentation",
    "29",
    "30",
    "31 Documentation (2)",
    "22.1 | CU or Transp Insp",
    "32 ControlUnit|Transponder (2)",
    "22.2 | CU or Transp Test",
    "22.3 + 22.4 | Voice & PS",
    "32.7",
    "32.8 Power Supply (2)",
    "22.5 | Power Supply(s)",
    "22.6 | Annunciator(s)",
    "22.7 | Annun & Seq Disp",
    "22.9 + 22.10 | Printer",
    "23.1 Field Device Legend",
    "23.2 Device Record",
    "23.3 CircuitFaultTolerance",
]


# Maps sheet names to their merge handler functions
SHEET_MERGE_HANDLERS = {
    "20.1 | Report": sheet_20_1.merge_20_1_report,
    # Add more mappings as you implement them
}


def get_merge_handler(sheet_name):
    return SHEET_MERGE_HANDLERS.get(sheet_name)

from handlers.handler_base import merge_cells, merge_checkbox_groups

# Placeholder cells to merge (you can update these later)
MERGE_CELLS = [
    "H25",
    "C26",
    "A9",
    "A10",
    "A11",
    "A12",
    "A13",
    "A14",
    "A15",
    "A16",
    "A17",
    "A18",
    "A19",
    "A20",
    "A21",
    "A22",
    "A23",
    "A24",
]

# Placeholder checkbox groups (YES/NO or YES/NO/N/A)
MERGE_CHECKBOX_GROUPS = [
    ["M9", "P9"],
    ["M10", "P10"],
    ["M11", "P11"],
    ["M12", "P12"],
    ["M13", "P13"],
    ["M14", "P14"],
    ["L15", "N15", "P15"],
    ["L16", "N16", "P16"],
    ["M17", "P17"],
    ["M18", "P18"],
    ["M19", "P19"],
    ["M20", "P20"],
    ["M21", "P21"],
    ["L22", "N22", "P22"],
    ["M23", "P23"],
    ["L24", "N24", "P24"],
]


def merge_21_documentation(ws_file_list, output_ws):
    """
    Merges the 21 | Documentation sheet from multiple technician workbooks.
    Handles cell-by-cell conflicts and exclusive checkbox group conflicts.
    Technician names are tagged in column 'S'.
    """
    merge_cells(ws_file_list, output_ws, MERGE_CELLS, tech_col_letter="S")
    merge_checkbox_groups(
        ws_file_list, output_ws, MERGE_CHECKBOX_GROUPS, tech_col_letter="S"
    )

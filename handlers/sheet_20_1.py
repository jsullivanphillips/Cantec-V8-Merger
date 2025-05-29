from handlers.handler_base import (
    merge_cells,
    merge_checkbox_groups,
)


# List of cells/ranges to merge for this sheet
# These will be copied cell-by-cell, with matching output locations
MERGE_CELLS = ["F7", "F8", "F9", "F10", "F11", "F12", "D13", "D14", "l25", "K28", "K35"]

MERGE_CHECKBOX_GROUPS = [
    ["D15", "G15"],
    ["N17", "R17"],
    ["N22", "R22"],
    ["N23", "R23"],
    ["N24", "R24"],
    ["N26", "R26"],
    ["M34", "Q34"],
]


def merge_20_1_report(ws_file_list, output_ws):
    """
    ws_file_list: List of tuples (worksheet, filename)
    output_ws: xlwings sheet where data will be merged
    """
    merge_cells(ws_file_list, output_ws, MERGE_CELLS, tech_col_letter="T")
    merge_checkbox_groups(
        ws_file_list, output_ws, MERGE_CHECKBOX_GROUPS, tech_col_letter="T"
    )

from handlers.handler_base import (
    merge_cells,
)


# List of cells/ranges to merge for this sheet
# These will be copied cell-by-cell, with matching output locations
MERGE_CELLS = [
    "L8",
    "N8",
    "L9",
    "N9",
    "L10",
    "N10",
    "F13",
    "H15",
    "L19",
    "N19",
    "L20",
    "N19",
    "L21",
    "N21",
    "L22",
    "N22",
    "L23",
    "N23",
    "L24",
    "N24",
    "L25",
    "N25",
    "L26",
    "N26",
    "L27",
    "N27",
    "L28",
    "N28",
    "L29",
    "N29",
    "L30",
    "N30",
    "L31",
    "N31",
    "L32",
    "N32",
    "L33",
    "N33",
    "L34",
    "N34",
    "L35",
    "N35",
    "L36",
    "N36",
    "L37",
    "N37",
    "L38",
    "N38",
]


def merge_23_1_field_device(ws_file_list, output_ws):
    """
    ws_file_list: List of tuples (worksheet, filename)
    output_ws: xlwings sheet where data will be merged
    """
    merge_cells(ws_file_list, output_ws, MERGE_CELLS, tech_col_letter="Q")

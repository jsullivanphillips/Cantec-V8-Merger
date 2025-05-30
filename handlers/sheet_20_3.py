from handlers.handler_base import merge_cells

# Cells A6 to A13 need to be merged and checked for conflicts
MERGE_CELLS = ["A6", "A7", "A8", "A9", "A10", "A11", "A12", "A13"]


def merge_20_3_recommendations(ws_file_list, output_ws):
    """
    Merges the 20.3 | Recommendations sheet from multiple technician workbooks.
    Performs cell-by-cell merging with conflict detection and technician tagging.
    """
    merge_cells(ws_file_list, output_ws, MERGE_CELLS, tech_col_letter="O")

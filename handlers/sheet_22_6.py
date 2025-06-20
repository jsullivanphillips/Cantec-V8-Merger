from handlers.handler_base import merge_cells, merge_checkbox_groups, is_page_meaningful

MERGE_CELLS = [
    # Page 1
    "F6",
    *[f"A{row}" for row in range(43, 46)],
]

MERGE_CHECKBOX_GROUPS = [
    # Page 1
    ["L15", "N15", "P15"],
]

MERGE_CELLS_BY_PAGE = [
    MERGE_CELLS[0:66],  # Page 1 (F6 to A45)
    MERGE_CELLS[66:126],  # Page 2 (F51 to A90)
    MERGE_CELLS[126:186],  # Page 3 (F96 to A135)
    MERGE_CELLS[186:246],  # Page 4 (F141 to A180)
    MERGE_CELLS[246:],  # Page 5 (F187 to A226)
]

MERGE_CHECKBOX_GROUPS_BY_PAGE = [
    MERGE_CHECKBOX_GROUPS[0:16],  # Page 1
    MERGE_CHECKBOX_GROUPS[16:32],  # Page 2
    MERGE_CHECKBOX_GROUPS[32:48],  # Page 3
    MERGE_CHECKBOX_GROUPS[48:64],  # Page 4
    MERGE_CHECKBOX_GROUPS[64:],  # Page 5
]

# --- Anchor cells used to detect meaningful pages ---
MEANINGFUL_ANCHORS = [
    ["K37", "M17"],  # Page 1
    ["M62", "K85"],  # Page 2
    ["M109", "K132"],  # Page 3
    ["M155", "K178"],  # Page 4
    ["M201", "K224"],  # Page 5
]

# Special row is for thinks that require a recorded value (i.e. 27.7 V dc)
# We check if there is a "meaningful value" where the value should be
# recorded. Then it checks if anything is highlighted on that row,
# signifying that this is a new value to be saved in the report.
SPECIAL_ROW_RANGES = [
    # Page 1
    {"rows": range(17, 23), "value_col": "M", "highlight_cols": ["A", "M"]},
    {"rows": range(37, 39), "value_col": "K", "highlight_cols": ["A", "K"]},
    {"rows": range(40, 41), "value_col": "K", "highlight_cols": ["A", "K"]},
    # Page 2
    {"rows": range(62, 68), "value_col": "M", "highlight_cols": ["A", "M"]},
    {"rows": range(82, 84), "value_col": "K", "highlight_cols": ["A", "K"]},
    {"rows": range(85, 86), "value_col": "K", "highlight_cols": ["A", "K"]},
    # Page 3
    {"rows": range(107, 113), "value_col": "M", "highlight_cols": ["A", "M"]},
    {"rows": range(127, 129), "value_col": "K", "highlight_cols": ["A", "K"]},
    {"rows": range(130, 131), "value_col": "K", "highlight_cols": ["A", "K"]},
]


def merge_22_6_annun(ws_file_list, output_ws):
    for i, (merge_cells_page, checkbox_groups_page) in enumerate(
        zip(MERGE_CELLS_BY_PAGE, MERGE_CHECKBOX_GROUPS_BY_PAGE)
    ):
        anchor_cells = MEANINGFUL_ANCHORS[i]
        if not is_page_meaningful(ws_file_list, anchor_cells):
            print(f"Page {i + 1} is blank — skipping this and all following pages.")
            break

        merge_cells(
            ws_file_list,
            output_ws,
            merge_cells_page,
            tech_col_letter="R",
            special_row_ranges=SPECIAL_ROW_RANGES,
        )
        merge_checkbox_groups(
            ws_file_list, output_ws, checkbox_groups_page, tech_col_letter="R"
        )

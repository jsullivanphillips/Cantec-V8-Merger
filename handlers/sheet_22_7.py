from handlers.handler_base import merge_cells, merge_checkbox_groups, is_page_meaningful


MERGE_CELLS = [
    # Page 1
    "J6" "G8",
    "G9",
    *[f"A{row}" for row in range(10, 23)],
    "J29",
    "G31",
    "G32"
    # Page 2
    "G38",
    "G39",
    *[f"A{row}" for row in range(40, 53)],
    # Page 3
    "G78",
    "G79",
    *[f"A{row}" for row in range(80, 93)],
    # Page 4
    "G118",
    "G119",
    *[f"A{row}" for row in range(120, 133)],
]

MERGE_CHECKBOX_GROUPS = [
    # Page 1
    ["L9", "N9", "P9"],
    ["L10", "N10", "P10"],
    ["L11", "N11", "P11"],
    ["L12", "N12", "P12"],
    ["L13", "N13", "P13"],
    ["L14", "N14", "P14"],
    ["L15", "N15", "P15"],
    ["L16", "N16", "P16"],
    ["L17", "N17", "P17"],
    ["L18", "N18", "P18"],
    ["L19", "N19", "P19"],
    ["L20", "N20", "P20"],
    ["L21", "N21", "P21"],
    ["L22", "N22", "P22"],
    # Page 2
    ["L40", "N40", "P40"],
    ["L41", "N41", "P41"],
    ["L42", "N42", "P42"],
    ["L43", "N43", "P43"],
    ["L44", "N44", "P44"],
    ["L45", "N45", "P45"],
    ["L46", "N46", "P46"],
    ["L47", "N47", "P47"],
    ["L48", "N48", "P48"],
    ["L49", "N49", "P49"],
    ["L50", "N50", "P50"],
    ["L51", "N51", "P51"],
    ["L52", "N52", "P52"],
    ["L53", "N53", "P53"],
    # Page 3
    ["L80", "N80", "P80"],
    ["L81", "N81", "P81"],
    ["L82", "N82", "P82"],
    ["L83", "N83", "P83"],
    ["L84", "N84", "P84"],
    ["L85", "N85", "P85"],
    ["L86", "N86", "P86"],
    ["L87", "N87", "P87"],
    ["L88", "N88", "P88"],
    ["L89", "N89", "P89"],
    ["L90", "N90", "P90"],
    ["L91", "N91", "P91"],
    ["L92", "N92", "P92"],
    ["L93", "N93", "P93"],
    # Page 4
    ["L120", "N120", "P120"],
    ["L121", "N121", "P121"],
    ["L122", "N122", "P122"],
    ["L123", "N123", "P123"],
    ["L124", "N124", "P124"],
    ["L125", "N125", "P125"],
    ["L126", "N126", "P126"],
    ["L127", "N127", "P127"],
    ["L128", "N128", "P128"],
    ["L129", "N129", "P129"],
    ["L130", "N130", "P130"],
    ["L131", "N131", "P131"],
    ["L132", "N132", "P132"],
    ["L133", "N133", "P133"],
]

# split MERGE_CELLS into four pages of 15 entries each (2 G-rows + 13 A-rows)
MERGE_CELLS_BY_PAGE = [
    MERGE_CELLS[0:15],  # Page 1: G7, G8, A9–A21
    MERGE_CELLS[15:30],  # Page 2: G38, G39, A40–A52
    MERGE_CELLS[30:45],  # Page 3: G78, G79, A80–A92
    MERGE_CELLS[45:60],  # Page 4: G118, G119, A120–A132
]

# split MERGE_CHECKBOX_GROUPS into four pages of 14 groups each
MERGE_CHECKBOX_GROUPS_BY_PAGE = [
    MERGE_CHECKBOX_GROUPS[0:14],  # Page 1: rows 9–22
    MERGE_CHECKBOX_GROUPS[14:28],  # Page 2: rows 40–53
    MERGE_CHECKBOX_GROUPS[28:42],  # Page 3: rows 80–93
    MERGE_CHECKBOX_GROUPS[42:56],  # Page 4: rows 120–133
]

# --- Anchor cells used to detect meaningful pages ---
MEANINGFUL_ANCHORS = [
    ["G7", "G8", "A9"],  # Page 1
    ["G38", "G39", "A40"],  # Page 2
    ["G78", "G79", "A80"],  # Page 3
    ["G118", "G119", "A120"],  # Page 4
]


def merge_22_7_seq(ws_file_list, output_ws):
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
        )
        merge_checkbox_groups(
            ws_file_list, output_ws, checkbox_groups_page, tech_col_letter="R"
        )

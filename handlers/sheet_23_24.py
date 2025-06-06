from handlers.handler_base import merge_cells, merge_checkbox_groups, is_page_meaningful

MERGE_CELLS = [
    # Page 1
    "F5",
    *[f"A{row}" for row in range(7, 24)],
    # PS on Page 1
    "F29",
    "F30",
    "G31",
    "G32",
    *[f"A{row}" for row in range(33, 41)],
    # Page 2
    "F42",
    "F43",
    "G44",
    "G45",
    *[f"A{row}" for row in range(46, 54)],
    # Page 3
    "F55",
    "F56",
    "G57",
    "G58",
    *[f"A{row}" for row in range(59, 67)],
    # Page 4
    "F68",
    "F69",
    "G70",
    "G71",
    *[f"A{row}" for row in range(72, 80)],
]

MERGE_CHECKBOX_GROUPS = [
    # Page 1
    ["L7", "N7", "P7"],
    ["L8", "N8", "P8"],
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
    ["L23", "N23", "P23"],
    # PS on Page 1
    ["L33", "N33", "P33"],
    ["L34", "N34", "P34"],
    ["L35", "N35", "P35"],
    ["L36", "N36", "P36"],
    ["L37", "N37", "P37"],
    ["L38", "N38", "P38"],
    ["L39", "N39", "P39"],
    ["L40", "N40", "P40"],
    # Page 2
    ["L46", "N46", "P46"],
    ["L47", "N47", "P47"],
    ["L48", "N48", "P48"],
    ["L49", "N49", "P49"],
    ["L50", "N50", "P50"],
    ["L51", "N51", "P51"],
    ["L52", "N52", "P52"],
    ["L53", "N53", "P53"],
    # Page 3
    ["L59", "N59", "P59"],
    ["L60", "N60", "P60"],
    ["L61", "N61", "P61"],
    ["L62", "N62", "P62"],
    ["L63", "N63", "P63"],
    ["L64", "N64", "P64"],
    ["L65", "N65", "P65"],
    ["L66", "N66", "P66"],
    # Page 4
    ["L68", "N68", "P68"],
    ["L69", "N69", "P69"],
    ["L70", "N70", "P70"],
    ["L71", "N71", "P71"],
    ["L72", "N72", "P72"],
    ["L73", "N73", "P73"],
    ["L74", "N74", "P74"],
    ["L75", "N75", "P75"],
]

MERGE_CELLS_BY_PAGE = [
    MERGE_CELLS[0 : 1 + 17 + 2 + 8],  # Page 1: F5 + A7–A23 + F29,F30,G31,G32 + A33–A40
    MERGE_CELLS[28 : 28 + 2 + 8],  # Page 2: F42,F43,G44,G45 + A46–A53
    MERGE_CELLS[38 : 38 + 2 + 8],  # Page 3: F55,F56,G57,G58 + A59–A66
    MERGE_CELLS[48 : 48 + 2 + 8],  # Page 4: F68,F69,G70,G71 + A72–A79
]


MERGE_CHECKBOX_GROUPS_BY_PAGE = [
    MERGE_CHECKBOX_GROUPS[0 : 17 + 8],  # Page 1: L7–L23 + L33–L40
    MERGE_CHECKBOX_GROUPS[25 : 25 + 8],  # Page 2: L46–L53
    MERGE_CHECKBOX_GROUPS[33 : 33 + 8],  # Page 3: L59–L66
    MERGE_CHECKBOX_GROUPS[41 : 41 + 8],  # Page 4: L68–L75
]

# Use two meaningful anchor cells per page (like a known header or unique field)
MEANINGFUL_ANCHORS = [
    ["F29", "F30"],  # Page 1
    ["F42", "F43"],  # Page 2
    ["F55", "F56"],  # Page 3
    ["F68", "F69"],  # Page 4
]


def merge_23_24_Voice_PS(ws_file_list, output_ws):
    for i, (merge_cells_page, checkbox_groups_page) in enumerate(
        zip(MERGE_CELLS_BY_PAGE, MERGE_CHECKBOX_GROUPS_BY_PAGE)
    ):
        anchor_cells = MEANINGFUL_ANCHORS[i]
        if not is_page_meaningful(ws_file_list, anchor_cells):
            print(f"Page {i + 1} is blank — skipping this and all following pages.")
            break

        merge_cells(ws_file_list, output_ws, merge_cells_page, tech_col_letter="R")
        merge_checkbox_groups(
            ws_file_list, output_ws, checkbox_groups_page, tech_col_letter="R"
        )

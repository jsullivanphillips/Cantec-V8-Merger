from handlers.handler_base import merge_cells, merge_checkbox_groups, is_page_meaningful


MERGE_CELLS = [
    # Page 1
    "J6",
    "G8",
    "G9",
    *[f"A{row}" for row in range(10, 23)],
    "J29",
    "G31",
    "G32",
    *[f"A{row}" for row in range(33, 37)],
    # Page 2
    "G39",  # 8  +31
    "G40",  # 9  +31
    *[f"A{row}" for row in range(41, 54)],  # 10→22 +31
    "G56",  # 31 +31
    "G57",  # 32 +31
    *[f"A{row}" for row in range(58, 62)],  # 33→36 +31
    # Page 3
    "G64",  # 8  +56
    "G65",  # 9  +56
    *[f"A{row}" for row in range(66, 79)],  # 10→22 +56
    "G81",  # 31 +56
    "G82",  # 32 +56
    *[f"A{row}" for row in range(83, 97)],  # 33→36 +56
]

# --- Anchor cells used to detect meaningful pages ---
MEANINGFUL_ANCHORS = [
    ["G8", "G9", "J6", "G31", "G32"],  # Page 1
    ["G39", "G40", "G56", "G57"],  # Page 2
    ["G64", "G65", "G81", "G82"],  # Page 3
]

MERGE_CHECKBOX_GROUPS = [
    # Page 1
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
    ["L33", "N33", "P33"],
    ["L34", "N34", "P34"],
    ["L35", "N35", "P35"],
    ["L36", "N36", "P36"],
    # Page 2
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
    ["L58", "N58", "P58"],
    ["L59", "N59", "P59"],
    ["L60", "N60", "P60"],
    ["L61", "N61", "P61"],
    # Page 3
    ["L66", "N66", "P66"],
    ["L67", "N67", "P67"],
    ["L68", "N68", "P68"],
    ["L69", "N69", "P69"],
    ["L70", "N70", "P70"],
    ["L71", "N71", "P71"],
    ["L72", "N72", "P72"],
    ["L73", "N73", "P73"],
    ["L74", "N74", "P74"],
    ["L75", "N75", "P75"],
    ["L76", "N76", "P76"],
    ["L77", "N77", "P77"],
    ["L78", "N78", "P78"],
    ["L83", "N83", "P83"],
    ["L84", "N84", "P84"],
    ["L85", "N85", "P85"],
    ["L86", "N86", "P86"],
]

# split MERGE_CELLS into four pages of 15 entries each (2 G-rows + 13 A-rows)
MERGE_CELLS_BY_PAGE = [
    MERGE_CELLS[0:23],  # Page 1 (23 entries)
    MERGE_CELLS[23:44],  # Page 2 (21 entries)
    MERGE_CELLS[44:],  # Page 3 (remaining entries)
]

MERGE_CHECKBOX_GROUPS_BY_PAGE = [
    # Page 1: rows 10–22 & 33–36 (17 groups)
    MERGE_CHECKBOX_GROUPS[0:17],
    # Page 2: rows 41–53 & 58–61 (17 groups)
    MERGE_CHECKBOX_GROUPS[17:34],
    # Page 3: rows 66–78 & 83–86 (17 groups)
    MERGE_CHECKBOX_GROUPS[34:51],
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

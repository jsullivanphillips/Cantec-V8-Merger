from handlers.handler_base import merge_cells, merge_checkbox_groups, is_page_meaningful

MERGE_CELLS = [
    # Page 1
    "F6",
    "F7",
    "F8",
    "G9",
    "C10",
    "H10",
    "J10",
    "N10",
    "H11",
    "J11",
    "M11",
    "P11",
    "E12",
    "G12",
    "I12",
    "M17",
    "M18",
    "M19",
    "M20",
    "M21",
    "M22",
    "M30",
    "K37",
    "K38",
    "K40",
    *[f"A{row}" for row in range(15, 41)],
    *[f"A{row}" for row in range(43, 46)],
    # Page 2 (values increased by 45)
    "F51",
    "F52",
    "F53",
    "G54",
    "C55",
    "H55",
    "J55",
    "N55",
    "H56",
    "J56",
    "M56",
    "P56",
    "E57",
    "G57",
    "I57",
    "M62",
    "M63",
    "M64",
    "M65",
    "M66",
    "M67",
    "M75",
    "K82",
    "K83",
    "K85",
    "A60",
    "A61",
    "A62",
    "A63",
    "A64",
    "A65",
    "A66",
    "A67",
    "A68",
    "A69",
    "A70",
    "A71",
    "A72",
    "A73",
    "A74",
    "A75",
    "A76",
    "A77",
    "A78",
    "A79",
    "A80",
    "A81",
    "A82",
    "A83",
    "A84",
    "A85",
    "A88",
    "A89",
    "A90",
    # Page 3 (values increased by 90)
    "F96",
    "F97",
    "F98",
    "G99",
    "C100",
    "H100",
    "J100",
    "N100",
    "H101",
    "J101",
    "M101",
    "P101",
    "E102",
    "G102",
    "I102",
    "M107",
    "M108",
    "M109",
    "M110",
    "M111",
    "M112",
    "M120",
    "K127",
    "K128",
    "K130",
    "A105",
    "A106",
    "A107",
    "A108",
    "A109",
    "A110",
    "A111",
    "A112",
    "A113",
    "A114",
    "A115",
    "A116",
    "A117",
    "A118",
    "A119",
    "A120",
    "A121",
    "A122",
    "A123",
    "A124",
    "A125",
    "A126",
    "A127",
    "A128",
    "A129",
    "A130",
    "A133",
    "A134",
    "A135",
    # Page 4 (values increased by 135)
    "F141",
    "F142",
    "F143",
    "G144",
    "C145",
    "H145",
    "J145",
    "N145",
    "H146",
    "J146",
    "M146",
    "P146",
    "E147",
    "G147",
    "I147",
    "M152",
    "M153",
    "M154",
    "M155",
    "M156",
    "M157",
    "M165",
    "K172",
    "K173",
    "K175",
    "A150",
    "A151",
    "A152",
    "A153",
    "A154",
    "A155",
    "A156",
    "A157",
    "A158",
    "A159",
    "A160",
    "A161",
    "A162",
    "A163",
    "A164",
    "A165",
    "A166",
    "A167",
    "A168",
    "A169",
    "A170",
    "A171",
    "A172",
    "A173",
    "A174",
    "A175",
    "A178",
    "A179",
    "A180",
    # Page 5
    "F187",
    "F188",
    "F189",
    "G190",
    "C191",
    "H191",
    "J191",
    "N191",
    "H192",
    "J192",
    "M192",
    "P192",
    "E193",
    "G193",
    "I193",
    "M198",
    "M199",
    "M200",
    "M201",
    "M202",
    "M203",
    "M211",
    "K218",
    "K219",
    "K221",
    "A196",
    "A197",
    "A198",
    "A199",
    "A200",
    "A201",
    "A202",
    "A203",
    "A204",
    "A205",
    "A206",
    "A207",
    "A208",
    "A209",
    "A210",
    "A211",
    "A212",
    "A213",
    "A214",
    "A215",
    "A216",
    "A217",
    "A218",
    "A219",
    "A220",
    "A221",
    "A224",
    "A225",
    "A226",
]

MERGE_CHECKBOX_GROUPS = [
    # Page 1
    ["L15", "N15", "P15"],
    ["L16", "N16", "P16"],
    ["L23", "N23", "P23"],
    ["L24", "N24", "P24"],
    ["L25", "N25", "P25"],
    ["L26", "N26", "P26"],
    ["L27", "N27", "P27"],
    ["L28", "N28", "P28"],
    ["L29", "N29", "P29"],
    ["L31", "N31", "P31"],
    ["L33", "N33", "P33"],
    ["L34", "N34", "P34"],
    ["L35", "N35", "P35"],
    ["L43", "N43", "P43"],
    ["L44", "N44", "P44"],
    ["L45", "N45", "P45"],
    # Page 2
    ["L60", "N60", "P60"],
    ["L61", "N61", "P61"],
    ["L68", "N68", "P68"],
    ["L69", "N69", "P69"],
    ["L70", "N70", "P70"],
    ["L71", "N71", "P71"],
    ["L72", "N72", "P72"],
    ["L73", "N73", "P73"],
    ["L74", "N74", "P74"],
    ["L76", "N76", "P76"],
    ["L78", "N78", "P78"],
    ["L79", "N79", "P79"],
    ["L80", "N80", "P80"],
    ["L88", "N88", "P88"],
    ["L89", "N89", "P89"],
    ["L90", "N90", "P90"],
    # Page 3
    ["L105", "N105", "P105"],
    ["L106", "N106", "P106"],
    ["L113", "N113", "P113"],
    ["L114", "N114", "P114"],
    ["L115", "N115", "P115"],
    ["L116", "N116", "P116"],
    ["L117", "N117", "P117"],
    ["L118", "N118", "P118"],
    ["L119", "N119", "P119"],
    ["L121", "N121", "P121"],
    ["L123", "N123", "P123"],
    ["L124", "N124", "P124"],
    ["L125", "N125", "P125"],
    ["L133", "N133", "P133"],
    ["L134", "N134", "P134"],
    ["L135", "N135", "P135"],
    # Page 4
    ["L150", "N150", "P150"],
    ["L151", "N151", "P151"],
    ["L158", "N158", "P158"],
    ["L159", "N159", "P159"],
    ["L160", "N160", "P160"],
    ["L161", "N161", "P161"],
    ["L162", "N162", "P162"],
    ["L163", "N163", "P163"],
    ["L164", "N164", "P164"],
    ["L166", "N166", "P166"],
    ["L168", "N168", "P168"],
    ["L169", "N169", "P169"],
    ["L170", "N170", "P170"],
    ["L178", "N178", "P178"],
    ["L179", "N179", "P179"],
    ["L180", "N180", "P180"],
    # Page 5
    ["L196", "N196", "P196"],
    ["L197", "N197", "P197"],
    ["L204", "N204", "P204"],
    ["L205", "N205", "P205"],
    ["L206", "N206", "P206"],
    ["L207", "N207", "P207"],
    ["L208", "N208", "P208"],
    ["L209", "N209", "P209"],
    ["L210", "N210", "P210"],
    ["L212", "N212", "P212"],
    ["L214", "N214", "P214"],
    ["L215", "N215", "P215"],
    ["L216", "N216", "P216"],
    ["L224", "N224", "P224"],
    ["L225", "N225", "P225"],
    ["L226", "N226", "P226"],
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


def merge_22_5_PS(ws_file_list, output_ws):
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

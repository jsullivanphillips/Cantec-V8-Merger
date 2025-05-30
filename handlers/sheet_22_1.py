from handlers.handler_base import merge_cells, merge_checkbox_groups

# Placeholder cells to merge (you can update these later)
MERGE_CELLS = [
    # Page 1
    "H15",
    "H16",  # Location and ID
    "A17",
    "A18",
    "A19",
    "A20",
    "A21",
    "Q22",
    "Q23",
    "T23",  # Firmware
    "Q24",
    "Q25",
    "T25",  # Software
    "A26",
    "A27",
    "A28",
    "A29",
    # Page 2
    "H58",
    "H59",  # Location and ID
    "A60",
    "A61",
    "A62",
    "A63",
    "A64",
    "Q65",
    "Q66",
    "T66",  # Firmware
    "Q67",
    "Q68",
    "T68",  # Software
    "A69",
    "A70",
    "A71",
    "A72",
    # Page 3
    "H104",
    "H105",  # Location and ID
    "A106",
    "A107",
    "A108",
    "A109",
    "A110",
    "Q111",
    "Q112",
    "T112",  # Firmware
    "Q113",
    "Q114",
    "T114",  # Software
    "A115",
    "A116",
    "A117",
    "A118",
    # Page 4
    "H149",
    "H150",  # Location and ID
    "A151",
    "A152",
    "A153",
    "A154",
    "A155",
    "Q156",
    "Q157",
    "T157",  # Firmware
    "Q158",
    "Q159",
    "T159",  # Software
    "A160",
    "A161",
    "A162",
    "A163",
    # Page 5
    "H194",
    "H195",  # Location and ID
    "A196",
    "A197",
    "A198",
    "A199",
    "A200",
    "Q201",
    "Q202",
    "T202",  # Firmware
    "Q203",
    "Q204",
    "T204",  # Software
    "A205",
    "A206",
    "A207",
    "A208",
]

# Placeholder checkbox groups (YES/NO or YES/NO/N/A)
MERGE_CHECKBOX_GROUPS = [
    # Page 1
    ["Q17", "S17", "U17"],
    ["Q18", "S18", "U18"],
    ["Q19", "S19", "U19"],
    ["Q20", "S20", "U20"],
    ["Q21", "S21", "U21"],
    ["Q26", "S26", "U26"],
    ["Q27", "S27", "U27"],
    ["Q28", "S28", "U28"],
    ["Q29", "S29", "U29"],
    # Page 2
    ["Q60", "S60", "U60"],
    ["Q61", "S61", "U61"],
    ["Q62", "S62", "U62"],
    ["Q63", "S63", "U63"],
    ["Q64", "S64", "U64"],
    ["Q69", "S69", "U69"],
    ["Q70", "S70", "U70"],
    ["Q71", "S71", "U71"],
    ["Q72", "S72", "U72"],
    # Page 3
    ["Q106", "S106", "U106"],
    ["Q107", "S107", "U107"],
    ["Q108", "S108", "U108"],
    ["Q109", "S109", "U109"],
    ["Q110", "S110", "U110"],
    ["Q115", "S115", "U115"],
    ["Q116", "S116", "U116"],
    ["Q117", "S117", "U117"],
    ["Q118", "S118", "U118"],
    # Page 4
    ["Q151", "S151", "U151"],
    ["Q152", "S152", "U152"],
    ["Q153", "S153", "U153"],
    ["Q154", "S154", "U154"],
    ["Q155", "S155", "U155"],
    ["Q160", "S160", "U160"],
    ["Q161", "S161", "U161"],
    ["Q162", "S162", "U162"],
    ["Q163", "S163", "U163"],
    # Page 5
    ["Q196", "S196", "U196"],
    ["Q197", "S197", "U197"],
    ["Q198", "S198", "U198"],
    ["Q199", "S199", "U199"],
    ["Q200", "S200", "U200"],
    ["Q205", "S205", "U205"],
    ["Q206", "S206", "U206"],
    ["Q207", "S207", "U207"],
    ["Q208", "S208", "U208"],
]


def merge_22_1_CU(ws_file_list, output_ws):
    """
    Merges the 22.1 | CU or Transp Insp sheet from multiple technician workbooks.
    Handles cell-by-cell conflicts and exclusive checkbox group conflicts.
    Technician names are tagged in column 'W'.
    """
    merge_cells(ws_file_list, output_ws, MERGE_CELLS, tech_col_letter="W")
    merge_checkbox_groups(
        ws_file_list, output_ws, MERGE_CHECKBOX_GROUPS, tech_col_letter="W"
    )

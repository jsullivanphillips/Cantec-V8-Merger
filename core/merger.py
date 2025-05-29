import xlwings as xw
from config.sheet_definitions import get_merge_handler


def merge(file_paths, output_path):
    output_wb = xw.Book(output_path)
    input_wbs = [xw.Book(path) for path in file_paths]

    try:
        # Collect all unique sheet names
        all_sheet_names = set()
        for wb in input_wbs:
            all_sheet_names.update(sheet.name for sheet in wb.sheets)

        # Merge using handlers
        for sheet_name in sorted(all_sheet_names):
            handler = get_merge_handler(sheet_name)
            if handler:
                print(f"Merging sheet: {sheet_name}")
                ws_file_list = [
                    (wb.sheets[sheet_name], wb.name)
                    for wb in input_wbs
                    if sheet_name in [s.name for s in wb.sheets]
                ]

                # Reuse or create target sheet in output
                if sheet_name in [s.name for s in output_wb.sheets]:
                    output_ws = output_wb.sheets[sheet_name]
                else:
                    output_ws = output_wb.sheets.add(name=sheet_name)

                handler(ws_file_list, output_ws)

        output_wb.save()

    finally:
        output_wb.close()
        for wb in input_wbs:
            wb.close()

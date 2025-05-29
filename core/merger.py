import xlwings as xw
from config.sheet_definitions import get_merge_handler
from handlers.handler_base import merge_conflict_log
import os


def merge(file_paths, output_path, progress_callback=None):
    output_wb = xw.Book(output_path)
    input_wbs = []
    file_open_steps = len(file_paths)

    if progress_callback:
        progress_callback(0, "Starting merge...")

    # Open files (first 50%)
    for idx, path in enumerate(file_paths):
        if progress_callback:
            pct = ((idx + 1) / file_open_steps) * 50
            progress_callback(pct, f"Opening file: {os.path.basename(path)}")
        wb = xw.Book(path)
        input_wbs.append(wb)

    # Get all handler-eligible sheets
    all_sheet_names = set()
    for wb in input_wbs:
        all_sheet_names.update(sheet.name for sheet in wb.sheets)

    handler_sheet_names = sorted(
        [name for name in all_sheet_names if get_merge_handler(name)]
    )
    sheet_merge_steps = len(handler_sheet_names)

    # Merge sheets (next 40%)
    for idx, sheet_name in enumerate(handler_sheet_names):
        handler = get_merge_handler(sheet_name)
        if handler:
            step_pct = ((idx + 1) / sheet_merge_steps) * 40
            pct = 50 + step_pct  # 50–90%
            if progress_callback:
                progress_callback(pct, f"Merging sheet: {sheet_name}")

            ws_file_list = [
                (wb.sheets[sheet_name], wb.name)
                for wb in input_wbs
                if sheet_name in [s.name for s in wb.sheets]
            ]

            output_ws = (
                output_wb.sheets[sheet_name]
                if sheet_name in [s.name for s in output_wb.sheets]
                else output_wb.sheets.add(name=sheet_name)
            )

            handler(ws_file_list, output_ws)

    # Saving and closing (last 10%)
    if progress_callback:
        progress_callback(95, "Saving merged workbook...")

    output_wb.save()

    if progress_callback:
        progress_callback(100, "Finalizing...")

    output_wb.close()
    for wb in input_wbs:
        wb.close()

    return merge_conflict_log

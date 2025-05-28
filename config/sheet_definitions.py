from handlers.sheet_20_1 import merge_20_1_report

SHEET_HANDLERS = {
    "20.1 | Report": merge_20_1_report,
}


def get_merge_handler(sheet_name):
    return SHEET_HANDLERS.get(sheet_name)

def merge_20_1_report(ws_list, output_ws):
    # Placeholder: implement your cell-copying logic here
    output_ws["A1"] = "Merged data from 20.1 | Report"
    row_offset = 2

    for ws in ws_list:
        # Hardcoded example: copy A2:D5 from each input ws
        for row in ws.iter_rows(min_row=2, max_row=5, min_col=1, max_col=4):
            for cell in row:
                output_ws.cell(row=row_offset, column=cell.col_idx, value=cell.value)
            row_offset += 1

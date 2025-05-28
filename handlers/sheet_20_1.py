def merge_20_1_report(ws_list, output_ws):
    # Add a title or starting label
    output_ws.range("A1").value = "Merged data from 20.1 | Report"
    row_offset = 2  # start copying from row 2

    for ws in ws_list:
        # Copy values from A2:D5 of each input worksheet
        data = ws.range("A2:D5").value  # returns a 2D list
        if data:
            output_ws.range(f"A{row_offset}").value = data
            row_offset += len(data)

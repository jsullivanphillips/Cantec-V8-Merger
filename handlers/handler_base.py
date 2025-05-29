from typing import List, Tuple
import re

merge_conflict_log = []


def get_cell_format_signature(cell) -> Tuple:
    """Returns a simplified signature of a cell's format."""
    bold = cell.api.Font.Bold
    font_rgb = cell.api.Font.Color
    fill_rgb = cell.api.Interior.Color

    # Normalize fill: treat None, white, or default fills the same
    if fill_rgb in (None, 0xFFFFFF, -4142):  # -4142 is xlNone
        fill_rgb = "NO_FILL"

    return (bold, font_rgb, fill_rgb)


def is_meaningful_value(val):
    """Returns False for any empty, false, or default state values."""
    if val is None:
        return False

    if isinstance(val, bool):
        return val  # False -> not meaningful, True -> meaningful

    if isinstance(val, str):
        return val.strip().lower() not in ("", "false")

    return True  # For numbers, dates, objects, etc.


def merge_checkbox_groups(
    ws_file_list, output_ws, checkbox_groups, tech_col_letter=None
):
    """
    Merges YES/NO/N/A checkbox-style groups.
    Flags conflicts when multiple boxes are checked by different techs.
    Allows merging if all techs agree on the same box.
    """
    for group in checkbox_groups:
        true_cells = []  # (address, value, filename, format_signature)

        for ws, filename in ws_file_list:
            for addr in group:
                val = ws.range(addr).value
                if isinstance(val, bool) and val is True:
                    sig = get_cell_format_signature(ws.range(addr))
                    true_cells.append((addr, val, filename, sig))

        if not true_cells:
            # No box checked by any tech — skip
            continue

        addresses = {entry[0] for entry in true_cells}

        if len(addresses) == 1:
            # ✅ All checkmarks are on the same cell — merge it
            addr, val, filename, sig = true_cells[0]
            cell = output_ws.range(addr)
            cell.value = True
            if sig[0]:
                cell.api.Font.Bold = True
            cell.api.Font.Color = sig[1]
            if sig[2] != "NO_FILL":
                cell.api.Interior.Color = sig[2]

            if insert_or_fill_technician_column:
                row_index = cell.row
                insert_or_fill_technician_column(
                    output_ws, row_index, filename, tech_col_letter
                )
        else:
            # ❌ Conflict — multiple different boxes selected in the same group
            for addr, val, filename, sig in true_cells:
                cell = output_ws.range(addr)
                apply_conflict_highlight(cell)
                add_conflict_comment(
                    cell,
                    [(entry[2], entry[1], entry[3]) for entry in true_cells],
                    output_ws=output_ws,
                    tech_col_letter=tech_col_letter,
                )


def merge_cells(ws_file_list, output_ws, merge_cells_list, tech_col_letter=None):
    """
    Merges cell-by-cell values with conflict checking.
    Optionally calls on_row_merged(output_ws, row_index, technician_filename)
    after each row write.
    """
    for cell_address in merge_cells_list:
        output_cell = output_ws.range(cell_address)
        row_index = output_cell.row
        recorded = None

        for ws, filename in ws_file_list:
            input_cell = ws.range(cell_address)

            if not is_meaningful_value(input_cell.value):
                continue

            if recorded is None:
                # First meaningful value — accept and write
                output_cell.value = input_cell.value

                # Optional formatting (if not using safe_transfer_formatting yet)
                try:
                    output_cell.api.Font.Bold = input_cell.api.Font.Bold
                    output_cell.api.Font.Color = input_cell.api.Font.Color
                    fill = input_cell.api.Interior.Color
                    if fill not in (None, 0xFFFFFF, -4142):
                        output_cell.api.Interior.Color = fill
                except Exception as e:
                    print(f"⚠️ Formatting error at {cell_address}: {e}")

                recorded = (
                    input_cell.value,
                    get_cell_format_signature(input_cell),
                    filename,
                )

                if insert_or_fill_technician_column:
                    insert_or_fill_technician_column(
                        output_ws, row_index, filename, tech_col_letter
                    )

            else:
                # Check for conflicts
                if not compare_cells(output_cell, input_cell):
                    apply_conflict_highlight(output_cell)
                    conflicts = [
                        ("Original", recorded[0], recorded[1]),
                        (
                            filename,
                            input_cell.value,
                            get_cell_format_signature(input_cell),
                        ),
                    ]
                    add_conflict_comment(
                        output_cell,
                        conflicts,
                        output_ws=output_ws,
                        tech_col_letter=tech_col_letter,
                    )


def compare_cells(cell1, cell2) -> bool:
    """Returns True if value and formatting match, prints debug info if they don't."""
    val1 = cell1.value
    val2 = cell2.value

    fmt1 = get_cell_format_signature(cell1)
    fmt2 = get_cell_format_signature(cell2)

    if val1 != val2 or fmt1 != fmt2:
        print(f"⚠️ Cell mismatch at {cell1.address}:")
        if val1 != val2:
            print(f"  - Value mismatch: {val1!r} vs {val2!r}")
        if fmt1 != fmt2:
            print("  - Format mismatch:")
            print(f"    - cell1: {format_signature_to_string(fmt1)}")
            print(f"    - cell2: {format_signature_to_string(fmt2)}")
        return False

    return True


def apply_conflict_highlight(cell):
    """Applies a pastel purple background to a cell."""
    pastel_purple_rgb = (221, 210, 233)  # Hex #DDD2E9
    cell.color = pastel_purple_rgb


def add_conflict_comment(
    cell, conflicts: List[Tuple[str, any, Tuple]], output_ws=None, tech_col_letter=None
):
    """
    Adds a readable comment showing only meaningful differences.
    Merges with any existing conflicts already on the cell.
    """
    print(f"📝 add_conflict_comment: Starting for cell {cell.address}")

    if len(conflicts) < 1:
        return

    # Try resolving "Original" to actual filename
    row_index = cell.row
    original_filename = "Original"
    if output_ws and tech_col_letter:
        tech_cell = output_ws.range(f"{tech_col_letter}{row_index}")
        if tech_cell.value:
            original_filename = tech_cell.value

    # Update "Original" label if needed
    updated_conflicts = []
    for filename, val, fmt in conflicts:
        source = clean_filename(
            original_filename if filename == "Original" else filename
        )
        updated_conflicts.append((source, val, fmt))

    # Load existing comment entries if any
    existing_entries = set()
    if cell.api.Comment:
        comment_text = cell.api.Comment.Text()
        match_lines = re.findall(r"'(.*?)' \((.*?)\)", comment_text)
        for val_str, filename in match_lines:
            existing_entries.add((filename.strip(), val_str.strip()))

    # Add new entries
    for filename, val, fmt in updated_conflicts:
        val_str = str(val).strip().strip("'")  # Normalize
        existing_entries.add((clean_filename(filename.strip()), val_str.strip()))

    # Build comment text
    lines = [f"'{val}' ({filename})" for filename, val in sorted(existing_entries)]
    if len(existing_entries) > 1:
        comment_text = "[Conflict]\n" + "\n".join(lines)
    else:
        print("✅ Only one unique value — no comment needed.")
        return

    # Replace comment
    try:
        if output_ws:
            clean_address = cell.address.replace("$", "")
            merge_conflict_log.append(f"{output_ws.name}, Cell {clean_address}")
    except Exception as e:
        print(f"⚠️ Failed to log conflict at {cell.address}: {e}")

    try:
        if cell.api.Comment:
            cell.api.Comment.Delete()
        cell.api.AddComment(comment_text)
        print(f"✅ Updated comment:\n{comment_text}")
    except Exception as e:
        print(f"❌ Failed to add comment: {e}")


def clean_filename(name: str) -> str:
    return name.replace(".xlsx", "").strip()


def short_format_description(fmt: Tuple) -> str:
    bold = "Bold" if fmt[0] else "Regular"
    fill = fmt[2]
    fill_str = "No Fill" if fill == "NO_FILL" else "Colored Fill"
    return f"{bold}, {fill_str}"


def format_signature_to_string(fmt: Tuple) -> str:
    """Formats the format tuple into a human-readable string."""
    bold = "Bold" if fmt[0] else "Regular"
    font_rgb = fmt[1]
    fill_rgb = fmt[2]

    if fill_rgb == "NO_FILL":
        fill_str = "No Fill"
    else:
        fill_str = f"FillRGB={fill_rgb}"

    return f"{bold}, FontRGB={font_rgb}, {fill_str}"


def insert_or_fill_technician_column(
    output_ws, row_index: int, technician_name: str, col_letter: str
):
    """
    Writes the technician name at the specified row in a target column,
    and hides the column if not already hidden.
    """
    cell = output_ws.range(f"{col_letter}{row_index}")
    cell.value = technician_name

    # Hide the column using the COM API layer
    try:
        cell.api.EntireColumn.Hidden = True
    except Exception as e:
        print(f"⚠️ Could not hide column {col_letter}: {e}")

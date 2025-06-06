from typing import List, Tuple
import re

from math import sqrt

merge_conflict_log = []


NAMED_COLORS = {
    "Red": (255, 0, 0),
    "Excel Red": (192, 80, 77),
    "Green": (0, 255, 0),
    "Dark Green": (0, 128, 0),
    "Yellow": (255, 255, 0),
}


def int_to_rgb(color_int):
    if not isinstance(color_int, int):
        return None
    b = color_int // 65536
    g = (color_int % 65536) // 256
    r = color_int % 256
    return (r, g, b)


def closest_named_color(rgb):
    if rgb is None:
        return "No fill"

    best_match = "Unknown"
    min_dist = float("inf")

    for name, std_rgb in NAMED_COLORS.items():
        if std_rgb is None:
            continue
        dist = sqrt(sum((c1 - c2) ** 2 for c1, c2 in zip(rgb, std_rgb)))
        if dist < min_dist:
            min_dist = dist
            best_match = name

    print(f"🎨 Color match for RGB {rgb} → {best_match}")
    return best_match


def get_cell_format_signature(cell) -> Tuple:
    bold = cell.api.Font.Bold
    font_rgb = cell.api.Font.Color
    fill_rgb = cell.api.Interior.Color

    # Normalize "no fill"
    if fill_rgb in (None, 0xFFFFFF, -4142):
        fill_rgb = "NO_FILL"
    else:
        try:
            fill_rgb = int(fill_rgb)
        except Exception as e:
            print(f"Could not parse rgb as int! {e}")
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


def is_cell_meaningful(val):
    """
    Returns True if the value is meaningful for merge logic:
    - Not None
    - Not empty string
    - Not "0" (string or int) if representing blank formula output
    """
    if val is None:
        return False
    if isinstance(val, str):
        val = val.strip()
        return val not in ("", "0")
    if isinstance(val, (int, float)):
        return val != 0
    return True


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
            # ❌ Conflict — multiple checkboxes selected by different techs

            # Assign semantic labels: YES, NO, N/A (left to right)
            labels = ["YES", "NO", "N/A"][: len(group)]
            addr_to_label = dict(zip(group, labels))

            # Build unified conflict data
            conflict_data = []
            for addr, val, filename, sig in true_cells:
                label = addr_to_label.get(addr, "UNKNOWN")
                tech = clean_filename(filename)
                conflict_data.append((label, tech))

            # Build shared comment
            comment_text = "[Conflict]\n" + "\n".join(
                f"{label}: {tech}" for label, tech in sorted(conflict_data)
            )

            # Apply comment and highlight to all checked cells
            for addr, val, filename, sig in true_cells:
                cell = output_ws.range(addr)
                apply_conflict_highlight(cell)
                try:
                    if cell.api.Comment:
                        cell.api.Comment.Delete()
                    cell.api.AddComment(comment_text)
                except Exception as e:
                    print(f"❌ Failed to write comment to {addr}: {e}")

                # Log conflict
                if output_ws:
                    merge_conflict_log.append(output_ws.name)


def is_page_meaningful(ws_file_list, cell_refs):
    """
    Check if any file has a meaningful value in any of the given cell references.
    If so, the page is considered meaningful.
    """
    for ws, _ in ws_file_list:
        for ref in cell_refs:
            if is_cell_meaningful(ws.range(ref).value):
                print(f"{ws.range(ref).value} is meaningful")
                return True
    return False


def merge_cells(ws_file_list, output_ws, merge_cells_list, tech_col_letter=None):
    """
    Merges cell-by-cell values with conflict checking.
    Optionally calls on_row_merged(output_ws, row_index, technician_filename)
    after each row write.
    """
    for cell_address in merge_cells_list:
        output_cell = output_ws.range(cell_address)
        row_index = output_cell.row
        conflicts = []
        winner_value = None
        winner_fmt = None
        winner_filename = None
        best_fill_color = None

        for ws, filename in ws_file_list:
            input_cell = ws.range(cell_address)

            if not is_meaningful_value(input_cell.value):
                continue

            input_val = input_cell.value
            input_fmt = get_cell_format_signature(input_cell)
            fill = input_cell.api.Interior.Color

            # Check if this cell has a non-white, non-default fill
            has_highlight = fill not in (None, 0xFFFFFF, -4142)

            if winner_value is None:
                # First valid input
                output_cell.value = input_val

                try:
                    output_cell.api.Font.Bold = input_cell.api.Font.Bold
                    output_cell.api.Font.Color = input_cell.api.Font.Color
                except Exception as e:
                    print(f"⚠️ Font formatting error at {cell_address}: {e}")

                winner_value = input_val
                winner_fmt = input_fmt
                winner_filename = filename

                if has_highlight:
                    best_fill_color = fill

                if insert_or_fill_technician_column:
                    insert_or_fill_technician_column(
                        output_ws, row_index, filename, tech_col_letter
                    )

            else:
                if has_highlight and best_fill_color in (None, 0xFFFFFF, -4142):
                    # Upgrade from plain fill to highlighted one
                    best_fill_color = fill

                if input_val != winner_value or not formats_equal(
                    input_fmt, winner_fmt
                ):
                    conflicts.append((filename, input_val, input_fmt))

        # Apply final fill after all comparisons
        if best_fill_color not in (None, 0xFFFFFF, -4142):
            try:
                output_cell.api.Interior.Color = best_fill_color
            except Exception as e:
                print(f"⚠️ Fill application error at {cell_address}: {e}")

        # Finalize conflict comment if needed
        if conflicts:
            apply_conflict_highlight(output_cell)
            conflicts.insert(
                0, (winner_filename, winner_value, winner_fmt)
            )  # prepend original
            add_conflict_comment(
                output_cell,
                conflicts,
                output_ws=output_ws,
                tech_col_letter=tech_col_letter,
            )


def formats_equal(fmt1, fmt2) -> bool:
    """
    Returns True if formatting is considered equivalent.
    Ignores fill conflicts if one side is NO_FILL.
    Assumes format is a tuple: (bold, font_rgb, fill_rgb)
    """
    bold1, font1, fill1 = fmt1
    bold2, font2, fill2 = fmt2

    if fill1 == "NO_FILL" or fill2 == "NO_FILL":
        fill_matches = True
    else:
        fill_matches = fill1 == fill2

    return bold1 == bold2 and font1 == font2 and fill_matches


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
        print(f"🧪 DEBUG: fmt[2] for {filename} = {fmt[2]}")  # <--- Add this
        val_str = str(val).strip().strip("'")  # Normalize
        existing_entries.add((clean_filename(filename.strip()), val_str.strip()))

    # Decide comment type
    if len({val for _, val, _ in updated_conflicts}) > 1:
        # Value conflict
        lines = [f"'{val}' ({filename})" for filename, val, _ in updated_conflicts]
        comment_text = "[Conflict]\n" + "\n".join(lines)
    else:
        # Format conflict
        # Filter meaningful fill differences
        non_empty = [
            (filename, format_signature_to_string(fmt))
            for filename, _, fmt in updated_conflicts
            if format_signature_to_string(fmt).lower() != "no fill"
        ]

        # If we have at least one fill color, ignore 'no fill' entries
        final_lines = (
            [f"{filename}: {desc}" for filename, desc in non_empty]
            if non_empty
            else [f"{filename}: No fill" for filename, _, _ in updated_conflicts]
        )

        comment_text = "[Format Conflict]\n" + "\n".join(final_lines)

    # Replace comment
    try:
        if output_ws:
            merge_conflict_log.append(output_ws.name)
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
    fill_rgb = fmt[2]

    if fill_rgb in ("NO_FILL", None, 0xFFFFFF, -4142):
        return "No fill"

    try:
        rgb = int_to_rgb(fill_rgb)
        return f"{closest_named_color(rgb)} fill"
    except Exception as e:
        print(f"⚠️ Failed to interpret fill color: {fill_rgb} → {e}")
        return "Unknown fill"


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

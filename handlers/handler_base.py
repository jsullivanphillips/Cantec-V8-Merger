from typing import List, Tuple

from math import sqrt

merge_conflict_log = []


NAMED_COLORS = {
    "Red 1": (255, 0, 0),
    "Red 2": (192, 80, 77),
    "Green 1": (0, 255, 0),
    "Green 2": (0, 128, 0),
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
                return True
    return False


def strip_cells(conflicts_with_cells):
    return [(filename, val, fmt) for filename, val, fmt, _ in conflicts_with_cells]


def merge_cells(
    ws_file_list: List[Tuple],
    output_ws,
    merge_cells_list: List[str],
    tech_col_letter: str = None,
    special_row_ranges: List[dict] = None,
):
    """
    Merges cell-by-cell values with conflict checking.
    Supports override logic for specific rows where value should come from
    the technician who has a highlight in one of several defined columns.

    Now treats differing highlight colors as conflicts and uses add_conflict_comment
    to report full conflict details (sheet, value, fill color).
    """

    def get_special_range(row_index: int, ranges: List[dict]):
        if not ranges:
            return None
        for entry in ranges:
            if row_index in entry.get("rows", []):
                return entry
        return None

    for cell_address in merge_cells_list:
        output_cell = output_ws.range(cell_address)
        row_index = output_cell.row

        # --- Special merging logic for highlighted rows ---
        special = get_special_range(row_index, special_row_ranges)
        if special:
            value_col = special["value_col"]
            highlight_cols = special["highlight_cols"]
            candidates: List[Tuple[str, any, Tuple, any]] = []

            for ws, filename in ws_file_list:
                val_cell = ws.range(f"{value_col}{row_index}")
                val = val_cell.value
                if not is_meaningful_value(val):
                    continue

                # detect any highlight in the row
                found_fill = None
                for col in highlight_cols:
                    fill = ws.range(f"{col}{row_index}").api.Interior.Color
                    if fill not in (None, 0xFFFFFF, -4142, 0xFFFF00, 65535):
                        found_fill = fill
                        break
                if found_fill is None:
                    continue

                fmt = get_cell_format_signature(val_cell)
                # prefer explicit fill from cell format if set
                fill_color = fmt[2] if fmt[2] != "NO_FILL" else found_fill
                candidates.append((filename, val, fmt, fill_color))

            # exactly one highlighted candidate
            if len(candidates) == 1:
                fn, val, fmt, fill = candidates[0]
                out_val_cell = output_ws.range(f"{value_col}{row_index}")
                out_val_cell.value = val
                # copy font
                try:
                    out_val_cell.api.Font.Bold = val_cell.api.Font.Bold
                    out_val_cell.api.Font.Color = val_cell.api.Font.Color
                except Exception as e:
                    print(f"⚠️ Font error on row {row_index}: {e}")
                # copy highlight cells
                source_ws = next((w for w, f in ws_file_list if f == fn), None)
                if source_ws:
                    for col in highlight_cols:
                        inp = source_ws.range(f"{col}{row_index}")
                        out = output_ws.range(f"{col}{row_index}")
                        try:
                            out.api.Interior.Color = inp.api.Interior.Color
                        except Exception as e:
                            print(f"⚠️ Fill copy error {col}{row_index}: {e}")
                # apply cell fill
                if fill not in (None, 0xFFFFFF, -4142):
                    try:
                        out_val_cell.api.Interior.Color = fill
                    except Exception as e:
                        print(f"⚠️ Fill apply error on {value_col}{row_index}: {e}")
                # technician column
                if insert_or_fill_technician_column:
                    insert_or_fill_technician_column(
                        output_ws, row_index, fn, tech_col_letter
                    )

            # conflict among multiple highlights
            elif len(candidates) > 1:
                apply_conflict_highlight(output_cell)
                # build 4-tuples for comment
                comment_entries = [
                    (fn, val, fmt, fill) for fn, val, fmt, fill in candidates
                ]
                add_conflict_comment(
                    output_cell,
                    comment_entries,
                    output_ws=output_ws,
                    tech_col_letter=tech_col_letter,
                )
            continue  # skip default logic for this row

        # --- Default merging logic below ---
        conflicts: List[Tuple[str, any, Tuple, any]] = []
        winner_value = None
        winner_fmt = None
        winner_filename = None
        best_fill_color = None

        for ws, filename in ws_file_list:
            cell = ws.range(cell_address)
            val = cell.value
            if not is_meaningful_value(val):
                continue
            fmt = get_cell_format_signature(cell)
            fill = cell.api.Interior.Color
            has_hi = fill not in (None, 0xFFFFFF, -4142)

            if winner_value is None:
                # adopt first meaningful value
                output_cell.value = val
                try:
                    output_cell.api.Font.Bold = cell.api.Font.Bold
                    output_cell.api.Font.Color = cell.api.Font.Color
                except Exception as e:
                    print(f"⚠️ Font error at {cell_address}: {e}")
                winner_value = val
                winner_fmt = fmt
                winner_filename = filename
                if has_hi:
                    best_fill_color = fill
                if insert_or_fill_technician_column:
                    insert_or_fill_technician_column(
                        output_ws, row_index, filename, tech_col_letter
                    )
            else:
                # value or format conflict
                if val != winner_value or not formats_equal(fmt, winner_fmt):
                    conflicts.append((filename, val, fmt, fill))
                # highlight color conflict
                if has_hi:
                    if best_fill_color in (None, 0xFFFFFF, -4142):
                        best_fill_color = fill
                    elif fill != best_fill_color:
                        conflicts.append((filename, val, fmt, fill))

        # apply best fill
        if best_fill_color not in (None, 0xFFFFFF, -4142):
            try:
                output_cell.api.Interior.Color = best_fill_color
            except Exception as e:
                print(f"⚠️ Fill error at {cell_address}: {e}")

        if conflicts:
            # prepend original
            conflicts.insert(
                0, (winner_filename, winner_value, winner_fmt, best_fill_color)
            )
            apply_conflict_highlight(output_cell)
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
    cell,
    conflicts: List[Tuple[str, any, Tuple, any]],
    output_ws=None,
    tech_col_letter=None,
):
    """
    Adds a readable comment showing only the conflicting parts of value or fill.
    Each comment begins with [Conflict] and lists only the attributes that differ.
    """
    # No conflicts -> nothing to do
    if not conflicts:
        return

    # Resolve "Original" to actual filename from technician column if available
    row_idx = cell.row
    orig_name = "Original"
    if output_ws and tech_col_letter:
        tech_val = output_ws.range(f"{tech_col_letter}{row_idx}").value
        if tech_val:
            orig_name = tech_val

    # Normalize entries to (sheet, value_str, fill_name)
    entries: List[Tuple[str, str, str]] = []
    for entry in conflicts:
        if len(entry) == 4:
            fn, val, fmt, fill = entry
        else:
            fn, val, fmt = entry
            fill = None
        name = clean_filename(orig_name if fn == "Original" else fn)
        val_str = str(val).strip()
        # Convert fill integer to rgb tuple (Excel is BGR)
        if fill in (None, 0xFFFFFF, -4142):
            rgb = None
        else:
            b = int(fill) & 0xFF
            g = (int(fill) >> 8) & 0xFF
            r = (int(fill) >> 16) & 0xFF
            rgb = (r, g, b)
        fill_name = closest_named_color(rgb)
        entries.append((name, val_str, fill_name))

    # Determine which attribute is actually in conflict
    distinct_vals = {v for _, v, _ in entries}
    distinct_fills = {f for _, _, f in entries}

    lines: List[str] = []
    if len(distinct_vals) > 1:
        # Value conflict
        for name, v, _ in entries:
            lines.append(f"{name}: '{v}'")
    elif len(distinct_fills) > 1:
        # Fill (color) conflict
        for name, _, f in entries:
            lines.append(f"{name}: fill={f}")
    else:
        # Fallback: show both
        for name, v, f in entries:
            lines.append(f"{name}: '{v}' | fill={f}")

    comment_text = "[Conflict]\n" + "\n".join(lines)

    # Log conflict
    try:
        if output_ws:
            merge_conflict_log.append(output_ws.name)
    except Exception:
        pass

    # Replace existing comment on the cell
    try:
        try:
            cell.api.ClearComments()
        except Exception:
            if getattr(cell.api, "Comment", None):
                cell.api.Comment.Delete()
        cell.api.AddComment(comment_text)
    except Exception as e:
        print(f"❌ Failed to add conflict comment on {cell.address}: {e}")


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

from typing import List, Tuple

from math import sqrt

merge_conflict_log = []

NAMED_COLORS = {
    "Red": [(255, 0, 0), (192, 80, 77), (218, 150, 148)],
    "Green": [(0, 255, 0), (0, 128, 0), (196, 215, 155), (0, 176, 80), (146, 208, 80)],
    "Yellow": [(255, 255, 0)],
}


def closest_named_color(rgb, threshold=90):
    if rgb is None:
        return "No fill"

    best_match = "Unknown"
    min_dist = float("inf")

    for name, color_list in NAMED_COLORS.items():
        for std_rgb in color_list:
            dist = sqrt(sum((c1 - c2) ** 2 for c1, c2 in zip(rgb, std_rgb)))
            if dist < min_dist:
                min_dist = dist
                best_match = name

    if min_dist > threshold:
        return f"RGB{rgb}"  # fallback to raw RGB if too far from any label
    return best_match


def int_to_rgb(color_int):
    """Converts Excel BGR color int into RGB tuple."""
    if not isinstance(color_int, int):
        return None
    r = color_int & 0xFF
    g = (color_int >> 8) & 0xFF
    b = (color_int >> 16) & 0xFF
    return (r, g, b)


def get_cell_format_signature(cell) -> Tuple:
    bold = cell.api.Font.Bold
    font_rgb = cell.api.Font.Color
    try:
        raw_fill = cell.api.DisplayFormat.Interior.Color
    except Exception:
        raw_fill = cell.api.Interior.Color  # fallback
    norm = normalize_fill(raw_fill)
    fill_rgb = norm if norm is not None else "NO_FILL"

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


def normalize_fill(fill):
    if fill in (None, "NO_FILL", 0xFFFFFF, -4142):
        return None
    try:
        return int(fill)
    except Exception as e:
        print(f"Failed to normalize fill with error {e}")
        return None


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
                    raw_fill = ws.range(f"{col}{row_index}").api.Interior.Color
                    norm_fill = normalize_fill(raw_fill)
                    if norm_fill is not None and norm_fill not in (
                        0xFFFF00,
                        65535,
                    ):  # skip yellow or Excel light yellow
                        found_fill = norm_fill
                        break
                if found_fill is None:
                    continue

                fmt = get_cell_format_signature(val_cell)
                # prefer explicit fill from cell format if set
                fill_color = found_fill
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
                fill = normalize_fill(fill)
                if fill is not None:
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
        best_fill_rgb = None  # For display / comparison
        best_fill_excel = None  # For setting in Excel
        winner_fill = None

        for ws, filename in ws_file_list:
            cell = ws.range(cell_address)
            val = cell.value
            if not is_meaningful_value(val):
                continue
            fmt = get_cell_format_signature(cell)
            try:
                raw_fill = cell.api.DisplayFormat.Interior.Color
            except Exception:
                raw_fill = cell.api.Interior.Color

            norm_fill = normalize_fill(raw_fill)
            rgb_fill = int_to_rgb(norm_fill)
            has_hi = rgb_fill is not None
            print(f"{cell_address} | norm_fill: {norm_fill} | rgb_fill: {rgb_fill}")

            if winner_fill is None:
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
                winner_fill = norm_fill

                if has_hi:
                    best_fill_rgb = rgb_fill
                    best_fill_excel = norm_fill

                if insert_or_fill_technician_column:
                    insert_or_fill_technician_column(
                        output_ws, row_index, filename, tech_col_letter
                    )
            else:
                conflict_entry = (filename, val, fmt, norm_fill)
                if (
                    val != winner_value
                    or not formats_equal(fmt, winner_fmt)
                    or (has_hi and rgb_fill != best_fill_rgb)
                ):
                    if conflict_entry not in conflicts:
                        conflicts.append(conflict_entry)

        # apply best fill
        if best_fill_excel not in (None, 0xFFFFFF, -4142):
            try:
                output_cell.api.Interior.Color = best_fill_excel
            except Exception as e:
                print(f"⚠️ Fill error at {cell_address}: {e}")

        if conflicts:
            # find the original cell and use its actual fill
            winner_ws = next(
                (ws for ws, fn in ws_file_list if fn == winner_filename), None
            )
            if winner_ws:
                winner_fmt = get_cell_format_signature(winner_ws.range(cell_address))
                winner_fill = best_fill_excel
            else:
                winner_fill = best_fill_excel
            conflicts.insert(
                0, (winner_filename, winner_value, winner_fmt, winner_fill)
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
    if not conflicts:
        return

    # Figure out the "Original" sheet name, if any
    row_idx = cell.row
    orig_name = "Original"
    if output_ws and tech_col_letter:
        tech_val = output_ws.range(f"{tech_col_letter}{row_idx}").value
        if tech_val:
            orig_name = tech_val

    # Build a list of (sheet, value_str, fill_name)
    entries: List[Tuple[str, str, str]] = []
    for entry in conflicts:
        # unpack the 4-tuple (filename, val, fmt, fill)
        fn, val, fmt, fill = entry
        name = clean_filename(orig_name if fn == "Original" else fn)
        val_str = str(val).strip()

        rgb = int_to_rgb(normalize_fill(fill))

        # map to the nearest name
        fill_name = closest_named_color(rgb)
        entries.append((name, val_str, fill_name))

    # figure out which attribute changed
    vals = {v for _, v, _ in entries}
    fills = {f for _, _, f in entries}

    lines: List[str] = []
    if len(vals) > 1:
        # value conflict
        for name, v, _ in entries:
            lines.append(f"{name}: '{v}'")
    elif len(fills) > 1:
        # fill conflict
        for name, _, f in entries:
            lines.append(f"{name}: fill={f}")
    else:
        # fallback: show both
        for name, v, f in entries:
            lines.append(f"{name}: '{v}' | fill={f}")

    comment_text = "[Conflict]\n" + "\n".join(lines)
    print(f"[DEBUG] {name}: raw_fill={fill}, rgb={rgb}, closest={fill_name}")

    # log and replace any existing comment
    try:
        if output_ws:
            merge_conflict_log.append(output_ws.name)
    except Exception:
        pass

    try:
        # clear old
        try:
            cell.api.ClearComments()
        except Exception as e:
            print(f"could not clear comment with error {e}. Deleting instead.")
            if getattr(cell.api, "Comment", None):
                cell.api.Comment.Delete()
        # add new
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

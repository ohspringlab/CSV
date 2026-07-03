"""
CSV parsing logic for cabinet cut list files.
Handles complex cabinet number formats and row expansion.
"""

import csv
import re
from fractions import Fraction


# File color palette (index 0 = first file, etc.)
FILE_COLORS = [
    "#FFFFFF",  # white
    "#FFFACD",  # light yellow
    "#E0F0FF",  # light blue
    "#E0FFE0",  # light green
    "#FFE4E1",  # light rose
    "#F0E6FF",  # light lavender
    "#FFE8CC",  # light orange
    "#E6FFFA",  # light teal
]

# Piece-name normalisation: raw CSV name → canonical name used in the app.
# Add entries here whenever a CSV uses a variant spelling.
# None means "drop this row entirely" (excluded parts).
PIECE_NAME_MAP: dict[str, str | None] = {
    "Front Stretcher": None,   # excluded
    "Rear Stretcher":  None,
    "Drw Stretcher":   None,
    "Front Toe":       None,
    "Side Toe":        None,
}


def decimal_to_fraction_str(value: float) -> str:
    """
    Convert a decimal number to a mixed number fraction string.
    e.g. 23.5 -> "23 1/2", 0.0625 -> "1/16", 35.5625 -> "35 9/16"
    Supported denominators: 2, 4, 8, 16, 32
    """
    if value == 0:
        return "0"

    negative = value < 0
    value = abs(value)
    whole = int(value)
    decimal_part = value - whole

    if decimal_part < 0.0001:
        result = str(whole)
    else:
        # Try to express decimal as a fraction with denominator up to 32
        frac = Fraction(decimal_part).limit_denominator(32)
        if frac.denominator == 1:
            # Rounds to a whole
            result = str(whole + frac.numerator)
        elif whole == 0:
            result = f"{frac.numerator}/{frac.denominator}"
        else:
            result = f"{whole} {frac.numerator}/{frac.denominator}"

    return f"-{result}" if negative else result


def parse_cabinet_num(raw: str) -> list:
    """
    Parse a cabinet_num field and return list of (cabinet_id, explicit_qty) tuples.

    Formats handled:
      "1"                -> [("1", None)]
      "N1"               -> [("N1", None)]
      "R2:1"             -> [("R2:1", None)]
      "2&3"              -> [("2", None), ("3", None)]
      "2(2)"             -> [("2", 2)]
      "R2:1&3"           -> [("R2:1", None), ("R2:3", None)]   ← prefix inherited
      "R2:1(2)&3(2)"     -> [("R2:1", 2),    ("R2:3", 2)]      ← prefix inherited
      "R2:1(2)&2(2)&3(2)"-> [("R2:1", 2),    ("R2:2", 2), ("R2:3", 2)]
      "2(2)&3(2)"        -> [("2", 2),        ("3", 2)]
      "1(2)&2(2)&3(2)"   -> [("1", 2),        ("2", 2),   ("3", 2)]

    Prefix inheritance rule: when the first cabinet has an alpha prefix
    (e.g. "R2:") and a subsequent part is a bare integer (optionally with
    explicit qty), the same prefix is prepended automatically.
    """
    raw = raw.strip()
    if not raw:
        return [("?", None)]

    parts = raw.split("&")
    result = []

    # Detect prefix of first part (e.g. "R2:" from "R2:1")
    first_part = parts[0].strip()
    # Strip trailing (n) to get the pure cabinet token
    first_token = re.sub(r'\(\d+\)$', '', first_part).strip()
    # Prefix = everything up to and including the last non-digit character
    prefix_match = re.match(r'^(.*[^0-9])(\d+)$', first_token)
    inherited_prefix = prefix_match.group(1) if prefix_match else ""

    for part in parts:
        part = part.strip()
        # cabinet ID (letters, digits, spaces, hyphens, colons)
        # followed by optional explicit qty in parentheses
        match = re.match(r'^([A-Za-z0-9\s\-:]+?)(?:\((\d+)\))?$', part)
        if match:
            cabinet_id  = match.group(1).strip()
            explicit_qty = int(match.group(2)) if match.group(2) else None

            # If this part is a bare integer AND the first part had a prefix,
            # inherit that prefix (e.g. "3" → "R2:3")
            if inherited_prefix and re.match(r'^\d+$', cabinet_id):
                cabinet_id = inherited_prefix + cabinet_id

            result.append((cabinet_id, explicit_qty))
        else:
            result.append((part, None))

    return result


def expand_rows(row_dict: dict) -> list:
    """
    Expand a single parsed row into multiple rows, one per cabinet.
    Parenthesized values are interpreted as the quantity for that cabinet:
      qty 6 + "2(3)&3(3)" -> cabinet 2 qty 3, cabinet 3 qty 3
      qty 3 + "2&3(2)"    -> cabinet 2 qty 1, cabinet 3 qty 2
    Unspecified quantities receive the remaining quantity, split evenly.
    Returns list of row dicts.
    """
    raw = row_dict.get("cabinet_num_raw", "")
    cabinets = parse_cabinet_num(raw)
    if not cabinets:
        return [row_dict]

    total_qty = row_dict.get("qty", 1)
    explicit_sum = sum(qty for _, qty in cabinets if qty is not None)
    unspecified_count = sum(1 for _, qty in cabinets if qty is None)
    remaining_qty = max(0, total_qty - explicit_sum)
    base_qty = remaining_qty // unspecified_count if unspecified_count else 0
    remainder = remaining_qty % unspecified_count if unspecified_count else 0

    expanded = []
    unspecified_seen = 0
    for cabinet_id, explicit_qty in cabinets:
        if explicit_qty is None:
            qty = base_qty + (1 if unspecified_seen < remainder else 0)
            unspecified_seen += 1
        else:
            qty = explicit_qty
        new_row = dict(row_dict)
        new_row["cabinet_id"] = cabinet_id
        new_row["unit_qty_spec"] = explicit_qty
        new_row["qty"] = qty
        expanded.append(new_row)

    return expanded


def _detect_col_offset(cols: list[str]) -> int:
    """
    Detect column offset for files that insert an extra numeric column
    before the piece name (e.g. door order CSVs with a depth/thickness col).

    Normal layout:  col3=piece_name, col4=cabinet_num
    Shifted layout: col3=number(depth), col4=piece_name, col5=cabinet_num

    Returns 0 for normal, 1 for shifted.
    """
    col3 = cols[3].strip() if len(cols) > 3 else ""
    col4 = cols[4].strip() if len(cols) > 4 else ""
    # col3 is purely numeric AND col4 looks like a piece name (letters / braces)
    if re.match(r'^-?\d*\.?\d+$', col3) and re.search(r'[A-Za-z{]', col4):
        return 1
    return 0


def parse_value(val: str):
    """Parse a string value to float, return 0.0 on failure."""
    val = val.strip()
    if not val:
        return 0.0
    try:
        return float(val)
    except ValueError:
        return 0.0


def parse_csv_file(filepath: str, source_color: str = "#FFFFFF") -> list:
    """
    Parse a CSV file (no headers).

    Supports two layouts (auto-detected per line):
      Normal:  qty, width, length, piece_name, cabinet_num, description, color_code
      Shifted: qty, width, length, depth(ignored), piece_name, cabinet_num, description, color_code

    Returns list of expanded row dicts.
    """
    rows = []

    with open(filepath, newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        for line_num, cols in enumerate(reader, start=1):
            if not cols or all(c.strip() == "" for c in cols):
                continue

            # Pad to at least 8 columns to accommodate shifted layout
            while len(cols) < 8:
                cols.append("")

            try:
                # Detect if this row uses the shifted layout (extra col before piece_name)
                offset = _detect_col_offset(cols)

                piece_name_raw = cols[3 + offset].strip()

                # Apply normalisation / exclusion map
                if piece_name_raw in PIECE_NAME_MAP:
                    canonical = PIECE_NAME_MAP[piece_name_raw]
                    if canonical is None:
                        continue   # excluded — skip row
                    piece_name_raw = canonical

                raw_row = {
                    "qty":             int(float(cols[0].strip())) if cols[0].strip() else 1,
                    "width":           parse_value(cols[1]),
                    "length":          parse_value(cols[2]),
                    "piece_name":      piece_name_raw,
                    "cabinet_num_raw": cols[4 + offset].strip(),
                    "description":     cols[5 + offset].strip(),
                    "color_code":      cols[6 + offset].strip(),
                    "source_file":     filepath,
                    "source_color":    source_color,
                }
                expanded = expand_rows(raw_row)
                rows.extend(expanded)
            except Exception as e:
                print(f"Warning: skipping line {line_num} in {filepath}: {e}")
                continue

    return rows

"""
Data transformation logic: pivot parsed rows into per-cabinet output rows.

Column mapping (from nomenclature.json piece_aliases):
  CÔTÉS          ← PIGNON (L) + PIGNON (R)   (merged when dims match)
  DESSUS/DESSOUS ← DESSUS/DESSOUS             (kept as-is, multi-row per source)
  DOS            ← DOS
  TAB AJUST      ← TAB AJUST
  ADJACENTS      ← TAB FIXE
  PARTITION      ← Door (panelized end), PANEL
  PORTES/FAÇADES ← Door(L), Door(R), Drawer   (merged L+R when dims match, then Drawer)
  Stile          ← Stile
  Toe Skin       ← Toe Skin
  Drw L Side, Drw R Side, Drw Front, Drw Back, Drw Bottom  → drawer sheet
"""

import json
import re
from pathlib import Path

_CONFIG_PATH = Path(__file__).resolve().parent.parent / "config" / "nomenclature.json"

# ── Merged-column definitions ──────────────────────────────────────────────
# These canonical column names aggregate multiple raw piece_names into one column.
# Door(L)+Door(R) pairs with matching dims are merged (qty×2); Drawer is appended.
_MERGED_COLS: dict[str, list[str]] = {
    "PORTES/FAÇADES": ["Door(L)", "Door(R)", "Drawer"],
}

# piece_aliases from config: raw_name → canonical_column_name  (or None = exclude)
# Rows whose raw piece_name maps to None are completely dropped.
# Rows whose raw piece_name maps to a column name are collected into that column.
_ALIASES: dict[str, str | None] = {}


def _load_aliases():
    """Load piece_aliases from nomenclature.json into _ALIASES."""
    global _ALIASES
    try:
        with _CONFIG_PATH.open("r", encoding="utf-8") as f:
            cfg = json.load(f)
        _ALIASES = cfg.get("piece_aliases", {})
    except Exception:
        _ALIASES = {}


_load_aliases()


def _cabinet_sort_key(cabinet_id: str):
    """
    Sort cabinet IDs:
      0 — pure numeric (1, 2, 3, 10…) sorted numerically
      1 — alpha-prefixed, grouped by prefix then number (N1,N2… R1,R2…)
          Handles colon sub-IDs: R2:1 < R2:2 < R2:10
          Handles colon+alpha sub-IDs: R5:N2 < R5:N4 < R5:N11
      2 — everything else lexicographically
    """
    if re.match(r'^\d+$', cabinet_id):
        return (0, "", int(cabinet_id), "", 0)

    # R5:N11 style — outer alpha+digit prefix + colon + inner alpha-prefix + number
    m = re.match(r'^([A-Za-z]+\d+):([A-Za-z]+)(\d+)$', cabinet_id)
    if m:
        return (1, m.group(1).upper(), 0, m.group(2).upper(), int(m.group(3)))

    # R2:1 style — alpha prefix + digit + colon + number
    m = re.match(r'^([A-Za-z]+)(\d+):(\d+)$', cabinet_id)
    if m:
        return (1, m.group(1).upper(), int(m.group(2)), "", int(m.group(3)))

    # N1, R2 style — simple alpha + number
    m = re.match(r'^([A-Za-z]+)(\d+)$', cabinet_id)
    if m:
        return (1, m.group(1).upper(), int(m.group(2)), "", 0)

    return (2, cabinet_id, 0, "", 0)


def _aggregate_piece_entries(piece_rows: list) -> list:
    """
    Aggregate rows for one raw piece_name.

    Key = (width, length, color_code, source_color) so identical pieces from
    the same file are summed, while pieces from different files (different
    source_color) stay as separate entries — preserving each file's highlight.
    """
    grouped: dict = {}
    for row in piece_rows:
        key = (
            row.get("width", ""),
            row.get("length", ""),
            row.get("color_code", ""),
            row.get("source_color", "#FFFFFF"),
        )
        if key not in grouped:
            grouped[key] = {
                "qty": 0,
                "width":        row.get("width", ""),
                "length":       row.get("length", ""),
                "source_color": row.get("source_color", "#FFFFFF"),
                "source_file":  row.get("source_file", ""),
                "color_code":   row.get("color_code", ""),
                "description":  row.get("description", ""),
            }
        grouped[key]["qty"] += row.get("qty", 0)

    return sorted(
        grouped.values(),
        key=lambda e: (
            str(e.get("source_file", "")),
            str(e.get("color_code", "")),
            float(e.get("width") or 0),
            float(e.get("length") or 0),
        ),
    )


def _try_merge_pair(entries_a: list, entries_b: list) -> list | None:
    """
    Merge two entry lists when every pair has identical dimensions.
    Returns merged list (qty summed per pair) or None if dims differ.
    Highlight color = whichever entry is non-white.
    """
    if len(entries_a) != len(entries_b):
        return None
    merged = []
    for ea, eb in zip(entries_a, entries_b):
        if ea.get("width") != eb.get("width") or ea.get("length") != eb.get("length"):
            return None
        sc_a = ea.get("source_color", "#FFFFFF")
        sc_b = eb.get("source_color", "#FFFFFF")
        merged.append({
            "qty":          (ea.get("qty") or 0) + (eb.get("qty") or 0),
            "width":        ea.get("width", ""),
            "length":       ea.get("length", ""),
            "source_color": sc_a if sc_a != "#FFFFFF" else sc_b,
            "source_file":  ea.get("source_file", ""),
            "color_code":   ea.get("color_code", ""),
            "description":  ea.get("description", ""),
        })
    return merged


def _collect_merged_col(col_name: str, raw_names: list[str], pieces: dict) -> list:
    """
    Build entries for a merged column (e.g. CÔTÉS from PIGNON L + PIGNON R).

    Strategy for a 2-source merge (L + R):
      - If dims match → merge into one set of entries (qty × 2)
      - If dims differ → concatenate both lists

    For 3+ sources (PORTES/FAÇADES: Door(L), Door(R), Drawer):
      - Try merge Door(L) + Door(R) first, then append Drawer entries
    """
    if len(raw_names) == 0:
        return []

    if len(raw_names) == 1:
        return _aggregate_piece_entries(pieces.get(raw_names[0], []))

    if len(raw_names) == 2:
        a = _aggregate_piece_entries(pieces.get(raw_names[0], []))
        b = _aggregate_piece_entries(pieces.get(raw_names[1], []))
        if a and b:
            merged = _try_merge_pair(a, b)
            return merged if merged is not None else (a + b)
        return a or b

    # 3 sources: merge first two, then append rest
    entries = _collect_merged_col(col_name, raw_names[:2], pieces)
    for extra_name in raw_names[2:]:
        extra = _aggregate_piece_entries(pieces.get(extra_name, []))
        entries = entries + extra
    return entries


def _set_piece(out_row: dict, pname: str, entries: list):
    """Write piece entries (or blanks) into out_row."""
    if entries:
        first = entries[0]
        out_row[f"{pname}_entries"] = entries
        out_row[f"{pname}_qty"]     = sum(e.get("qty", 0) for e in entries)
        out_row[f"{pname}_w"]       = first.get("width", "")
        out_row[f"{pname}_l"]       = first.get("length", "")
        out_row[f"{pname}_color"]   = first.get("source_color", "#FFFFFF")
    else:
        out_row[f"{pname}_entries"] = []
        out_row[f"{pname}_qty"]     = ""
        out_row[f"{pname}_w"]       = ""
        out_row[f"{pname}_l"]       = ""
        out_row[f"{pname}_color"]   = ""


def transform(rows: list, piece_names: list) -> list:
    """
    Pivot expanded rows into one dict per cabinet_id.

    Column resolution order for each pname:
    1. _MERGED_COLS  → special multi-source merge (PORTES/FAÇADES)
    2. alias_raw     → one or more raw names redirect to this column
                       (e.g. TAB FIXE rows are appended to DESSUS/DESSOUS)
                       The column's own direct rows are ALSO included first.
    3. direct match  → raw piece_name == pname
    """
    # Build reverse alias map: canonical_col → [raw_names that point to it]
    # Skip aliases whose target is None (those rows are excluded at parse time)
    # or whose target is already handled by _MERGED_COLS.
    alias_raw: dict[str, list[str]] = {}
    for raw, canonical in _ALIASES.items():
        if canonical is None:
            continue                         # excluded — parser already drops them
        if canonical in _MERGED_COLS:
            continue                         # handled by merged-col logic
        alias_raw.setdefault(canonical, []).append(raw)

    # Group raw rows by cabinet → raw piece_name
    cabinets: dict = {}
    for row in rows:
        cid   = row.get("cabinet_id", "?")
        pname = row.get("piece_name", "")
        if cid not in cabinets:
            cabinets[cid] = {}
        if pname not in cabinets[cid]:
            cabinets[cid][pname] = []
        cabinets[cid][pname].append(row)

    output = []
    for cabinet_id in sorted(cabinets.keys(), key=_cabinet_sort_key):
        pieces  = cabinets[cabinet_id]
        out_row = {"cabinet_id": cabinet_id}

        for pname in piece_names:

            if pname in _MERGED_COLS:
                # e.g. PORTES/FAÇADES ← Door(L)+Door(R)+Drawer (merge L+R when dims match)
                entries = _collect_merged_col(pname, _MERGED_COLS[pname], pieces)
                # Also append any aliased raw names pointing to this column
                # (e.g. {Door Size}, {Drawer Front Size} from door-order CSVs)
                for raw, target in _ALIASES.items():
                    if target == pname and raw not in _MERGED_COLS[pname]:
                        entries = entries + _aggregate_piece_entries(pieces.get(raw, []))

            elif pname in alias_raw:
                # Column gathers its own direct rows PLUS all aliased rows.
                # Example: DESSUS/DESSOUS gets raw DESSUS/DESSOUS rows first,
                # then TAB FIXE rows appended below them.
                direct  = _aggregate_piece_entries(pieces.get(pname, []))
                aliased: list = []
                for raw in alias_raw[pname]:
                    aliased += _aggregate_piece_entries(pieces.get(raw, []))
                entries = direct + aliased

            else:
                # Direct name match
                entries = _aggregate_piece_entries(pieces.get(pname, []))

            _set_piece(out_row, pname, entries)

        output.append(out_row)

    return output

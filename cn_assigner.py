"""
cn_assigner.py — CN (Calculation Number) Assignment Engine for Level 1 lines.

A CN is a Caesar II model boundary: a group of piping lines that must be analyzed
together because thermal forces transfer between them.

This module PROPOSES CN groupings. The engineer is always the final decision maker.
Every proposal includes a written reason and a review flag.

Only Level 1 lines receive CN assignments. Level 2 and Level 3 are untouched.
"""

from collections import defaultdict, deque
from typing import Optional
import re

import pandas as pd


# ─────────────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────────────

def assign_cns(
    classified_df: pd.DataFrame,
    rules: dict,
) -> tuple[pd.DataFrame, list[dict]]:
    """
    Assign Calculation Numbers to Level 1 lines.

    Returns:
        result_df   — classified_df with four CN columns added for Level 1 rows;
                      non-Level-1 rows have empty strings in those columns.
        cn_proposals — list of dicts, one per proposed CN, used for the
                       CN Proposals and Dashboard output sheets.
    """
    cn_cfg = rules.get("cn_settings", {})
    if not cn_cfg:
        return classified_df, []

    result_df = classified_df.copy()
    for col in ("Proposed_CN", "CN_Reason", "CN_Review_Flag", "CN_Status"):
        result_df[col] = ""

    l1_indices = result_df.index[result_df["Level"] == "Level 1"].tolist()
    if not l1_indices:
        return result_df, []

    l1_df = result_df.loc[l1_indices]
    lines = _build_line_records(l1_df, cn_cfg)
    groups = _group_into_cns(lines, cn_cfg)
    cn_proposals = _finalize_cn_assignments(groups, lines, cn_cfg)

    for proposal in cn_proposals:
        for df_idx in proposal["line_df_indices"]:
            result_df.at[df_idx, "Proposed_CN"] = proposal["cn_number"]
            result_df.at[df_idx, "CN_Reason"] = proposal["grouping_reason"]
            result_df.at[df_idx, "CN_Review_Flag"] = proposal["review_flag"]
            result_df.at[df_idx, "CN_Status"] = "PROPOSED"

    return result_df, cn_proposals


# ─────────────────────────────────────────────────────────────────────────────
# Step 1 — Build line records from the Level 1 DataFrame rows
# ─────────────────────────────────────────────────────────────────────────────

def _build_line_records(l1_df: pd.DataFrame, cn_cfg: dict) -> list[dict]:
    equip_prefixes = cn_cfg.get("equipment_type_prefixes", {})
    area_col_override = (cn_cfg.get("column_mapping") or {}).get("area_code")

    records = []
    for df_idx, row in l1_df.iterrows():
        line_num = _clean_str(row.get("line_number"))
        from_str = _clean_str(row.get("from_equipment"))
        to_str   = _clean_str(row.get("to_equipment"))

        from_tags = _parse_tags(from_str)
        to_tags   = _parse_tags(to_str)
        all_tags  = from_tags | to_tags

        # Area code: read from dedicated column if configured, else extract from line number
        if area_col_override and area_col_override in row.index:
            area_code = _clean_str(row[area_col_override]) or None
        else:
            area_code = _extract_area_code(line_num, cn_cfg)

        records.append({
            "df_idx":               df_idx,
            "line_number":          line_num,
            "from_str":             from_str,
            "to_str":               to_str,
            "from_tags":            from_tags,
            "to_tags":              to_tags,
            "all_tags":             all_tags,
            "temperature":          _safe_float(row.get("design_temperature")),
            "has_expansion_joint":  _is_true(_clean_str(row.get("expansion_joint", ""))),
            "area_code":            area_code,
            "missing_data":         (not from_str or not to_str),
            "equip_types":          {t: _get_equip_type(t, equip_prefixes) for t in all_tags},
        })
    return records


# ─────────────────────────────────────────────────────────────────────────────
# Step 2 — Group lines into CNs via boundary rules
# ─────────────────────────────────────────────────────────────────────────────

def _group_into_cns(lines: list[dict], cn_cfg: dict) -> list[dict]:
    """
    Apply CN boundary rules in priority order and return a list of group dicts.

    Boundary priority:
      1. Missing equipment data → isolated, REVIEW-MISSING-DATA
      2. Expansion joint → isolated, hard split
      3. Pump-first pass → each unique pump tag gets its own CN
      4. Graph-based grouping for remaining lines
         4a. Area boundary split → REVIEW-AREA-CONFLICT if cross-area
    """
    equip_prefixes = cn_cfg.get("equipment_type_prefixes", {})
    pump_equip_types = {"centrifugal_pump", "reciprocating_pump"}
    n = len(lines)
    assigned = set()
    groups = []

    # ── Boundary 1: missing equipment data ───────────────────────────────────
    for i in range(n):
        if lines[i]["missing_data"]:
            groups.append(_make_group(
                indices=[i],
                group_type="unassigned",
                notes=["Missing FROM or TO equipment tag — connectivity cannot be determined"],
                area_conflict=False,
            ))
            assigned.add(i)

    # ── Boundary 2: expansion joints ─────────────────────────────────────────
    for i in range(n):
        if i in assigned:
            continue
        if lines[i]["has_expansion_joint"]:
            groups.append(_make_group(
                indices=[i],
                group_type="expansion_joint",
                notes=["Expansion joint present — hard CN boundary; no thermal force transfer across this point"],
                area_conflict=False,
            ))
            assigned.add(i)

    # ── Boundary 3: pump-first pass ───────────────────────────────────────────
    pump_tag_to_indices: dict[str, list[int]] = defaultdict(list)
    for i in range(n):
        if i in assigned:
            continue
        for tag, equip_type in lines[i]["equip_types"].items():
            if equip_type in pump_equip_types:
                pump_tag_to_indices[tag].append(i)

    for pump_tag, candidate_indices in pump_tag_to_indices.items():
        unassigned = [i for i in candidate_indices if i not in assigned]
        if not unassigned:
            continue
        equip_type = _get_equip_type(pump_tag, equip_prefixes) or "pump"
        groups.append(_make_group(
            indices=unassigned,
            group_type=equip_type,
            notes=[
                f"Pump CN: all lines directly connecting to {pump_tag} grouped together.",
                f"Suction and discharge of {pump_tag} must be in the same Caesar II model.",
            ],
            area_conflict=False,
        ))
        for i in unassigned:
            assigned.add(i)

    # ── Boundary 4: graph-based grouping for remaining lines ──────────────────
    remaining = [i for i in range(n) if i not in assigned]

    # Build adjacency: two remaining lines are connected if they share any equipment tag
    tag_to_lines: dict[str, list[int]] = defaultdict(list)
    for i in remaining:
        for tag in lines[i]["all_tags"]:
            tag_to_lines[tag].append(i)

    adj: dict[int, set] = defaultdict(set)
    for tag, connected in tag_to_lines.items():
        for a in connected:
            for b in connected:
                if a != b:
                    adj[a].add(b)

    # BFS connected components
    visited: set[int] = set()
    for start in remaining:
        if start in visited:
            continue
        component: list[int] = []
        queue = deque([start])
        visited.add(start)
        while queue:
            node = queue.popleft()
            component.append(node)
            for neighbor in adj[node]:
                if neighbor not in visited:
                    visited.add(neighbor)
                    queue.append(neighbor)

        # Boundary 4a: split by area code
        area_sub_groups = _split_by_area(component, lines)
        cross_area = len(area_sub_groups) > 1

        for sub_group in area_sub_groups:
            # Determine group type based on connectivity
            has_equip = any(lines[i]["all_tags"] for i in sub_group)
            if cross_area:
                gtype = "area_conflict"
                notes = [
                    "Lines from different area/unit codes appear in the same connected network.",
                    "Auto-split by area code. Engineer must verify on P&ID whether physical "
                    "connection crosses area boundary.",
                ]
            elif len(sub_group) > 1:
                gtype = "equipment_system"
                shared_tags = _shared_equipment_tags(sub_group, lines)
                notes = [
                    f"Lines connected through shared equipment: {', '.join(sorted(shared_tags))}.",
                ]
            elif has_equip:
                gtype = "area_group"
                notes = [
                    "Single line with equipment connections but no adjacent Level 1 lines sharing "
                    "the same equipment. Standalone — engineer must confirm CN boundary on P&ID.",
                ]
            else:
                gtype = "area_group"
                notes = [
                    "No equipment connectivity data linking this line to other Level 1 lines. "
                    "Standalone — engineer must confirm CN boundary on P&ID.",
                ]

            groups.append(_make_group(
                indices=sub_group,
                group_type=gtype,
                notes=notes,
                area_conflict=cross_area,
            ))

    return groups


# ─────────────────────────────────────────────────────────────────────────────
# Step 3 — Number CNs, apply soft flags, build proposal dicts
# ─────────────────────────────────────────────────────────────────────────────

def _finalize_cn_assignments(
    groups: list[dict],
    lines: list[dict],
    cn_cfg: dict,
) -> list[dict]:
    project_code    = cn_cfg.get("project_code", "PROJ")
    cn_format       = cn_cfg.get("cn_number_format", "{project_code}-CN-{number:03d}")
    max_lines       = int(cn_cfg.get("max_lines_per_cn", 12))
    temp_delta_flag = float(cn_cfg.get("temp_delta_flag", 80))
    cn_ranges       = cn_cfg.get("cn_number_ranges", {})

    pump_start     = int(cn_ranges.get("centrifugal_pump_start", 1))
    comp_start     = int(cn_ranges.get("compressor_turbine_start", 100))
    equip_start    = int(cn_ranges.get("equipment_system_start", 200))
    area_start     = int(cn_ranges.get("area_group_start", 300))
    unassigned_start = int(cn_ranges.get("unassigned_start", 900))

    counters = {
        "pump":         pump_start,
        "equipment":    equip_start,
        "area":         area_start,
        "unassigned":   unassigned_start,
        "compressor":   comp_start,
    }

    proposals = []
    for group in groups:
        indices    = group["indices"]
        group_type = group["group_type"]
        notes      = list(group["notes"])
        area_conflict = group["area_conflict"]

        # Collect per-group data
        group_lines  = [lines[i] for i in indices]
        line_numbers = [l["line_number"] for l in group_lines]
        df_indices   = [l["df_idx"]      for l in group_lines]
        all_equip    = sorted({t for l in group_lines for t in l["all_tags"]})
        temps        = [l["temperature"] for l in group_lines if l["temperature"] is not None]
        area_codes   = sorted({l["area_code"] for l in group_lines if l["area_code"]})
        line_count   = len(indices)

        min_temp  = min(temps) if temps else None
        max_temp  = max(temps) if temps else None
        delta_t   = (max_temp - min_temp) if (min_temp is not None and max_temp is not None) else None

        # ── Determine CN number range ─────────────────────────────────────────
        if group_type in ("unassigned",):
            cn_num = counters["unassigned"]
            counters["unassigned"] += 1
        elif group_type in ("centrifugal_pump", "reciprocating_pump"):
            cn_num = counters["pump"]
            counters["pump"] += 1
        elif group_type in ("compressor", "turbine"):
            cn_num = counters["compressor"]
            counters["compressor"] += 1
        elif group_type in ("equipment_system", "area_conflict"):
            cn_num = counters["equipment"]
            counters["equipment"] += 1
        else:
            cn_num = counters["area"]
            counters["area"] += 1

        cn_number = cn_format.format(project_code=project_code, number=cn_num)

        # ── Determine review flag ─────────────────────────────────────────────
        if group_type == "unassigned":
            review_flag = "[REVIEW-MISSING-DATA]"
            notes.append("CN assigned to 9XX range. Engineer must determine correct grouping from P&ID.")
        elif area_conflict:
            review_flag = "[REVIEW-AREA-CONFLICT]"
        elif delta_t is not None and delta_t > temp_delta_flag:
            review_flag = "[REVIEW-TEMPERATURE]"
            notes.append(
                f"Temperature delta = {delta_t:.0f}°C between connected lines (threshold: {temp_delta_flag:.0f}°C). "
                f"Engineer must confirm these lines belong in the same model."
            )
        elif line_count > max_lines:
            review_flag = "[REVIEW-MODEL-SIZE]"
            notes.append(
                f"CN contains {line_count} lines (limit: {max_lines}). "
                f"Engineer must propose a split point based on P&ID routing."
            )
        elif line_count == 1 and group_type not in ("centrifugal_pump", "reciprocating_pump", "expansion_joint"):
            review_flag = "[REVIEW-MANUAL]"
            notes.append(
                "Single line in CN with no confirmed adjacent Level 1 connectivity. "
                "Engineer must confirm boundary on P&ID."
            )
        else:
            review_flag = "[AUTO-CONFIRMED]"

        # ── Build grouping reason text ────────────────────────────────────────
        grouping_reason = " | ".join(notes)

        proposals.append({
            "cn_number":        cn_number,
            "review_flag":      review_flag,
            "line_numbers":     line_numbers,
            "line_df_indices":  df_indices,
            "equipment_tags":   all_equip,
            "area_codes":       area_codes,
            "min_temperature":  min_temp,
            "max_temperature":  max_temp,
            "delta_t":          delta_t,
            "line_count":       line_count,
            "grouping_reason":  grouping_reason,
            "group_type":       group_type,
        })

    # Sort proposals by CN number string for deterministic output
    proposals.sort(key=lambda p: p["cn_number"])
    return proposals


# ─────────────────────────────────────────────────────────────────────────────
# Helper: group factory
# ─────────────────────────────────────────────────────────────────────────────

def _make_group(
    indices: list[int],
    group_type: str,
    notes: list[str],
    area_conflict: bool,
) -> dict:
    return {
        "indices":      indices,
        "group_type":   group_type,
        "notes":        notes,
        "area_conflict": area_conflict,
    }


# ─────────────────────────────────────────────────────────────────────────────
# Helper: area-based split
# ─────────────────────────────────────────────────────────────────────────────

def _split_by_area(component: list[int], lines: list[dict]) -> list[list[int]]:
    """
    Split a connected component into per-area-code sub-groups.
    Lines with no area code share a special '_UNKNOWN_' bucket.
    Returns a list of sub-groups (each sub-group is a list of indices).
    """
    buckets: dict[str, list[int]] = defaultdict(list)
    for i in component:
        key = lines[i]["area_code"] or "_UNKNOWN_"
        buckets[key].append(i)

    if len(buckets) == 1:
        return [component]
    return list(buckets.values())


# ─────────────────────────────────────────────────────────────────────────────
# Helper: find equipment tags shared by multiple lines in a group
# ─────────────────────────────────────────────────────────────────────────────

def _shared_equipment_tags(indices: list[int], lines: list[dict]) -> set[str]:
    tag_counts: dict[str, int] = defaultdict(int)
    for i in indices:
        for tag in lines[i]["all_tags"]:
            tag_counts[tag] += 1
    return {t for t, cnt in tag_counts.items() if cnt > 1}


# ─────────────────────────────────────────────────────────────────────────────
# Helper: equipment type lookup
# ─────────────────────────────────────────────────────────────────────────────

def _get_equip_type(tag: str, equip_prefixes: dict) -> Optional[str]:
    """Return the equipment type name for a tag, or None if unrecognised."""
    tag_upper = tag.upper()
    for equip_type, prefixes in equip_prefixes.items():
        for prefix in prefixes:
            if tag_upper.startswith(prefix.upper()):
                return equip_type
    return None


# ─────────────────────────────────────────────────────────────────────────────
# Helper: area code extraction from line number string
# ─────────────────────────────────────────────────────────────────────────────

def _extract_area_code(line_number: str, cn_cfg: dict) -> Optional[str]:
    """
    Extract area code from line number using the configured extraction rule.
    Rule format: "characters_N_to_M_of_line_number" (1-based, inclusive).
    """
    rule = cn_cfg.get("area_code_extraction", "")
    if not rule or not line_number:
        return None
    m = re.match(r"characters_(\d+)_to_(\d+)_of_line_number", rule)
    if not m:
        return None
    start = int(m.group(1)) - 1  # convert to 0-based
    end   = int(m.group(2))      # exclusive in Python slice
    if len(line_number) < end:
        return None
    code = line_number[start:end].strip()
    return code if code else None


# ─────────────────────────────────────────────────────────────────────────────
# Helpers: string / numeric parsing
# ─────────────────────────────────────────────────────────────────────────────

_EMPTY_VALUES = {"", "nan", "none", "-", "n/a", "tbd", "na"}


def _clean_str(value) -> str:
    """Return stripped string; convert None/NaN to empty string."""
    if value is None:
        return ""
    import math
    try:
        if math.isnan(float(value)):
            return ""
    except (TypeError, ValueError):
        pass
    s = str(value).strip()
    return "" if s.lower() in _EMPTY_VALUES else s


def _parse_tags(raw: str) -> set[str]:
    """Split a FROM/TO cell on / & , and newline separators; return non-empty tags."""
    if not raw:
        return set()
    parts = re.split(r"[/&,\n]", raw)
    result = set()
    for p in parts:
        p = p.strip()
        if p and p.lower() not in _EMPTY_VALUES:
            result.add(p)
    return result


def _safe_float(value) -> Optional[float]:
    """Convert a cell value to float, stripping °C suffixes. Returns None on failure."""
    if value is None:
        return None
    s = str(value).strip().rstrip("°C").rstrip("°").rstrip("C").strip()
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def _is_true(value: str) -> bool:
    """Return True if the string represents a truthy boolean (Yes/TRUE/1/X)."""
    return value.lower() in {"yes", "true", "1", "x"}

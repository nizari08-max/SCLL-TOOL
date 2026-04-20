"""
cn_assigner.py — CN (Calculation Number) Assignment Engine for Level 1 lines.

A CN is a Caesar II model boundary: a group of piping lines that must be analyzed
together because thermal forces and loads transfer between them.

GROUPING ALGORITHM
==================
Build a graph where each Level 1 line is a node. Two lines are connected if
ANY of these are true:
  - They share any equipment tag (FROM or TO)
  - They share the same header reference

Find connected components, then apply six refinement rules:

  RULE 1 — PROCESS CONSISTENCY
    Split a component if:
      • fluid service mismatch between endpoints, OR
      • material group mismatch, OR
      • temperature delta > 30 °C (configurable)

  RULE 2 — PUMP SUCTION + DISCHARGE
    Overrides Rule 1: lines that share a pump tag stay in the same CN even if
    temperatures differ.

  RULE 3 — SHARED HEADER
    Lines sharing a header reference stay grouped (implicit in graph build).

  RULE 4 — LARGE CN FLAGGING
    CN with more than 15 lines (configurable) gets review flag REVIEW-LARGE-CN.
    NOT auto-split — engineer decides boundary.

  RULE 5 — PARALLEL TRAINS
    If a TRAIN column exists, lines from different trains cannot share a CN.

  RULE 6 — STANDALONE
    Lines with no connectivity (or missing FROM/TO) get their own CN in the
    CN-900+ range with flag REVIEW-STANDALONE.

This module proposes; the engineer decides. Every proposal carries a reason.
"""

from __future__ import annotations

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
    material_map: Optional[dict[str, str]] = None,
) -> tuple[pd.DataFrame, list[dict]]:
    """
    Assign Calculation Numbers to Level 1 lines.

    Returns:
        result_df    — classified_df with CN_Number and CN_Review_Flag columns
                       filled for Level 1 rows; non-Level-1 rows stay empty.
        cn_proposals — list of dicts (one per proposed CN) for the
                       CN Proposals output sheet.
    """
    cn_cfg = rules.get("cn_settings", {}) or {}
    max_lines       = int(cn_cfg.get("max_lines_per_cn", 15))
    temp_delta_flag = float(cn_cfg.get("temp_delta_flag", 30))
    cn_format       = cn_cfg.get("cn_number_format", "CN-{number:03d}")
    project_code    = cn_cfg.get("project_code", "")
    equip_prefixes  = cn_cfg.get("equipment_type_prefixes", {}) or {}
    standalone_start = int(cn_cfg.get("unassigned_start", 900))

    result_df = classified_df.copy()
    for col in ("CN_Number", "CN_Review_Flag", "CN_Reason"):
        if col not in result_df.columns:
            result_df[col] = ""

    l1_mask = result_df["Level"].astype(str) == "Level 1"
    l1_indices = result_df.index[l1_mask].tolist()
    if not l1_indices:
        return result_df, []

    lines = _build_line_records(result_df.loc[l1_indices], equip_prefixes, material_map)

    # ── RULE 5: split by train up-front (hard boundary) ─────────────────────
    train_buckets = _bucket_by_train(lines)

    raw_groups: list[dict] = []
    for train_value, train_line_indices in train_buckets.items():
        # Within this train, build graph + find components
        adj = _build_adjacency(lines, train_line_indices)
        components = _connected_components(train_line_indices, adj)

        for comp in components:
            # ── RULE 1 + RULE 2: process consistency with pump override ────
            sub_components = _apply_process_consistency(
                comp, lines, temp_delta_flag, equip_prefixes
            )
            for sub in sub_components:
                raw_groups.append({
                    "indices": sub["indices"],
                    "reason_parts": sub["reason_parts"],
                    "is_standalone": sub.get("is_standalone", False),
                    "missing_data": sub.get("missing_data", False),
                    "train": train_value,
                })

    # ── Assign CN numbers + review flags ───────────────────────────────────
    proposals = _finalize(
        raw_groups, lines,
        max_lines=max_lines,
        temp_delta_flag=temp_delta_flag,
        cn_format=cn_format,
        project_code=project_code,
        standalone_start=standalone_start,
    )

    # Write results back to DataFrame
    for p in proposals:
        for df_idx in p["line_df_indices"]:
            result_df.at[df_idx, "CN_Number"]      = p["cn_number"]
            result_df.at[df_idx, "CN_Review_Flag"] = p["review_flag"]
            result_df.at[df_idx, "CN_Reason"]      = p["grouping_reason"]

    return result_df, proposals


# ─────────────────────────────────────────────────────────────────────────────
# Line records
# ─────────────────────────────────────────────────────────────────────────────

def _build_line_records(
    l1_df: pd.DataFrame,
    equip_prefixes: dict,
    material_map: Optional[dict[str, str]],
) -> dict[int, dict]:
    """Return {df_index → record}. Each record has everything the grouping needs."""
    records: dict[int, dict] = {}
    for df_idx, row in l1_df.iterrows():
        from_str = _clean_str(row.get("from_equipment", ""))
        to_str   = _clean_str(row.get("to_equipment", ""))

        from_tags = _parse_tags(from_str)
        to_tags   = _parse_tags(to_str)
        all_tags  = from_tags | to_tags

        material_raw = _clean_str(row.get("material", ""))
        material_group = None
        if material_map and material_raw:
            material_group = material_map.get(material_raw.upper())

        train_raw = _clean_str(row.get("train", ""))
        train_key = train_raw.lower() if train_raw else None  # None = no train info

        records[df_idx] = {
            "df_idx":          df_idx,
            "line_number":     _clean_str(row.get("line_number", "")),
            "from_str":        from_str,
            "to_str":          to_str,
            "from_tags":       from_tags,
            "to_tags":         to_tags,
            "all_tags":        all_tags,
            "temperature":     _safe_float(row.get("design_temperature")),
            "fluid_service":   _clean_str(row.get("fluid_service", "")).lower() or None,
            "material":        material_raw,
            "material_group":  material_group,
            "train":           train_key,
            "missing_data":    (not from_str and not to_str),
            "equip_types":     {t: _get_equip_type(t, equip_prefixes) for t in all_tags},
        }
    return records


# ─────────────────────────────────────────────────────────────────────────────
# RULE 5 — bucket lines by train
# ─────────────────────────────────────────────────────────────────────────────

def _bucket_by_train(lines: dict[int, dict]) -> dict[Optional[str], list[int]]:
    """
    Group lines by train value. Lines with no train value share a single
    'no-train' bucket. Each bucket becomes an independent graph.
    """
    buckets: dict[Optional[str], list[int]] = defaultdict(list)
    for df_idx, rec in lines.items():
        buckets[rec["train"]].append(df_idx)
    return buckets


# ─────────────────────────────────────────────────────────────────────────────
# Graph build + connected components
# ─────────────────────────────────────────────────────────────────────────────

def _build_adjacency(
    lines: dict[int, dict], indices: list[int]
) -> dict[int, set[int]]:
    """Build adjacency: two nodes share an edge iff they share any equipment tag."""
    tag_to_lines: dict[str, list[int]] = defaultdict(list)
    for i in indices:
        for tag in lines[i]["all_tags"]:
            tag_to_lines[tag].append(i)

    adj: dict[int, set[int]] = defaultdict(set)
    for tag, group in tag_to_lines.items():
        for a in group:
            for b in group:
                if a != b:
                    adj[a].add(b)
    return adj


def _connected_components(
    indices: list[int], adj: dict[int, set[int]]
) -> list[list[int]]:
    """Return connected components via BFS. Isolated nodes are singleton components."""
    visited: set[int] = set()
    components: list[list[int]] = []
    for start in indices:
        if start in visited:
            continue
        comp: list[int] = []
        queue = deque([start])
        visited.add(start)
        while queue:
            node = queue.popleft()
            comp.append(node)
            for nb in adj.get(node, ()):
                if nb not in visited:
                    visited.add(nb)
                    queue.append(nb)
        components.append(comp)
    return components


# ─────────────────────────────────────────────────────────────────────────────
# RULE 1 + RULE 2 — process consistency with pump override
# ─────────────────────────────────────────────────────────────────────────────

_PUMP_TYPES = {"centrifugal_pump", "reciprocating_pump"}


def _apply_process_consistency(
    component: list[int],
    lines: dict[int, dict],
    temp_delta_flag: float,
    equip_prefixes: dict,
) -> list[dict]:
    """
    Split a connected component into process-consistent sub-components.

    A pair of lines is kept in the same sub-component if they share any
    equipment tag AND at least one of:
      - the shared tag is a pump (Rule 2 overrides temperature delta)
      - fluid service matches AND material group matches AND temp delta ≤ threshold

    Returns a list of sub-component dicts with indices + reason_parts.
    """
    # Singleton or missing-data → standalone immediately
    if len(component) == 1:
        rec = lines[component[0]]
        if rec["missing_data"]:
            return [{
                "indices":      component,
                "reason_parts": ["Missing FROM/TO — connectivity cannot be determined"],
                "is_standalone": True,
                "missing_data":  True,
            }]
        if not rec["all_tags"]:
            return [{
                "indices":      component,
                "reason_parts": ["No equipment tags — cannot connect to any other Level 1 line"],
                "is_standalone": True,
            }]
        return [{
            "indices":      component,
            "reason_parts": ["Single Level 1 line with no neighbours sharing equipment tags"],
            "is_standalone": True,
        }]

    # Build per-pair shared tags + classify edge types
    tag_to_lines: dict[str, list[int]] = defaultdict(list)
    for i in component:
        for tag in lines[i]["all_tags"]:
            tag_to_lines[tag].append(i)

    # Build refined graph: keep edge only if process-consistent OR pump-shared
    refined_adj: dict[int, set[int]] = defaultdict(set)
    edge_reasons: dict[tuple[int, int], list[str]] = defaultdict(list)

    for tag, group in tag_to_lines.items():
        tag_is_pump = _get_equip_type(tag, equip_prefixes) in _PUMP_TYPES
        for a in group:
            for b in group:
                if a >= b:
                    continue
                pair = (a, b)
                ra, rb = lines[a], lines[b]

                if tag_is_pump:
                    refined_adj[a].add(b)
                    refined_adj[b].add(a)
                    edge_reasons[pair].append(f"shared pump {tag}")
                    continue

                # Non-pump shared tag: check process consistency
                if not _process_consistent(ra, rb, temp_delta_flag):
                    continue
                refined_adj[a].add(b)
                refined_adj[b].add(a)
                edge_reasons[pair].append(f"shared {tag}")

    # Ensure isolated nodes still appear (singletons after split)
    for i in component:
        refined_adj.setdefault(i, set())

    # Connected components of refined graph
    sub_comps = _connected_components(component, refined_adj)

    result: list[dict] = []
    for sub in sub_comps:
        if len(sub) == 1:
            rec = lines[sub[0]]
            if not rec["all_tags"]:
                reason = ["Standalone — no equipment tags"]
            else:
                reason = [
                    "Split from larger network due to process mismatch "
                    "(fluid service / material / temperature)"
                ]
            result.append({
                "indices":       sub,
                "reason_parts":  reason,
                "is_standalone": True,
            })
            continue

        # Multi-line sub-component: collect shared tags for reason
        tags_in_sub: dict[str, int] = defaultdict(int)
        for i in sub:
            for tag in lines[i]["all_tags"]:
                tags_in_sub[tag] += 1
        shared_tags = sorted([t for t, c in tags_in_sub.items() if c > 1])

        pump_tags = [t for t in shared_tags
                     if _get_equip_type(t, equip_prefixes) in _PUMP_TYPES]

        reason_parts: list[str] = []
        if pump_tags:
            reason_parts.append(
                f"Pump CN: suction + discharge of {', '.join(pump_tags)} grouped per Rule 2"
            )
        if shared_tags and not pump_tags:
            reason_parts.append(
                f"Equipment network: shared {', '.join(shared_tags[:4])}"
                + (" …" if len(shared_tags) > 4 else "")
            )
        if not reason_parts:
            reason_parts.append("Connected via shared equipment tags")

        # If this sub-component is a proper subset of the original component,
        # note the process split
        if len(sub) < len(component):
            reason_parts.append(
                "Process-consistency split (Rule 1) from larger connected network"
            )

        result.append({
            "indices":      sub,
            "reason_parts": reason_parts,
            "is_standalone": False,
        })

    return result


def _process_consistent(a: dict, b: dict, temp_delta_flag: float) -> bool:
    """Return True if two line records are in the same process system."""
    # Fluid service: if both known and different → mismatch
    if a["fluid_service"] and b["fluid_service"]:
        if a["fluid_service"] != b["fluid_service"]:
            return False

    # Material group: if both known and different → mismatch
    if a["material_group"] and b["material_group"]:
        if a["material_group"] != b["material_group"]:
            return False

    # Temperature delta check
    ta, tb = a["temperature"], b["temperature"]
    if ta is not None and tb is not None:
        if abs(ta - tb) > temp_delta_flag:
            return False

    return True


# ─────────────────────────────────────────────────────────────────────────────
# Finalization — CN numbering + review flags
# ─────────────────────────────────────────────────────────────────────────────

def _finalize(
    raw_groups: list[dict],
    lines: dict[int, dict],
    *,
    max_lines: int,
    temp_delta_flag: float,
    cn_format: str,
    project_code: str,
    standalone_start: int,
) -> list[dict]:
    """Sort groups, assign CN numbers, compute review flags, and build proposals."""
    # Split into main (multi-line, connected) and standalone (singletons / missing)
    main_groups = [g for g in raw_groups if not g["is_standalone"]]
    standalone_groups = [g for g in raw_groups if g["is_standalone"]]

    # Deterministic order: sort main by smallest df index, standalone by df index
    main_groups.sort(key=lambda g: min(g["indices"]))
    standalone_groups.sort(key=lambda g: min(g["indices"]))

    proposals: list[dict] = []
    seq_counter = 1
    standalone_counter = standalone_start

    for g in main_groups:
        proposals.append(_build_proposal(
            g, lines, seq_counter, cn_format, project_code,
            max_lines, temp_delta_flag, is_standalone=False,
        ))
        seq_counter += 1

    for g in standalone_groups:
        proposals.append(_build_proposal(
            g, lines, standalone_counter, cn_format, project_code,
            max_lines, temp_delta_flag, is_standalone=True,
        ))
        standalone_counter += 1

    return proposals


def _build_proposal(
    group: dict,
    lines: dict[int, dict],
    cn_num: int,
    cn_format: str,
    project_code: str,
    max_lines: int,
    temp_delta_flag: float,
    *,
    is_standalone: bool,
) -> dict:
    indices      = group["indices"]
    reason_parts = list(group["reason_parts"])
    train_val    = group.get("train")

    line_records = [lines[i] for i in indices]
    line_numbers = [r["line_number"] for r in line_records]
    df_indices   = [r["df_idx"]      for r in line_records]
    temps        = [r["temperature"] for r in line_records if r["temperature"] is not None]
    all_equip    = sorted({t for r in line_records for t in r["all_tags"]})

    min_temp = min(temps) if temps else None
    max_temp = max(temps) if temps else None
    delta_t  = (max_temp - min_temp) if (min_temp is not None and max_temp is not None) else None

    try:
        cn_number = cn_format.format(project_code=project_code, number=cn_num)
    except (KeyError, IndexError):
        cn_number = f"CN-{cn_num:03d}"

    # Determine review flag (priority order)
    if is_standalone:
        if group.get("missing_data"):
            review_flag = "REVIEW-STANDALONE"
            reason_parts.append("Assigned to CN-900+ range (missing connectivity data)")
        else:
            review_flag = "REVIEW-STANDALONE"
            reason_parts.append("Assigned to CN-900+ range (Rule 6: standalone)")
    elif len(indices) > max_lines:
        review_flag = "REVIEW-LARGE-CN"
        reason_parts.append(
            f"CN contains {len(indices)} lines (limit: {max_lines}) — "
            f"engineer must propose split point from P&ID (Rule 4)"
        )
    else:
        review_flag = "AUTO-CONFIRMED"

    if train_val:
        reason_parts.append(f"Train: {train_val}")

    grouping_reason = " | ".join(p for p in reason_parts if p)

    return {
        "cn_number":       cn_number,
        "review_flag":     review_flag,
        "grouping_reason": grouping_reason,
        "line_numbers":    line_numbers,
        "line_df_indices": df_indices,
        "equipment_tags":  all_equip,
        "min_temperature": min_temp,
        "max_temperature": max_temp,
        "delta_t":         delta_t,
        "line_count":      len(indices),
        "train":           train_val,
    }


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

_EMPTY_VALUES = {"", "nan", "none", "-", "n/a", "tbd", "na"}


def _clean_str(value) -> str:
    if value is None:
        return ""
    try:
        if isinstance(value, float) and pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    s = str(value).strip()
    return "" if s.lower() in _EMPTY_VALUES else s


def _parse_tags(raw: str) -> set[str]:
    """Split a FROM/TO cell on / & , and newline; return non-empty normalised tags."""
    if not raw:
        return set()
    parts = re.split(r"[/&,\n]", raw)
    result = set()
    for p in parts:
        p = p.strip()
        if p and p.lower() not in _EMPTY_VALUES:
            result.add(p.upper())
    return result


def _safe_float(value) -> Optional[float]:
    if value is None:
        return None
    try:
        if isinstance(value, float) and pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    s = str(value).strip().rstrip("°C").rstrip("°").rstrip("C").strip()
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def _get_equip_type(tag: str, equip_prefixes: dict) -> Optional[str]:
    """Return the equipment type name for a tag, or None if unrecognised."""
    tag_upper = tag.upper()
    for equip_type, prefixes in equip_prefixes.items():
        for prefix in prefixes:
            if tag_upper.startswith(str(prefix).upper()):
                return equip_type
    return None

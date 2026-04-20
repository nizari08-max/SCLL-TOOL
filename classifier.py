"""
classifier.py — 5-step deterministic classification engine.

Waterfall (strict priority; first match wins, no later step overrides an earlier one):

  STEP 1 — FRP / non-metallic material            → Level 1
  STEP 2 — Exception flags (relief line, jacketed, vibration, …)  → Level 1
  STEP 3 — Strain-sensitive equipment (Chart 3)
  STEP 4 — Material chart lookup (Chart 1 / 2 / 4)
  STEP 5 — Missing data → Level "?" + Data_Quality_Flag = "MISSING: …"

Graceful handling of missing columns:
  - Exception flag rules whose 'flag' column is absent from the file are silently skipped.
  - Missing size / temperature / material → row marked NEEDS REVIEW, not rejected.
  - Missing FROM/TO → Step 3 short-circuits, classification proceeds with Step 4.

This module performs no I/O. All thresholds come from the rules dict.
"""

from __future__ import annotations

import pandas as pd


# ─────────────────────────────────────────────────────────────────────────────
# Material group resolution
# ─────────────────────────────────────────────────────────────────────────────

def resolve_material_group(pipe_class: str, material_map: dict[str, str]) -> str:
    if not pipe_class or pd.isna(pipe_class):
        return "UNKNOWN"
    return material_map.get(str(pipe_class).strip().upper(), "UNKNOWN")


# ─────────────────────────────────────────────────────────────────────────────
# Equipment tag parsing
# ─────────────────────────────────────────────────────────────────────────────

def _parse_tag_fragments(tag_value) -> list[str]:
    if not tag_value or (isinstance(tag_value, float) and pd.isna(tag_value)):
        return []
    raw = str(tag_value).strip()
    for sep in ("/", "&", ";", ","):
        raw = raw.replace(sep, "|")
    return [t.strip() for t in raw.split("|") if t.strip()]


def _match_tag_prefix(tag: str, tag_patterns: list[dict]) -> tuple[bool, str]:
    tag_upper = tag.strip().upper()
    for p in tag_patterns:
        if tag_upper.startswith(str(p["prefix"]).upper()):
            return True, p["type"]
    return False, ""


def _strain_sensitive_prefix(tag_value, rules: dict) -> tuple[bool, str]:
    patterns = rules["strain_sensitive_equipment"].get("tag_patterns", [])
    priority = ["centrifugal_pump", "reciprocating_pump", "compressor",
                "turbine", "heater", "air_cooler"]
    found: list[str] = []
    for t in _parse_tag_fragments(tag_value):
        matched, equip_type = _match_tag_prefix(t, patterns)
        if matched:
            found.append(equip_type)
    if not found:
        return False, ""
    for p in priority:
        if p in found:
            return True, p
    return True, found[0]


def _strain_sensitive_keyword(tag_value, rules: dict) -> tuple[bool, str]:
    if not tag_value or (isinstance(tag_value, float) and pd.isna(tag_value)):
        return False, ""
    tag_upper = str(tag_value).strip().upper()
    for entry in rules["strain_sensitive_equipment"].get("keyword_patterns", []):
        for kw in entry.get("keywords", []):
            if str(kw).upper() in tag_upper:
                return True, entry["type"]
    return False, ""


def is_strain_sensitive_equipment(tag_value, rules: dict) -> tuple[bool, str]:
    mode = rules["strain_sensitive_equipment"].get("detection_mode", "tag_prefix")
    if mode == "keyword":
        return _strain_sensitive_keyword(tag_value, rules)
    return _strain_sensitive_prefix(tag_value, rules)


def _non_strain_sensitive_keyword(tag_value, rules: dict) -> bool:
    if not tag_value or (isinstance(tag_value, float) and pd.isna(tag_value)):
        return False
    tag_upper = str(tag_value).strip().upper()
    kws = rules["strain_sensitive_equipment"].get("non_strain_sensitive_keywords", [])
    return any(str(k).upper() in tag_upper for k in kws)


def _has_equipment_connection(tag_value, rules: dict) -> bool:
    if not tag_value or (isinstance(tag_value, float) and pd.isna(tag_value)):
        return False
    s = str(tag_value).strip()
    if not s or s.lower() == "nan":
        return False
    mode = rules["strain_sensitive_equipment"].get("detection_mode", "tag_prefix")
    if mode == "keyword":
        sensitive, _ = _strain_sensitive_keyword(tag_value, rules)
        return sensitive or _non_strain_sensitive_keyword(tag_value, rules)
    # prefix mode: any non-empty tag counts as equipment
    return True


# ─────────────────────────────────────────────────────────────────────────────
# Numeric parsing (tolerates "60/-20", "150°C")
# ─────────────────────────────────────────────────────────────────────────────

def _to_float(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip().rstrip("°C").rstrip("°").rstrip("C").strip()
    if not s:
        return None
    if "/" in s:  # compound max/min temp like "60/-20"
        nums = []
        for p in s.split("/"):
            p = p.strip().rstrip("°C").rstrip("°").rstrip("C").strip()
            try:
                nums.append(float(p))
            except (ValueError, TypeError):
                pass
        return max(nums) if nums else None
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


# ─────────────────────────────────────────────────────────────────────────────
# Step 2 — Exception flags
# ─────────────────────────────────────────────────────────────────────────────

def _eval_extra(row_val, condition: dict | None) -> bool:
    if condition is None:
        return True
    v = _to_float(row_val)
    if v is None:
        return False
    threshold = float(condition["value"])
    op = condition["operator"]
    return {">": v > threshold, ">=": v >= threshold,
            "<": v < threshold, "<=": v <= threshold}.get(op, False)


def _check_exception_flags(row: pd.Series, rules: dict) -> tuple[bool, str]:
    true_vals = {str(v).strip().lower() for v in rules.get("true_values", [])}
    for exc in rules.get("exception_flags", []):
        flag_col = exc["flag"]
        # If the flag column doesn't exist in this file, skip the rule silently
        if flag_col not in row.index:
            continue
        cell = row.get(flag_col)
        if cell is None or (isinstance(cell, float) and pd.isna(cell)):
            continue
        if str(cell).strip().lower() not in true_vals:
            continue
        extra = exc.get("extra_condition")
        if extra is not None:
            if extra["field"] not in row.index:
                continue
            if not _eval_extra(row.get(extra["field"]), extra):
                continue
        return True, exc["reason"]
    return False, ""


# ─────────────────────────────────────────────────────────────────────────────
# Chart lookups
# ─────────────────────────────────────────────────────────────────────────────

def _lookup_chart(chart_name: str, size: float, temp: float, rules: dict) -> str:
    for band in rules[chart_name]["size_bands"]:
        if size <= float(band["max_size"]):
            if temp > float(band["l1_threshold"]):
                return "Level 1"
            if temp > float(band["l2_threshold"]):
                return "Level 2"
            return "Level 3"
    return "Level 3"


def _apply_chart_3(size: float, temp: float, equip_type: str, rules: dict) -> tuple[str, str]:
    for rule in rules["strain_sensitive_equipment"]["chart_3"]:
        ctype = rule["condition_type"]
        if ctype == "equip_type" and equip_type in rule["equip_types"]:
            return rule["result"], rule["reason"]
        if ctype == "temp_threshold":
            op, val = rule["temp_operator"], float(rule["temp_value"])
            if (op == ">="  and temp >= val) or (op == ">" and temp > val):
                return rule["result"], rule["reason"]
        if ctype == "size_temp":
            s_op, s_val = rule["size_operator"], float(rule["size_value"])
            t_op, t_val = rule["temp_operator"], float(rule["temp_value"])
            s_ok = (s_op == "<=" and size <= s_val) or (s_op == ">" and size > s_val)
            t_ok = (t_op == ">=" and temp >= t_val) or (t_op == "<" and temp < t_val)
            if s_ok and t_ok:
                return rule["result"], rule["reason"]
        if ctype == "fallback":
            return rule["result"], rule["reason"]
    return "Level 3", "Strain-sensitive equipment — fallback Level 3"


# ─────────────────────────────────────────────────────────────────────────────
# Row classifier
# ─────────────────────────────────────────────────────────────────────────────

def classify_row(row: pd.Series, material_map: dict, rules: dict) -> dict:
    """
    Returns {level, reason, data_quality}.
    level         — "Level 1" | "Level 2" | "Level 3" | "" (unclassifiable)
    reason        — human-readable explanation of which rule fired
    data_quality  — "" | "MISSING: <fields>" | "AMBIGUOUS"
    """

    # ── Check required fields present ────────────────────────────────────────
    raw_size     = row.get("size")
    raw_temp     = row.get("design_temperature")
    raw_material = row.get("material")

    size     = _to_float(raw_size)
    temp     = _to_float(raw_temp)
    material_str = ""
    if raw_material is not None and not (isinstance(raw_material, float) and pd.isna(raw_material)):
        material_str = str(raw_material).strip()

    missing = []
    if size is None:
        missing.append("Size")
    if temp is None:
        missing.append("Design Temperature")
    if not material_str:
        missing.append("Material")

    if missing:
        return {
            "level":        "",
            "reason":       f"Missing required field(s): {', '.join(missing)} — cannot classify",
            "data_quality": f"MISSING: {', '.join(missing)}",
        }

    # ── STEP 1: FRP / non-metallic → Level 1 ─────────────────────────────────
    material_group = resolve_material_group(material_str, material_map)
    if material_group == "FRP":
        return {
            "level":        "Level 1",
            "reason":       f"Non-metallic material ({material_str}) → Level 1 (Step 1)",
            "data_quality": "",
        }
    if material_group == "UNKNOWN":
        return {
            "level":        "",
            "reason":       f"Material code '{material_str}' not in material_mapping.yaml — cannot classify",
            "data_quality": "AMBIGUOUS",
        }

    # ── STEP 2: Exception flags ──────────────────────────────────────────────
    triggered, reason = _check_exception_flags(row, rules)
    if triggered:
        return {"level": "Level 1", "reason": reason + " (Step 2)", "data_quality": ""}

    # ── STEP 3: Strain-sensitive equipment (Chart 3) ─────────────────────────
    from_tag = row.get("from_equipment") if "from_equipment" in row.index else None
    to_tag   = row.get("to_equipment")   if "to_equipment"   in row.index else None

    from_sens, from_type = is_strain_sensitive_equipment(from_tag, rules)
    to_sens,   to_type   = is_strain_sensitive_equipment(to_tag, rules)

    if from_sens or to_sens:
        priority = ["centrifugal_pump", "reciprocating_pump", "compressor",
                    "turbine", "heater", "air_cooler"]
        equip_type = from_type if from_sens else to_type
        for p in priority:
            if p in (from_type, to_type):
                equip_type = p
                break
        level, step3_reason = _apply_chart_3(size, temp, equip_type, rules)
        return {"level": level, "reason": step3_reason + " (Step 3)", "data_quality": ""}

    # ── STEP 4: Material chart lookup ────────────────────────────────────────
    chart_name = rules["material_chart_map"].get(material_group)
    if not chart_name:
        return {
            "level":        "",
            "reason":       f"No chart defined for material group '{material_group}'",
            "data_quality": "AMBIGUOUS",
        }

    rank = {"Level 1": 1, "Level 2": 2, "Level 3": 3}
    mat_level = _lookup_chart(chart_name, size, temp, rules)

    has_equip = (_has_equipment_connection(from_tag, rules)
                 or _has_equipment_connection(to_tag, rules))

    if has_equip:
        c4_level = _lookup_chart("chart_4", size, temp, rules)
        if rank[c4_level] <= rank[mat_level]:
            return {
                "level":  c4_level,
                "reason": (f"Non-strain-sensitive equipment connection, "
                           f"D={size}\", T={temp}°C → {c4_level} (chart_4, Step 4)"),
                "data_quality": "",
            }
        return {
            "level":  mat_level,
            "reason": (f"{material_group}, D={size}\", T={temp}°C "
                       f"→ {mat_level} ({chart_name}, Step 4)"),
            "data_quality": "",
        }

    return {
        "level":  mat_level,
        "reason": (f"{material_group}, D={size}\", T={temp}°C "
                   f"→ {mat_level} ({chart_name}, Step 4)"),
        "data_quality": "",
    }


def classify_dataframe(df: pd.DataFrame, material_map: dict, rules: dict) -> pd.DataFrame:
    """Apply classify_row to every row; returns a copy with 3 new columns."""
    if df.empty:
        out = df.copy()
        for col in ("Level", "Classification_Reason", "Data_Quality_Flag"):
            out[col] = ""
        return out

    results = df.apply(
        lambda row: classify_row(row, material_map, rules),
        axis=1,
        result_type="expand",
    )
    out = df.copy()
    out["Level"]                 = results["level"]
    out["Classification_Reason"] = results["reason"]
    out["Data_Quality_Flag"]     = results["data_quality"]
    return out

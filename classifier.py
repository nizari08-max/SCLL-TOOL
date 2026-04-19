"""
classifier.py — 5-step deterministic classification engine.

Steps (strict priority — early return on match):
  1. FRP/GRE material → always Level 1
  2. Exception flags → Level 1 (17 checks)
  3. Strain-sensitive equipment override (chart 3 logic)
  4. Material chart lookup (chart_1 / chart_2 / chart_4)
  5. Missing data → NEEDS REVIEW

No file I/O here. All numeric thresholds come from rules dict (loaded from rules.yaml).

Equipment detection supports two modes (set in rules['strain_sensitive_equipment']['detection_mode']):
  "prefix"  (default) — tag starts with a configured prefix (e.g. P-, E-A)
  "keyword" — tag contains a configured keyword (e.g. "PUMP", "HEATER")
              Used when FROM/TO holds full descriptive names.
"""

import warnings
import pandas as pd

from parser import coerce_numeric


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def resolve_material_group(pipe_class: str, material_map: dict[str, str]) -> str:
    if not pipe_class or pd.isna(pipe_class):
        return "UNKNOWN"
    return material_map.get(str(pipe_class).strip().upper(), "UNKNOWN")


def _parse_tag_fragments(tag_value) -> list[str]:
    """Split a FROM/TO cell into individual tag fragments (handles '/', '&', ';')."""
    if not tag_value or pd.isna(tag_value):
        return []
    raw = str(tag_value).strip()
    for sep in ["/", "&", ";"]:
        raw = raw.replace(sep, "|")
    return [t.strip() for t in raw.split("|") if t.strip()]


# ── Prefix-based detection ────────────────────────────────────────────────────

def _match_tag_prefix(tag: str, tag_patterns: list[dict]) -> tuple[bool, str]:
    """Check if a tag starts with any strain-sensitive prefix. Returns (matched, equip_type)."""
    tag_upper = tag.strip().upper()
    for pattern in tag_patterns:
        prefix = str(pattern["prefix"]).upper()
        if tag_upper.startswith(prefix):
            return True, pattern["type"]
    return False, ""


def _is_strain_sensitive_prefix(tag_value, rules: dict) -> tuple[bool, str]:
    """Prefix mode: check one FROM/TO cell. Returns (is_sensitive, equip_type)."""
    tag_patterns = rules["strain_sensitive_equipment"]["tag_patterns"]
    priority_order = ["centrifugal_pump", "reciprocating_pump", "compressor",
                      "turbine", "heater", "air_cooler"]

    found_types: list[str] = []
    for tag in _parse_tag_fragments(tag_value):
        matched, equip_type = _match_tag_prefix(tag, tag_patterns)
        if matched:
            found_types.append(equip_type)

    if not found_types:
        return False, ""
    for ptype in priority_order:
        if ptype in found_types:
            return True, ptype
    return True, found_types[0]


# ── Keyword-based detection ───────────────────────────────────────────────────

def _is_strain_sensitive_keyword(tag_value, rules: dict) -> tuple[bool, str]:
    """
    Keyword mode: check if the FROM/TO descriptive name contains a strain-sensitive keyword.
    Patterns evaluated in order; first match wins (allows "RECIPROCATING PUMP" before "PUMP").
    Returns (is_sensitive, equip_type).
    """
    if not tag_value or pd.isna(tag_value):
        return False, ""
    tag_upper = str(tag_value).strip().upper()

    keyword_patterns = rules["strain_sensitive_equipment"].get("keyword_patterns", [])
    for pattern in keyword_patterns:
        for kw in pattern["keywords"]:
            if kw.upper() in tag_upper:
                return True, pattern["type"]
    return False, ""


def _is_non_strain_sensitive_equipment_keyword(tag_value, rules: dict) -> bool:
    """
    Keyword mode: check if FROM/TO refers to non-strain-sensitive equipment (vessel, tank, etc.).
    Returns True only when the tag explicitly matches a known equipment keyword.
    Pipe line numbers and generic locations (DRAIN, SAFE LOCATION, etc.) return False.
    """
    if not tag_value or pd.isna(tag_value):
        return False
    tag_upper = str(tag_value).strip().upper()
    if not tag_upper or tag_upper == "NAN":
        return False

    ns_keywords = rules["strain_sensitive_equipment"].get("non_strain_sensitive_keywords", [])
    return any(kw.upper() in tag_upper for kw in ns_keywords)


# ── Public interface ──────────────────────────────────────────────────────────

def is_strain_sensitive_equipment(tag_value, rules: dict) -> tuple[bool, str]:
    """
    Dispatch to prefix or keyword detection based on rules config.
    Returns (is_sensitive, equip_type).
    """
    mode = rules["strain_sensitive_equipment"].get("detection_mode", "prefix")
    if mode == "keyword":
        return _is_strain_sensitive_keyword(tag_value, rules)
    return _is_strain_sensitive_prefix(tag_value, rules)


def _has_equipment_connection(tag_value, rules: dict) -> bool:
    """
    Returns True if FROM/TO represents an actual equipment connection
    (strain-sensitive OR non-strain-sensitive).

    Prefix mode: any non-empty tag = equipment (original behaviour).
    Keyword mode: only when tag matches a recognized equipment keyword.
                  Pipe line numbers and unrecognised strings return False.
    """
    if not tag_value or pd.isna(tag_value):
        return False
    s = str(tag_value).strip()
    if not s or s.lower() == "nan":
        return False

    mode = rules["strain_sensitive_equipment"].get("detection_mode", "prefix")
    if mode == "keyword":
        sensitive, _ = _is_strain_sensitive_keyword(tag_value, rules)
        if sensitive:
            return True
        return _is_non_strain_sensitive_equipment_keyword(tag_value, rules)

    # prefix mode: any non-empty tag = equipment
    return True


# ─────────────────────────────────────────────────────────────────────────────
# Exception flags
# ─────────────────────────────────────────────────────────────────────────────

def _eval_extra_condition(row_val, condition: dict | None) -> bool:
    if condition is None:
        return True
    raw = row_val
    if raw is None or pd.isna(raw) if not isinstance(raw, str) else raw.strip() == "":
        return False
    try:
        numeric_val = float(str(raw).strip().rstrip("°C").rstrip("°").rstrip("C").strip())
    except (ValueError, TypeError):
        return False
    threshold = float(condition["value"])
    op = condition["operator"]
    if op == ">":
        return numeric_val > threshold
    if op == ">=":
        return numeric_val >= threshold
    if op == "<":
        return numeric_val < threshold
    if op == "<=":
        return numeric_val <= threshold
    return False


def check_exception_flags(row: pd.Series, rules: dict) -> tuple[bool, str]:
    """
    Apply all STEP 2 exception flag checks in order.
    Returns (triggered, reason). First triggered exception wins.
    """
    true_vals = {str(v).strip().lower() for v in rules.get("true_values", [])}

    for exc in rules["exception_flags"]:
        flag_col = exc["flag"]
        cell = row.get(flag_col, None)

        if cell is None or pd.isna(cell) if not isinstance(cell, str) else cell.strip() == "":
            continue
        if str(cell).strip().lower() not in true_vals:
            continue

        extra = exc.get("extra_condition")
        if extra is not None:
            extra_col = extra["field"]
            extra_val = row.get(extra_col, None)
            if not _eval_extra_condition(extra_val, extra):
                continue

        return True, exc["reason"]

    return False, ""


# ─────────────────────────────────────────────────────────────────────────────
# Chart lookups
# ─────────────────────────────────────────────────────────────────────────────

def lookup_chart(chart_name: str, size: float, temp: float, rules: dict) -> str:
    """
    Look up classification level from a chart (chart_1, chart_2, or chart_4).
    Finds the first size band where size <= max_size, then applies thresholds.
    Returns 'Level 1', 'Level 2', or 'Level 3'.
    """
    if chart_name not in rules:
        raise KeyError(
            f"Chart '{chart_name}' not found in rules.yaml. "
            f"Available charts: {[k for k in rules if k.startswith('chart_')]}"
        )

    bands = rules[chart_name]["size_bands"]
    for band in bands:
        if size <= float(band["max_size"]):
            l1 = float(band["l1_threshold"])
            l2 = float(band["l2_threshold"])
            if temp > l1:
                return "Level 1"
            if temp > l2:
                return "Level 2"
            return "Level 3"

    return "Level 3"


def _apply_chart_3(size: float, temp: float, equip_type: str, rules: dict) -> tuple[str, str]:
    """Apply strain-sensitive equipment chart 3 rules. Returns (level, reason)."""
    for rule in rules["strain_sensitive_equipment"]["chart_3"]:
        ctype = rule["condition_type"]

        if ctype == "equip_type":
            if equip_type in rule["equip_types"]:
                return rule["result"], rule["reason"]

        elif ctype == "temp_threshold":
            op = rule["temp_operator"]
            val = float(rule["temp_value"])
            if op == ">=" and temp >= val:
                return rule["result"], rule["reason"]
            if op == ">" and temp > val:
                return rule["result"], rule["reason"]

        elif ctype == "size_temp":
            s_op = rule["size_operator"]
            s_val = float(rule["size_value"])
            t_op = rule["temp_operator"]
            t_val = float(rule["temp_value"])
            s_ok = (s_op == "<=" and size <= s_val) or (s_op == ">" and size > s_val)
            t_ok = (t_op == ">=" and temp >= t_val) or (t_op == "<" and temp < t_val)
            if s_ok and t_ok:
                return rule["result"], rule["reason"]

        elif ctype == "fallback":
            return rule["result"], rule["reason"]

    return "Level 3", "Strain-sensitive equipment, no specific rule matched → Level 3"


# ─────────────────────────────────────────────────────────────────────────────
# Core row classifier
# ─────────────────────────────────────────────────────────────────────────────

def classify_row(row: pd.Series, material_map: dict, rules: dict) -> dict:
    """
    Apply 5-step classification to a single row.
    Returns dict with keys: level, reason, review_flag.
    """

    # ── STEP 5 pre-check: required numeric fields ────────────────────────────
    raw_size = row.get("size", None)
    raw_temp = row.get("design_temperature", None)
    raw_material = row.get("material", None)

    def _to_float(val):
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return None
        s = str(val).strip().rstrip("°C").rstrip("°").rstrip("C").strip()
        # Compound temps like "60/-20" or "186/-20": use the maximum (design max temp)
        if "/" in s:
            parts = [p.strip().rstrip("°C").rstrip("°").rstrip("C").strip() for p in s.split("/")]
            nums = []
            for p in parts:
                try:
                    nums.append(float(p))
                except (ValueError, TypeError):
                    pass
            return max(nums) if nums else None
        try:
            return float(s)
        except (ValueError, TypeError):
            return None

    size = _to_float(raw_size)
    temp = _to_float(raw_temp)
    material_str = str(raw_material).strip() if raw_material and not (
        isinstance(raw_material, float) and pd.isna(raw_material)) else ""

    missing_fields = []
    if size is None or str(raw_size).strip() == "":
        missing_fields.append("Size")
    if temp is None or str(raw_temp).strip() == "":
        missing_fields.append("Design Temperature")
    if not material_str:
        missing_fields.append("Material")

    if missing_fields:
        return {
            "level": "",
            "reason": f"Missing required field(s): {', '.join(missing_fields)} — cannot classify",
            "review_flag": "NEEDS REVIEW",
        }

    # ── STEP 1: FRP/GRE check ────────────────────────────────────────────────
    material_group = resolve_material_group(material_str, material_map)

    if material_group == "FRP":
        return {
            "level": "Level 1",
            "reason": "FRP/GRE material → always Level 1 (Step 1)",
            "review_flag": "",
        }

    if material_group == "UNKNOWN":
        return {
            "level": "",
            "reason": f"Material code '{material_str}' not found in material_mapping.yaml — cannot classify",
            "review_flag": "NEEDS REVIEW",
        }

    # ── STEP 2: Exception flags → Level 1 ────────────────────────────────────
    triggered, reason = check_exception_flags(row, rules)
    if triggered:
        return {"level": "Level 1", "reason": reason + " (Step 2)", "review_flag": ""}

    # ── STEP 3: Strain-sensitive equipment override ───────────────────────────
    from_tag = row.get("from_equipment", None)
    to_tag = row.get("to_equipment", None)

    from_sensitive, from_type = is_strain_sensitive_equipment(from_tag, rules)
    to_sensitive, to_type = is_strain_sensitive_equipment(to_tag, rules)

    if from_sensitive or to_sensitive:
        equip_type = from_type if from_sensitive else to_type
        if from_sensitive and to_sensitive:
            priority = ["centrifugal_pump", "reciprocating_pump", "compressor",
                        "turbine", "heater", "air_cooler"]
            equip_type = from_type
            for ptype in priority:
                if ptype in (from_type, to_type):
                    equip_type = ptype
                    break

        level, step3_reason = _apply_chart_3(size, temp, equip_type, rules)
        return {"level": level, "reason": step3_reason + " (Step 3)", "review_flag": ""}

    # ── STEP 4: Material chart lookup ─────────────────────────────────────────
    _level_rank = {"Level 1": 1, "Level 2": 2, "Level 3": 3}

    mat_chart_name = rules["material_chart_map"].get(material_group)
    if not mat_chart_name:
        return {
            "level": "",
            "reason": f"No chart defined for material group '{material_group}' — cannot classify",
            "review_flag": "NEEDS REVIEW",
        }

    mat_level = lookup_chart(mat_chart_name, size, temp, rules)

    has_equipment = (
        _has_equipment_connection(from_tag, rules) or
        _has_equipment_connection(to_tag, rules)
    )

    if has_equipment:
        chart4_level = lookup_chart("chart_4", size, temp, rules)
        if _level_rank[chart4_level] <= _level_rank[mat_level]:
            level = chart4_level
            reason = (
                f"Non-strain-sensitive equipment connection, size={size}\", temp={temp}°C "
                f"→ {level} (chart_4 is stricter, Step 4)"
            )
        else:
            level = mat_level
            reason = (
                f"Non-strain-sensitive equipment connection, size={size}\", temp={temp}°C "
                f"→ {level} ({mat_chart_name} is stricter than chart_4, Step 4)"
            )
    else:
        level = mat_level
        reason = (
            f"Material group {material_group}, size={size}\", temp={temp}°C "
            f"→ {level} ({mat_chart_name}, Step 4)"
        )

    return {"level": level, "reason": reason, "review_flag": ""}


def classify_dataframe(df: pd.DataFrame, material_map: dict, rules: dict) -> pd.DataFrame:
    """
    Apply classify_row to every row in df.
    Appends columns: Level, Classification_Reason, Review_Flag.
    Returns augmented DataFrame.
    """
    results = df.apply(
        lambda row: classify_row(row, material_map, rules),
        axis=1,
        result_type="expand",
    )
    df = df.copy()
    df["Level"] = results["level"]
    df["Classification_Reason"] = results["reason"]
    df["Review_Flag"] = results["review_flag"]
    return df

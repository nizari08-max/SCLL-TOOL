"""
parser.py — Format-agnostic Excel line list reader.

Reads the input workbook using the layout information in `detected_config`
(from format_detector.detect_format). Renames columns to the internal field
names the classifier expects, converts mm → NPS inches when required, and
splits rows into (in_scope, excluded) using the detected scope strategy.

This module performs NO classification and has NO hardcoded column names,
scope values, or project knowledge. Every project-variable setting comes
from detected_config.
"""

from __future__ import annotations

import os
import sys
from typing import Optional

import pandas as pd
import yaml


# ─────────────────────────────────────────────────────────────────────────────
# Config loaders
# ─────────────────────────────────────────────────────────────────────────────

def load_rules(rules_path: str) -> dict:
    """Load rules.yaml and return as-is. Structure is validated by classifier.py, not here."""
    with open(rules_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def load_material_map(map_path: str) -> dict[str, str]:
    """Return flat dict: pipe_class_code.upper() → material_group."""
    with open(map_path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)
    flat: dict[str, str] = {}
    for group, payload in data.get("groups", {}).items():
        for code in payload.get("codes", []):
            flat[str(code).strip().upper()] = group
    return flat


# ─────────────────────────────────────────────────────────────────────────────
# Excel reader
# ─────────────────────────────────────────────────────────────────────────────

def read_linelist(filepath: str, detected_config: dict, rules: dict) -> pd.DataFrame:
    """
    Read the line list Excel using detected_config for layout + column mapping.

    detected_config keys used:
      sheet_name, header_row, skip_rows, column_mappings, size_unit

    rules keys used:
      size_config.mm_to_nps (for mm → NPS conversion when size_unit == 'mm')
    """
    sheet_name = detected_config.get("sheet_name", 0)
    header_row = int(detected_config.get("header_row", 0))
    skip_rows  = detected_config.get("skip_rows", [])

    try:
        df = pd.read_excel(
            filepath,
            sheet_name=sheet_name,
            header=header_row,
            skiprows=skip_rows if skip_rows else None,
            dtype=str,
        )
    except FileNotFoundError:
        print(f"ERROR: Input file not found: {filepath}", file=sys.stderr)
        sys.exit(2)
    except Exception as exc:
        print(f"ERROR reading '{filepath}': {exc}", file=sys.stderr)
        sys.exit(2)

    df.columns = [str(c).strip() for c in df.columns]
    df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
    df = df.dropna(how="all").reset_index(drop=True)

    # Rename detected columns → internal field names
    col_mappings: dict[str, str] = detected_config.get("column_mappings", {})
    reverse = {v.strip().lower(): k for k, v in col_mappings.items() if v}
    rename: dict[str, str] = {}
    for col in df.columns:
        key = str(col).strip().lower()
        if key in reverse:
            rename[col] = reverse[key]
    df.rename(columns=rename, inplace=True)

    # mm → NPS conversion when needed
    if detected_config.get("size_unit") == "mm" and "size" in df.columns:
        mm_to_nps = rules.get("size_config", {}).get("mm_to_nps", {})
        df["size"] = _convert_mm_to_nps(df["size"], mm_to_nps)

    return df


def _convert_mm_to_nps(size_series: pd.Series, mm_to_nps: dict) -> pd.Series:
    """Convert mm → NPS using the lookup table. Unmappable values → None (NEEDS REVIEW)."""
    lookup = {int(k): float(v) for k, v in mm_to_nps.items()}

    def _cvt(val):
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return None
        s = str(val).strip()
        if s == "" or s.upper() in ("XX", "TBD", "N/A", "-"):
            return None
        try:
            return lookup.get(int(float(s)))
        except (ValueError, TypeError):
            return None

    return size_series.map(_cvt)


# ─────────────────────────────────────────────────────────────────────────────
# Scope filter (detected_config-driven)
# ─────────────────────────────────────────────────────────────────────────────

def filter_scope(
    df: pd.DataFrame, detected_config: dict
) -> tuple[pd.DataFrame, pd.DataFrame, pd.Series]:
    """
    Split df into (in_scope, excluded, exclusion_labels).

    exclusion_labels: per-row Series with values "VENDOR"/"CLIENT"/"" (one entry
    per row in the ORIGINAL df, aligned by index). Used by output.py to populate
    the DATA QUALITY FLAG column with "SCOPE: VENDOR" etc.
    """
    mode = detected_config.get("scope_mode", "assume_all_in_scope")
    col  = _internal_col_for(detected_config, detected_config.get("scope_column", ""))

    labels = pd.Series([""] * len(df), index=df.index, dtype=object)

    if mode == "assume_all_in_scope" or col is None or col not in df.columns:
        return df.copy(), df.iloc[0:0].copy(), labels

    if mode == "include_values":
        include = {str(v).strip().lower() for v in detected_config.get("scope_include_values", [])}
        def _label(val):
            if val is None or (isinstance(val, float) and pd.isna(val)):
                return "UNKNOWN"
            s = str(val).strip()
            if s.lower() in include or s == "":
                return ""
            return s.upper()
        labels = df[col].map(_label)

    elif mode == "text_keywords":
        keywords = [str(k).lower() for k in detected_config.get("scope_exclude_keywords", [])]
        def _label(val):
            if val is None or (isinstance(val, float) and pd.isna(val)):
                return ""
            s = str(val).strip().lower()
            if not s:
                return ""
            for kw in keywords:
                if kw in s:
                    return kw.upper()
            return ""
        labels = df[col].map(_label)

    elif mode == "column_exclude_values":
        excludes = {str(v).strip().lower() for v in detected_config.get("scope_exclude_values", [])}
        def _label(val):
            if val is None or (isinstance(val, float) and pd.isna(val)):
                return ""
            return str(val).strip().upper() if str(val).strip().lower() in excludes else ""
        labels = df[col].map(_label)

    mask_excluded = labels.astype(str).str.len() > 0
    in_scope = df[~mask_excluded].copy()
    excluded = df[mask_excluded].copy()
    return in_scope, excluded, labels


def _internal_col_for(detected_config: dict, excel_header: str) -> Optional[str]:
    """Given an Excel header, return its internal field name (if mapped)."""
    if not excel_header:
        return None
    col_mappings = detected_config.get("column_mappings", {})
    for internal, header in col_mappings.items():
        if header and header.strip().lower() == excel_header.strip().lower():
            return internal
    return None


# ─────────────────────────────────────────────────────────────────────────────
# Numeric coercion (used by classifier)
# ─────────────────────────────────────────────────────────────────────────────

def coerce_numeric(df: pd.DataFrame, col: str) -> pd.Series:
    """Return numeric Series for col; strips °C-like suffixes; NaN on failure."""
    if col not in df.columns:
        return pd.Series([None] * len(df), index=df.index)

    def _parse(val):
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return None
        s = str(val).strip().rstrip("°C").rstrip("°").rstrip("C").strip()
        if not s:
            return None
        try:
            return float(s)
        except ValueError:
            return None

    return df[col].map(_parse)

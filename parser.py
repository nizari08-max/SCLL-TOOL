"""
parser.py — Line list Excel reader, column validator, and scope filter.

Responsibilities:
- Load rules.yaml and material_mapping.yaml
- Read input .xlsx into a DataFrame (supports multi-row headers via excel_config)
- Rename columns using rules['column_mappings']
- Convert size from mm to NPS inches when size_config.unit == "mm"
- Validate required columns are present
- Split rows into in-scope and excluded

No classification logic here.
"""

import re
import sys
import warnings
import yaml
import pandas as pd


def load_rules(rules_path: str) -> dict:
    """Load and minimally validate rules.yaml. Raises ValueError on missing top-level keys."""
    with open(rules_path, "r", encoding="utf-8") as f:
        rules = yaml.safe_load(f)

    required_keys = [
        "column_mappings", "scope", "required_columns", "exception_flags",
        "strain_sensitive_equipment", "chart_1", "chart_2", "chart_4",
        "material_chart_map",
    ]
    missing = [k for k in required_keys if k not in rules]
    if missing:
        raise ValueError(
            f"rules.yaml is missing required top-level keys: {missing}\n"
            f"Check the file at: {rules_path}"
        )
    return rules


def load_material_map(map_path: str) -> dict[str, str]:
    """
    Load material_mapping.yaml and return a flat dict: pipe_class_code → group_name.
    Example: {"BA1": "CS", "BD1": "SS", "BK1": "FRP"}
    """
    with open(map_path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)

    flat: dict[str, str] = {}
    for group_name, group_data in data.get("groups", {}).items():
        for code in group_data.get("codes", []):
            flat[str(code).strip().upper()] = group_name
    return flat


def _build_reverse_column_map(rules: dict) -> dict[str, str]:
    """
    Build a mapping from Excel header (case-insensitive, stripped) → internal field name.
    Empty-string mappings are skipped (columns intentionally omitted from a rules file).
    """
    reverse: dict[str, str] = {}
    for internal_name, excel_header in rules["column_mappings"].items():
        if excel_header:
            reverse[str(excel_header).strip().lower()] = internal_name
    return reverse


def read_linelist(filepath: str, rules: dict) -> pd.DataFrame:
    """
    Read the line list Excel file into a DataFrame.

    Supports multi-row headers via optional rules['excel_config']:
      sheet_name  — sheet to read (default: 0 = first sheet)
      header_row  — 0-indexed row to use as column names (default: 0)
      skip_rows   — list of 0-indexed row numbers to skip after the header row
                    (used to drop units rows and blank rows between header and data)

    Renames columns from Excel headers to internal field names using column_mappings.
    Returns the renamed DataFrame (all rows, unfiltered).
    """
    excel_cfg = rules.get("excel_config", {})
    sheet_name = excel_cfg.get("sheet_name", 0)
    header_row = int(excel_cfg.get("header_row", 0))
    skip_rows = excel_cfg.get("skip_rows", [])

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
        print(f"ERROR: Could not read input file '{filepath}': {exc}", file=sys.stderr)
        sys.exit(2)

    # Strip leading/trailing whitespace from all column headers
    df.columns = [str(c).strip() for c in df.columns]

    # Strip whitespace from all cell values
    df = df.map(lambda x: x.strip() if isinstance(x, str) else x)

    # Drop rows that are entirely empty (can appear between header and data in multi-row layouts)
    df = df.dropna(how="all").reset_index(drop=True)

    reverse_map = _build_reverse_column_map(rules)

    # Rename using case-insensitive header matching
    rename_dict: dict[str, str] = {}
    for col in df.columns:
        key = col.lower()
        if key in reverse_map:
            rename_dict[col] = reverse_map[key]
    df.rename(columns=rename_dict, inplace=True)

    # Apply mm → NPS conversion if configured
    size_cfg = rules.get("size_config", {})
    if size_cfg.get("unit", "inches") == "mm" and "size" in df.columns:
        df["size"] = _convert_size_mm_to_nps(df["size"], size_cfg["mm_to_nps"])

    # Warn about columns that did not get mapped
    known = set(rules["column_mappings"].keys())
    unmapped = [c for c in df.columns if c not in known]
    if unmapped:
        warnings.warn(
            f"The following columns were not recognised and will be carried through unchanged: {unmapped}",
            stacklevel=2,
        )

    return df


def _convert_size_mm_to_nps(size_series: pd.Series, mm_to_nps: dict) -> pd.Series:
    """
    Convert a Series of pipe sizes from mm to NPS inches using the lookup table.
    Values not in the table (e.g. "XX", "TBD", unmapped mm values) are set to None
    so the classifier will flag them as NEEDS REVIEW.
    """
    # Build a lookup with int keys for flexible matching
    lookup = {int(k): float(v) for k, v in mm_to_nps.items()}

    def _convert(val):
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return None
        s = str(val).strip()
        if s == "" or s.upper() in ("XX", "TBD", "N/A", "-"):
            return None
        try:
            mm = int(float(s))
        except (ValueError, TypeError):
            return None
        return lookup.get(mm)  # None if mm value not in table

    return size_series.map(_convert)


def validate_columns(df: pd.DataFrame, rules: dict) -> list[str]:
    """
    Check that all required_columns (internal names) are present in the DataFrame.
    Returns a list of missing internal column names (empty list = all present).
    """
    return [col for col in rules["required_columns"] if col not in df.columns]


def _parse_bool_value(value, rules: dict) -> bool:
    """Return True if value matches any of the configured true_values."""
    true_vals = {str(v).strip().lower() for v in rules.get("true_values", [])}
    return str(value).strip().lower() in true_vals


def coerce_numeric(df: pd.DataFrame, col: str) -> pd.Series:
    """
    Return a numeric Series for column col, coercing errors to NaN.
    Handles values like "150°C" or "150 C" by stripping non-numeric suffixes.
    """
    if col not in df.columns:
        return pd.Series([None] * len(df), index=df.index)

    def _parse(val):
        if pd.isna(val) or str(val).strip() == "":
            return None
        clean = str(val).strip().rstrip("°C").rstrip("°").rstrip("C").strip()
        try:
            return float(clean)
        except ValueError:
            return None

    return df[col].map(_parse)


def filter_scope(df: pd.DataFrame, rules: dict) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Split the DataFrame into (in_scope_df, excluded_df).

    Three modes (set via rules['scope']['mode']):

    "include_values" (default):
        Rows whose Scope value is in include_values are kept.
        All other rows are marked EXCLUDED.

    "text_keywords":
        Blank/null scope cell → in scope.
        Cell containing any exclude_keyword (case-insensitive) → EXCLUDED.
        Used when scope is embedded in a free-text NOTES/REMARKS column.

    "column_exclude_values":
        Checks an arbitrary already-mapped column (by internal name) for exact
        value matches. Matching rows → EXCLUDED.  No 'scope' column needed.
        Used when exclusion indicator is in a non-dedicated column (e.g. MOC).
    """
    scope_cfg = rules.get("scope", {})
    mode = scope_cfg.get("mode", "include_values")

    # ── column_exclude_values: check any named column, no 'scope' required ────
    if mode == "column_exclude_values":
        col_name = scope_cfg.get("column", "scope")
        exclude_vals_lower = {str(v).strip().lower() for v in scope_cfg.get("exclude_values", [])}
        if col_name in df.columns:
            mask_excluded = df[col_name].map(
                lambda x: str(x).strip().lower() in exclude_vals_lower
                if (x is not None and not (isinstance(x, float) and pd.isna(x)))
                else False
            )
        else:
            warnings.warn(
                f"column_exclude_values: column '{col_name}' not found. Treating all rows as in-scope.",
                stacklevel=2,
            )
            mask_excluded = pd.Series([False] * len(df), index=df.index)
        in_scope = df[~mask_excluded].copy()
        excluded = df[mask_excluded].copy()
        return in_scope, excluded

    if "scope" not in df.columns:
        warnings.warn(
            "Column 'scope' not found in input. Treating all rows as in-scope.",
            stacklevel=2,
        )
        return df.copy(), pd.DataFrame(columns=df.columns)

    if mode == "text_keywords":
        exclude_keywords = [str(k).lower() for k in scope_cfg.get("exclude_keywords", [])]

        def _is_excluded(cell) -> bool:
            if cell is None or (isinstance(cell, float) and pd.isna(cell)):
                return False
            cell_str = str(cell).strip()
            if cell_str == "" or cell_str.lower() == "nan":
                return False
            cell_lower = cell_str.lower()
            return any(kw in cell_lower for kw in exclude_keywords)

        mask_excluded = df["scope"].map(_is_excluded)

    else:
        include_vals = {str(v).strip().lower() for v in scope_cfg.get("include_values", [])}
        mask_excluded = df["scope"].map(
            lambda x: str(x).strip().lower() not in include_vals if pd.notna(x) else True
        )

    in_scope = df[~mask_excluded].copy()
    excluded = df[mask_excluded].copy()
    return in_scope, excluded

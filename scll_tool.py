"""
scll_tool.py — CLI entry point for the Stress Critical Line List Tool.

Usage:
    python scll_tool.py --input linelist.xlsx --output scll_output.xlsx
    python scll_tool.py --input linelist.xlsx --output out.xlsx --rules rules_q37027.yaml
    python scll_tool.py --input linelist.xlsx --output out.xlsx --mode classify
    python scll_tool.py --input linelist.xlsx --output out.xlsx --mode cn
    python scll_tool.py --input linelist.xlsx --output out.xlsx --mode full

Modes:
    classify — classify only, no CN assignment
    cn       — CN assignment only (assumes input already has Level column)
    full     — classify then assign CNs (default)

Output format:
    If the input has a "Coversheet" + "Line List" sheet structure (JESA SCLL format),
    or if the rules file contains output_config.format == "jesa", the JESA format
    writer (output_jesa.py) is used automatically.
    Otherwise the generic writer (output.py) is used.

Exit codes:
    0 = success
    1 = validation or configuration error
    2 = file I/O error
"""

import argparse
import sys
import os

import openpyxl

from parser import (
    load_rules,
    load_material_map,
    read_linelist,
    validate_columns,
    filter_scope,
)
from classifier import classify_dataframe
from cn_assigner import assign_cns
from output import write_output
from output_jesa import write_jesa_output


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog="scll_tool",
        description="JESA Stress Critical Line List — automated classification tool",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python scll_tool.py --input linelist.xlsx --output scll_output.xlsx
  python scll_tool.py --input Q37027-file.xlsx --output out.xlsx --rules rules_q37027.yaml
  python scll_tool.py --input linelist.xlsx --output out.xlsx --mode classify
        """,
    )
    p.add_argument("--input",        required=True,  metavar="FILE",
                   help="Path to the input line list .xlsx file")
    p.add_argument("--output",       required=True,  metavar="FILE",
                   help="Path to write the classified output .xlsx file")
    p.add_argument("--rules",        default="rules.yaml", metavar="FILE",
                   help="Path to rules.yaml (default: rules.yaml in current directory)")
    p.add_argument("--material-map", default="material_mapping.yaml", metavar="FILE",
                   help="Path to material_mapping.yaml (default: material_mapping.yaml)")
    p.add_argument("--mode",         default="full",
                   choices=["classify", "cn", "full"],
                   help="classify=classify only | cn=CN only | full=both (default: full)")
    return p.parse_args()


def _check_file_exists(path: str, label: str) -> None:
    if not os.path.isfile(path):
        print(f"ERROR: {label} not found: {path}", file=sys.stderr)
        sys.exit(2)


def _is_jesa_format(input_path: str, rules: dict) -> bool:
    """
    Return True if the input should be processed with the JESA format writer.
    True when rules file declares output_config.format == "jesa", OR when the
    input workbook contains both a "Coversheet" and a "Line List" sheet.
    """
    if rules.get("output_config", {}).get("format") == "jesa":
        return True
    try:
        wb = openpyxl.load_workbook(input_path, read_only=True, data_only=True)
        sheet_names = [s.lower() for s in wb.sheetnames]
        wb.close()
        return "coversheet" in sheet_names and "line list" in sheet_names
    except Exception:
        return False


def main() -> None:
    args = parse_args()

    # ── Prevent overwriting the input file ───────────────────────────────────
    if os.path.abspath(args.input) == os.path.abspath(args.output):
        print("ERROR: --output path must differ from --input path. "
              "The tool will never overwrite the input file.", file=sys.stderr)
        sys.exit(1)

    # ── File existence checks ─────────────────────────────────────────────────
    _check_file_exists(args.input,        "Input file")
    _check_file_exists(args.rules,        "Rules file")
    _check_file_exists(args.material_map, "Material mapping file")

    # ── Load configuration ────────────────────────────────────────────────────
    print(f"Loading rules from:          {args.rules}")
    print(f"Loading material mapping from: {args.material_map}")
    try:
        rules        = load_rules(args.rules)
        material_map = load_material_map(args.material_map)
    except Exception as exc:
        print(f"ERROR loading configuration: {exc}", file=sys.stderr)
        sys.exit(1)

    jesa_format = _is_jesa_format(args.input, rules)
    if jesa_format:
        print("Detected JESA SCLL format — using JESA output writer.")

    # ── Read input ────────────────────────────────────────────────────────────
    print(f"Reading input file:          {args.input}")
    df = read_linelist(args.input, rules)

    if df.empty:
        print("WARNING: Input file has no data rows. Output will be empty.")

    # ── Validate required columns ─────────────────────────────────────────────
    missing_cols = validate_columns(df, rules)
    if missing_cols:
        print(
            f"ERROR: Required column(s) missing from input file: {missing_cols}\n"
            f"Check that rules.yaml column_mappings match your Excel headers.",
            file=sys.stderr,
        )
        sys.exit(1)

    # ── Filter scope ──────────────────────────────────────────────────────────
    import warnings
    with warnings.catch_warnings(record=True) as w_list:
        warnings.simplefilter("always")
        in_scope_df, excluded_df = filter_scope(df, rules)
    for w in w_list:
        print(f"WARNING: {w.message}", file=sys.stderr)

    print(f"  Total rows:      {len(df)}")
    print(f"  In scope:        {len(in_scope_df)}")
    print(f"  Excluded:        {len(excluded_df)}")

    # ── Classify ──────────────────────────────────────────────────────────────
    classified_df = in_scope_df.copy()
    if args.mode in ("classify", "full"):
        print("Classifying in-scope lines...")
        if not in_scope_df.empty:
            classified_df = classify_dataframe(in_scope_df, material_map, rules)
        else:
            classified_df["Level"] = ""
            classified_df["Classification_Reason"] = ""
            classified_df["Review_Flag"] = ""
    else:
        # cn-only mode: Level column must already exist or is set empty
        if "Level" not in classified_df.columns:
            classified_df["Level"] = ""
            classified_df["Classification_Reason"] = ""
            classified_df["Review_Flag"] = ""

    # ── Print classification summary ──────────────────────────────────────────
    if "Level" in classified_df.columns and not classified_df.empty:
        l1 = int((classified_df["Level"] == "Level 1").sum())
        l2 = int((classified_df["Level"] == "Level 2").sum())
        l3 = int((classified_df["Level"] == "Level 3").sum())
        rv = int((classified_df.get("Review_Flag", "") == "NEEDS REVIEW").sum())
        print(f"  Level I:         {l1}")
        print(f"  Level II:        {l2}")
        print(f"  Level III:       {l3}")
        print(f"  Needs Review:    {rv}")

    # ── CN assignment ─────────────────────────────────────────────────────────
    cn_proposals = []
    if args.mode in ("cn", "full") and rules.get("cn_settings"):
        print("Assigning Calculation Numbers to Level 1 lines...")
        classified_df, cn_proposals = assign_cns(classified_df, rules)
        auto     = sum(1 for p in cn_proposals if p["review_flag"] == "[AUTO-CONFIRMED]")
        flagged  = len(cn_proposals) - auto
        missing  = sum(1 for p in cn_proposals if "[REVIEW-MISSING-DATA]" in p["review_flag"])
        print(f"  Proposed CNs:    {len(cn_proposals)}")
        print(f"  Auto-confirmed:  {auto}")
        print(f"  Flagged:         {flagged}  (including {missing} missing-data)")
    elif not rules.get("cn_settings"):
        print("CN settings not configured — skipping CN assignment.")

    # ── Write output ──────────────────────────────────────────────────────────
    print(f"Writing output to:           {args.output}")
    try:
        if jesa_format:
            write_jesa_output(classified_df, excluded_df, args.output, rules, cn_proposals)
        else:
            write_output(classified_df, excluded_df, args.output, rules, cn_proposals)
    except PermissionError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(2)
    except Exception as exc:
        import traceback
        print(f"ERROR writing output: {exc}", file=sys.stderr)
        traceback.print_exc()
        sys.exit(2)

    print("Done.")


if __name__ == "__main__":
    main()

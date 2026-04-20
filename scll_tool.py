"""
scll_tool.py — CLI entry point for the Stress Critical Line List Tool.

Usage:
    python scll_tool.py --input linelist.xlsx --output scll_output.xlsx

The tool is format-agnostic: format_detector.py inspects the workbook and
builds a detected_config that drives the rest of the pipeline. No project-
specific rules files are needed.

Exit codes:
    0 = success
    1 = validation or configuration error
    2 = file I/O error
"""

from __future__ import annotations

import argparse
import os
import sys
import warnings

# Force UTF-8 on Windows consoles so unicode in detection summary / logs doesn't crash
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

from parser import load_rules, load_material_map, read_linelist, filter_scope
from classifier import classify_dataframe
from cn_assigner import assign_cns
from format_detector import detect_format, apply_detection_to_rules, load_mapping
from output import write_enriched_output


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog="scll_tool",
        description="SCLL Tool — format-agnostic line list classifier",
    )
    p.add_argument("--input",  required=True, metavar="FILE",
                   help="Path to the input line list .xlsx file")
    p.add_argument("--output", required=True, metavar="FILE",
                   help="Path to write the enriched output .xlsx file")
    p.add_argument("--rules", default="rules.yaml", metavar="FILE",
                   help="Path to rules.yaml (default: rules.yaml)")
    p.add_argument("--material-map", default="material_mapping.yaml", metavar="FILE",
                   help="Path to material_mapping.yaml (default: material_mapping.yaml)")
    p.add_argument("--project-code", default="", metavar="CODE",
                   help="Optional project code (e.g. Q37027) — used in CN numbering")
    return p.parse_args()


def _check_file(path: str, label: str) -> None:
    if not os.path.isfile(path):
        print(f"ERROR: {label} not found: {path}", file=sys.stderr)
        sys.exit(2)


def main() -> None:
    args = parse_args()

    if os.path.abspath(args.input) == os.path.abspath(args.output):
        print("ERROR: --output path must differ from --input path.", file=sys.stderr)
        sys.exit(1)

    _check_file(args.input,        "Input file")
    _check_file(args.rules,        "Rules file")
    _check_file(args.material_map, "Material mapping file")

    print(f"Detecting format:       {args.input}")
    try:
        detected_config, summary = detect_format(args.input)
    except Exception as exc:
        print(f"ERROR detecting format: {exc}", file=sys.stderr)
        sys.exit(1)
    print(summary)

    print(f"Loading configuration:  {args.rules} + {args.material_map}")
    try:
        rules        = load_rules(args.rules)
        mapping      = load_mapping()
        material_map = load_material_map(args.material_map)
    except Exception as exc:
        print(f"ERROR loading configuration: {exc}", file=sys.stderr)
        sys.exit(1)

    rules = apply_detection_to_rules(rules, detected_config, mapping)
    if args.project_code and rules.get("cn_settings"):
        rules["cn_settings"]["project_code"] = args.project_code

    print(f"Reading input file:     {args.input}")
    with warnings.catch_warnings(record=True) as w_list:
        warnings.simplefilter("always")
        raw_df = read_linelist(args.input, detected_config, rules)
    for w in w_list:
        print(f"  WARNING: {w.message}", file=sys.stderr)

    in_scope_df, excluded_df, exclusion_labels = filter_scope(raw_df, detected_config)
    print(f"  Total rows:   {len(raw_df)}")
    print(f"  In scope:     {len(in_scope_df)}")
    print(f"  Excluded:     {len(excluded_df)}")

    print("Classifying in-scope lines...")
    if not in_scope_df.empty:
        classified_df = classify_dataframe(in_scope_df, material_map, rules)
    else:
        classified_df = in_scope_df.copy()
        for col in ("Level", "Classification_Reason", "Data_Quality_Flag"):
            classified_df[col] = ""

    if "Level" in classified_df.columns and not classified_df.empty:
        l1 = int((classified_df["Level"] == "Level 1").sum())
        l2 = int((classified_df["Level"] == "Level 2").sum())
        l3 = int((classified_df["Level"] == "Level 3").sum())
        dq = classified_df.get("Data_Quality_Flag", "").astype(str)
        missing = int(dq.str.startswith("MISSING").sum())
        print(f"  Level I:      {l1}")
        print(f"  Level II:     {l2}")
        print(f"  Level III:    {l3}")
        print(f"  Missing data: {missing}")

    print("Assigning Calculation Numbers to Level I lines...")
    classified_df, cn_proposals = assign_cns(classified_df, rules, material_map)
    auto  = sum(1 for p in cn_proposals if p["review_flag"] == "AUTO-CONFIRMED")
    large = sum(1 for p in cn_proposals if p["review_flag"] == "REVIEW-LARGE-CN")
    stand = sum(1 for p in cn_proposals if p["review_flag"] == "REVIEW-STANDALONE")
    print(f"  Proposed CNs: {len(cn_proposals)}  ({auto} auto, {large} large, {stand} standalone)")

    # Merge classification results back into raw_df so scope-excluded rows are preserved
    import pandas as pd
    enriched_df = raw_df.copy()
    for col in ("Level", "Classification_Reason", "Data_Quality_Flag",
                "CN_Number", "CN_Review_Flag"):
        enriched_df[col] = ""

    for idx in classified_df.index:
        for col in ("Level", "Classification_Reason", "Data_Quality_Flag",
                    "CN_Number", "CN_Review_Flag"):
            if col in classified_df.columns:
                enriched_df.at[idx, col] = classified_df.at[idx, col]

    for idx in excluded_df.index:
        label = exclusion_labels.at[idx] if idx in exclusion_labels.index else ""
        if label:
            enriched_df.at[idx, "Data_Quality_Flag"] = f"SCOPE: {label}"

    print(f"Writing output to:      {args.output}")
    try:
        write_enriched_output(
            input_path=args.input,
            output_path=args.output,
            detected_config=detected_config,
            enriched_df=enriched_df,
            cn_proposals=cn_proposals,
        )
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

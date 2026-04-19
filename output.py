"""
output.py — Colored Excel writer for SCLL classified output.

Produces:
  Sheet 1 "Classified Lines": full line list with Level, Reason, Review_Flag,
    CN columns for Level 1 rows, color-coded by row status.
  Sheet 2 "Summary": count breakdown by level/status.
  Sheet 3 "CN Proposals": one row per proposed CN with full grouping detail.
  Sheet 4 "Dashboard": CN summary statistics.

No classification logic here. Receives already-classified DataFrames.
"""

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd


# ─────────────────────────────────────────────────────────────────────────────
# Color palette (ARGB format for openpyxl)
# ─────────────────────────────────────────────────────────────────────────────
COLORS = {
    "Level 1":     "FFFF9999",   # light red
    "Level 2":     "FFFFC266",   # light orange
    "Level 3":     "FF92D050",   # light green
    "EXCLUDED":    "FFD3D3D3",   # light grey
    "NEEDS REVIEW": "FFFFFF99",  # light yellow
    "HEADER":      "FF1F4E79",   # dark blue header
    "SUMMARY_HDR": "FF2E75B6",   # medium blue
}

FONT_WHITE = Font(color="FFFFFFFF", bold=True)
FONT_BOLD  = Font(bold=True)
FONT_DARK  = Font(color="FF000000")

THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)


def _make_fill(hex_color: str) -> PatternFill:
    return PatternFill(fill_type="solid", fgColor=hex_color)


def _apply_row_fill(ws, row_idx: int, fill: PatternFill, num_cols: int) -> None:
    for col_idx in range(1, num_cols + 1):
        cell = ws.cell(row=row_idx, column=col_idx)
        cell.fill = fill
        cell.border = THIN_BORDER


def _auto_size_columns(ws, max_width: int = 50) -> None:
    """Set column widths based on content length."""
    for col_cells in ws.columns:
        max_len = 0
        for cell in col_cells:
            try:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            except Exception:
                pass
        col_letter = get_column_letter(col_cells[0].column)
        ws.column_dimensions[col_letter].width = min(max_len + 2, max_width)


def _row_color_key(row: pd.Series) -> str:
    """Determine the color key for a classified row."""
    scope_result = str(row.get("Scope_Filter_Result", "")).strip()
    if scope_result == "Excluded":
        return "EXCLUDED"
    review = str(row.get("Review_Flag", "")).strip()
    if review == "NEEDS REVIEW":
        return "NEEDS REVIEW"
    level = str(row.get("Level", "")).strip()
    if level in COLORS:
        return level
    return "NEEDS REVIEW"


def write_classified_sheet(
    wb: openpyxl.Workbook,
    classified_df: pd.DataFrame,
    excluded_df: pd.DataFrame,
) -> None:
    """
    Write the 'Classified Lines' sheet.
    In-scope rows first, excluded rows appended at the bottom.
    """
    ws = wb.active
    ws.title = "Classified Lines"

    # Build combined DataFrame
    classified_with_scope = classified_df.copy()
    classified_with_scope["Scope_Filter_Result"] = "In Scope"

    excluded_with_scope = excluded_df.copy()
    excluded_with_scope["Scope_Filter_Result"] = "Excluded"
    excluded_with_scope["Level"] = ""
    excluded_with_scope["Classification_Reason"] = "Line excluded from stress scope"
    excluded_with_scope["Review_Flag"] = "EXCLUDED"

    combined = pd.concat([classified_with_scope, excluded_with_scope], ignore_index=True)

    # Reorder columns: put new columns near the front after line_number and scope
    priority_cols = ["line_number", "scope", "Scope_Filter_Result", "Level",
                     "Classification_Reason", "Review_Flag"]
    other_cols = [c for c in combined.columns if c not in priority_cols]
    ordered_cols = priority_cols + other_cols
    ordered_cols = [c for c in ordered_cols if c in combined.columns]
    combined = combined[ordered_cols]

    # ── Write header row ─────────────────────────────────────────────────────
    header_fill = _make_fill(COLORS["HEADER"])
    for col_idx, col_name in enumerate(combined.columns, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.value = col_name
        cell.fill = header_fill
        cell.font = FONT_WHITE
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = THIN_BORDER

    # ── Write data rows ───────────────────────────────────────────────────────
    for row_idx, (_, data_row) in enumerate(combined.iterrows(), start=2):
        color_key = _row_color_key(data_row)
        fill = _make_fill(COLORS[color_key])

        for col_idx, col_name in enumerate(combined.columns, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            val = data_row[col_name]
            cell.value = "" if pd.isna(val) else val
            cell.fill = fill
            cell.border = THIN_BORDER
            cell.alignment = Alignment(wrap_text=False)

        _apply_row_fill(ws, row_idx, fill, len(combined.columns))

    # ── Formatting ────────────────────────────────────────────────────────────
    ws.freeze_panes = "A2"
    _auto_size_columns(ws)


def write_summary_sheet(
    wb: openpyxl.Workbook,
    classified_df: pd.DataFrame,
    excluded_df: pd.DataFrame,
) -> None:
    """Write the 'Summary' sheet with count breakdowns."""
    ws = wb.create_sheet(title="Summary")

    total_in_scope = len(classified_df)
    total_excluded = len(excluded_df)
    total_all = total_in_scope + total_excluded

    level1_count = int((classified_df.get("Level", pd.Series()) == "Level 1").sum())
    level2_count = int((classified_df.get("Level", pd.Series()) == "Level 2").sum())
    level3_count = int((classified_df.get("Level", pd.Series()) == "Level 3").sum())
    review_count = int((classified_df.get("Review_Flag", pd.Series()) == "NEEDS REVIEW").sum())

    rows = [
        ("SCLL Tool — Classification Summary", None),
        (None, None),
        ("Metric", "Count"),
        ("Total lines received", total_all),
        ("Lines in JESA scope (classified)", total_in_scope),
        ("Lines excluded (Vendor / Client)", total_excluded),
        (None, None),
        ("Level 1 — Rigorous analysis required", level1_count),
        ("Level 2 — Normal analysis required", level2_count),
        ("Level 3 — Visual / approximate check", level3_count),
        ("Flagged: NEEDS REVIEW (missing data)", review_count),
        (None, None),
        ("Color Legend", ""),
        ("Level 1", "Red"),
        ("Level 2", "Orange"),
        ("Level 3", "Green"),
        ("Excluded", "Grey"),
        ("Needs Review", "Yellow"),
    ]

    header_fill  = _make_fill(COLORS["SUMMARY_HDR"])
    l1_fill      = _make_fill(COLORS["Level 1"])
    l2_fill      = _make_fill(COLORS["Level 2"])
    l3_fill      = _make_fill(COLORS["Level 3"])
    excl_fill    = _make_fill(COLORS["EXCLUDED"])
    review_fill  = _make_fill(COLORS["NEEDS REVIEW"])

    color_map = {
        "Level 1": l1_fill,
        "Level 2": l2_fill,
        "Level 3": l3_fill,
        "Excluded": excl_fill,
        "Needs Review": review_fill,
    }

    for r_idx, (label, value) in enumerate(rows, start=1):
        c1 = ws.cell(row=r_idx, column=1, value=label)
        c2 = ws.cell(row=r_idx, column=2, value=value)

        if r_idx == 1:
            c1.font = Font(bold=True, size=13, color="FF1F4E79")
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2)
        elif label == "Metric":
            for cell in (c1, c2):
                cell.fill = header_fill
                cell.font = FONT_WHITE
                cell.alignment = Alignment(horizontal="center")
                cell.border = THIN_BORDER
        elif label in color_map:
            c1.fill = color_map[label]
            c2.fill = color_map[label]
            c1.border = THIN_BORDER
            c2.border = THIN_BORDER
        elif label and label.startswith("Level"):
            c1.border = THIN_BORDER
            c2.border = THIN_BORDER
        elif label and label not in (None, "Color Legend", "SCLL Tool — Classification Summary"):
            c1.border = THIN_BORDER
            c2.border = THIN_BORDER

    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 12


def write_cn_proposals_sheet(wb: openpyxl.Workbook, cn_proposals: list) -> None:
    """Write the 'CN Proposals' sheet — one row per proposed CN."""
    ws = wb.create_sheet(title="CN Proposals")

    headers = [
        "CN Number", "Review Flag", "Lines in CN", "Equipment Tags",
        "Area / Unit", "Min Temp (°C)", "Max Temp (°C)", "Delta T (°C)",
        "Line Count", "Grouping Reason", "Engineer Notes",
    ]

    flag_colors = {
        "[AUTO-CONFIRMED]":     "FF92D050",  # green
        "[REVIEW-TEMPERATURE]": "FFFFFF99",  # yellow
        "[REVIEW-MODEL-SIZE]":  "FFFFC266",  # orange
        "[REVIEW-MISSING-DATA]": "FFFF9999", # red
        "[REVIEW-AREA-CONFLICT]": "FFFFC266",
        "[REVIEW-MANUAL]":      "FFFFFF99",
    }

    header_fill = _make_fill(COLORS["HEADER"])
    for col_idx, hdr in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=hdr)
        cell.fill = header_fill
        cell.font = FONT_WHITE
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = THIN_BORDER

    for row_idx, proposal in enumerate(cn_proposals, start=2):
        flag   = proposal.get("review_flag", "")
        color  = flag_colors.get(flag, "FFFFFFFF")
        fill   = _make_fill(color)

        values = [
            proposal.get("cn_number", ""),
            flag,
            ", ".join(proposal.get("line_numbers", [])),
            ", ".join(proposal.get("equipment_tags", [])),
            ", ".join(proposal.get("area_codes", [])) or "—",
            proposal.get("min_temperature", ""),
            proposal.get("max_temperature", ""),
            proposal.get("delta_t", ""),
            proposal.get("line_count", ""),
            proposal.get("grouping_reason", ""),
            "",  # Engineer Notes — blank for engineer to fill
        ]

        for col_idx, val in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.value = "" if val is None else val
            cell.fill = fill
            cell.border = THIN_BORDER
            cell.alignment = Alignment(wrap_text=(col_idx == 10))  # wrap reason column

    ws.freeze_panes = "A2"
    _auto_size_columns(ws)
    # Widen the reason and lines columns
    ws.column_dimensions["C"].width = 40
    ws.column_dimensions["J"].width = 60
    ws.column_dimensions["K"].width = 30


def write_dashboard_sheet(
    wb: openpyxl.Workbook,
    classified_df: pd.DataFrame,
    cn_proposals: list,
) -> None:
    """Write the 'Dashboard' sheet with CN summary statistics."""
    ws = wb.create_sheet(title="Dashboard")

    l1_total = int((classified_df.get("Level", pd.Series()) == "Level 1").sum())
    total_cns = len(cn_proposals)

    flag_counts: dict[str, int] = {}
    for p in cn_proposals:
        flag = p.get("review_flag", "")
        flag_counts[flag] = flag_counts.get(flag, 0) + 1

    auto_confirmed  = flag_counts.get("[AUTO-CONFIRMED]", 0)
    rev_temp        = flag_counts.get("[REVIEW-TEMPERATURE]", 0)
    rev_size        = flag_counts.get("[REVIEW-MODEL-SIZE]", 0)
    rev_missing     = flag_counts.get("[REVIEW-MISSING-DATA]", 0)
    rev_area        = flag_counts.get("[REVIEW-AREA-CONFLICT]", 0)
    rev_manual      = flag_counts.get("[REVIEW-MANUAL]", 0)
    total_flagged   = total_cns - auto_confirmed

    assigned_lines   = sum(p["line_count"] for p in cn_proposals if "[REVIEW-MISSING-DATA]" not in p["review_flag"])
    unassigned_lines = sum(p["line_count"] for p in cn_proposals if "[REVIEW-MISSING-DATA]" in p["review_flag"])

    rows = [
        ("SCLL Tool — CN Assignment Dashboard", None, None),
        (None, None, None),
        ("Metric", "Count", "Notes"),
        ("Total Level 1 lines",            l1_total,        "Eligible for CN assignment"),
        ("Lines with CN assigned",         assigned_lines,  ""),
        ("Lines unassigned (missing data)", unassigned_lines, "Engineer action required"),
        (None, None, None),
        ("Total proposed CNs",             total_cns,       ""),
        ("CNs auto-confirmed",             auto_confirmed,  "All boundary rules applied cleanly"),
        ("CNs flagged for review",         total_flagged,   "Engineer must approve before analysis"),
        (None, None, None),
        ("  — REVIEW-TEMPERATURE",         rev_temp,        f"Delta T > threshold between connected lines"),
        ("  — REVIEW-MODEL-SIZE",          rev_size,        "CN exceeds max lines per model"),
        ("  — REVIEW-MISSING-DATA",        rev_missing,     "Equipment tags missing"),
        ("  — REVIEW-AREA-CONFLICT",       rev_area,        "Lines from different areas in same network"),
        ("  — REVIEW-MANUAL",              rev_manual,      "Single line, P&ID confirmation needed"),
        (None, None, None),
        ("CN Status Legend", "", ""),
        ("PROPOSED",   "", "Initial tool output — engineer has not reviewed yet"),
        ("CONFIRMED",  "", "Engineer approved the CN grouping"),
        ("REVISED",    "", "Engineer changed the grouping"),
    ]

    header_fill   = _make_fill(COLORS["SUMMARY_HDR"])
    title_font    = Font(bold=True, size=13, color="FF1F4E79")
    section_fill  = _make_fill("FFE2EFDA")   # light green tint for confirmed
    flag_fill     = _make_fill(COLORS["NEEDS REVIEW"])
    auto_fill     = _make_fill(COLORS["Level 3"])

    for r_idx, (label, value, note) in enumerate(rows, start=1):
        c1 = ws.cell(row=r_idx, column=1, value=label)
        c2 = ws.cell(row=r_idx, column=2, value=value)
        c3 = ws.cell(row=r_idx, column=3, value=note)

        if r_idx == 1:
            c1.font = title_font
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3)
        elif label == "Metric":
            for cell in (c1, c2, c3):
                cell.fill = header_fill
                cell.font = FONT_WHITE
                cell.alignment = Alignment(horizontal="center")
                cell.border = THIN_BORDER
        elif label and "REVIEW" in str(label):
            for cell in (c1, c2, c3):
                cell.fill = flag_fill
                cell.border = THIN_BORDER
        elif label == "CNs auto-confirmed":
            for cell in (c1, c2, c3):
                cell.fill = auto_fill
                cell.border = THIN_BORDER
        elif label and label not in (None, "CN Status Legend", "SCLL Tool — CN Assignment Dashboard"):
            for cell in (c1, c2, c3):
                cell.border = THIN_BORDER

    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 45


def write_output(
    classified_df: pd.DataFrame,
    excluded_df: pd.DataFrame,
    output_path: str,
    rules: dict,
    cn_proposals: list | None = None,
) -> None:
    """
    Create the output workbook.
    Sheets written:
      1. Classified Lines
      2. Summary
      3. CN Proposals  (only if cn_proposals is non-empty)
      4. Dashboard     (only if cn_proposals is non-empty)
    """
    if cn_proposals is None:
        cn_proposals = []

    wb = openpyxl.Workbook()
    write_classified_sheet(wb, classified_df, excluded_df)
    write_summary_sheet(wb, classified_df, excluded_df)
    if cn_proposals:
        write_cn_proposals_sheet(wb, cn_proposals)
        write_dashboard_sheet(wb, classified_df, cn_proposals)

    try:
        wb.save(output_path)
    except PermissionError:
        raise PermissionError(
            f"Cannot write to '{output_path}'. "
            f"If the file is open in Excel, please close it first."
        )

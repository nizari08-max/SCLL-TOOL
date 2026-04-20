"""
app.py — Flask web server for the SCLL Tool.

Three-step flow (matches the 3-step UI):
    1. UPLOAD      — user drops any .xlsx; we auto-detect format
    2. PROGRESS    — SSE streams pipeline logs
    3. RESULTS     — table, CN cards, inline download of enriched Excel

Routes:
    GET  /                      Single-page UI
    POST /upload                Accept .xlsx, run format_detector, return summary
    POST /run                   Kick off the pipeline (returns job_id)
    GET  /job-status/<job_id>   SSE stream of log messages
    GET  /results/<job_id>      Results JSON after job completes
    GET  /download/<filename>   Serve the enriched output file

Run with:
    python app.py
"""

from __future__ import annotations

import json
import os
import sys
import threading
import traceback
import uuid
import warnings
from datetime import datetime

import pandas as pd
from flask import (
    Flask, Response, jsonify, render_template,
    request, send_file, stream_with_context,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = os.path.join(BASE_DIR, "uploads")
app.config["OUTPUT_FOLDER"] = os.path.join(BASE_DIR, "outputs")
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
os.makedirs(app.config["OUTPUT_FOLDER"], exist_ok=True)

jobs: dict = {}

CN_COLORS = [
    "#3b82f6", "#ef4444", "#22c55e", "#f59e0b", "#8b5cf6",
    "#06b6d4", "#f97316", "#ec4899", "#10b981", "#6366f1",
    "#84cc16", "#14b8a6", "#a855f7", "#eab308", "#0ea5e9",
    "#d946ef", "#f43f5e", "#34d399", "#fbbf24", "#60a5fa",
]


# ─────────────────────────────────────────────────────────────────────────────
# Routes
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400

    f = request.files["file"]
    if not f.filename or not f.filename.lower().endswith(".xlsx"):
        return jsonify({"error": "Only .xlsx files are accepted"}), 400

    safe_name = f"{uuid.uuid4().hex}_{f.filename}"
    filepath = os.path.join(app.config["UPLOAD_FOLDER"], safe_name)
    f.save(filepath)

    try:
        from format_detector import detect_format

        detected_config, summary = detect_format(filepath)
        badges = _build_detection_badges(detected_config)
        project_code = _guess_project_code(f.filename)

        return jsonify({
            "filename":            safe_name,
            "original_name":       f.filename,
            "row_count":           detected_config.get("row_count", "?"),
            "sheet_names":         detected_config.get("sheet_names", []),
            "detected_config":     detected_config,
            "detection_summary":   summary,
            "detection_badges":    badges,
            "not_found_columns":   detected_config.get("not_found_columns", []),
            "project_code":        project_code,
        })
    except Exception as exc:
        try:
            os.unlink(filepath)
        except Exception:
            pass
        return jsonify({"error": f"Failed to process file: {exc}"}), 400


@app.route("/run", methods=["POST"])
def run_analysis():
    data = request.get_json(force=True) or {}
    filename = data.get("filename")
    if not filename:
        return jsonify({"error": "No filename provided"}), 400

    input_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
    if not os.path.exists(input_path):
        return jsonify({"error": "Uploaded file not found — please re-upload"}), 404

    project_code    = (data.get("project_code") or "").strip()
    detected_config = data.get("detected_config") or {}

    job_id = uuid.uuid4().hex
    output_filename = f"scll_output_{job_id[:8]}.xlsx"
    output_path = os.path.join(app.config["OUTPUT_FOLDER"], output_filename)

    jobs[job_id] = {"status": "running", "messages": [], "result": None, "error": None}

    t = threading.Thread(
        target=_run_pipeline,
        args=(job_id, input_path, output_path, project_code, detected_config),
        daemon=True,
    )
    t.start()

    return jsonify({"job_id": job_id})


@app.route("/job-status/<job_id>")
def job_status_sse(job_id):
    def generate():
        import time
        last_idx = 0
        while True:
            job = jobs.get(job_id)
            if not job:
                yield f"data: {json.dumps({'type': 'error', 'message': 'Job not found'})}\n\n"
                return

            msgs = job.get("messages", [])
            while last_idx < len(msgs):
                payload = json.dumps({"type": "log", "message": msgs[last_idx]})
                yield f"data: {payload}\n\n"
                last_idx += 1

            if job["status"] == "done":
                yield f"data: {json.dumps({'type': 'done', 'job_id': job_id})}\n\n"
                return
            if job["status"] == "error":
                err = job.get("error", "Unknown error")
                yield f"data: {json.dumps({'type': 'error', 'message': err})}\n\n"
                return

            time.sleep(0.15)

    return Response(
        stream_with_context(generate()),
        content_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.route("/results/<job_id>")
def get_results(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    if job["status"] == "error":
        return jsonify({"error": job.get("error", "Pipeline error")}), 500
    if job["status"] != "done":
        return jsonify({"error": "Job still running", "status": job["status"]}), 202
    return jsonify(job["result"])


@app.route("/download/<path:filename>")
def download_file(filename):
    output_folder = os.path.abspath(app.config["OUTPUT_FOLDER"])
    filepath      = os.path.abspath(os.path.join(output_folder, filename))
    if not filepath.startswith(output_folder + os.sep) and filepath != output_folder:
        return jsonify({"error": "Invalid path"}), 403
    if not os.path.exists(filepath):
        return jsonify({"error": "File not found"}), 404
    return send_file(filepath, as_attachment=True, download_name=os.path.basename(filepath))


# ─────────────────────────────────────────────────────────────────────────────
# Pipeline
# ─────────────────────────────────────────────────────────────────────────────

def _run_pipeline(job_id, input_path, output_path, project_code, detected_config):
    msgs = jobs[job_id]["messages"]

    def log(msg):
        msgs.append(msg)

    try:
        from parser import load_rules, load_material_map, read_linelist, filter_scope
        from classifier import classify_dataframe
        from cn_assigner import assign_cns
        from format_detector import apply_detection_to_rules, load_mapping
        from output import write_enriched_output

        log("⟳ Loading configuration...")
        rules = load_rules(os.path.join(BASE_DIR, "rules.yaml"))
        mapping = load_mapping()
        material_map = load_material_map(os.path.join(BASE_DIR, "material_mapping.yaml"))

        # Inject detection-driven fields into rules (detection_mode, keyword_patterns, size_unit)
        rules = apply_detection_to_rules(rules, detected_config, mapping)

        if project_code and rules.get("cn_settings"):
            rules["cn_settings"]["project_code"] = project_code

        log(f"✓ Rules + mapping loaded ({len(detected_config.get('column_mappings', {}))} columns mapped)")

        # ── Parse ────────────────────────────────────────────────────────────
        log("⟳ Reading input file...")
        with warnings.catch_warnings(record=True) as warn_list:
            warnings.simplefilter("always")
            raw_df = read_linelist(input_path, detected_config, rules)
        for w in warn_list:
            log(f"  ⚠ {w.message}")
        log(f"✓ Parsed {len(raw_df)} rows × {len(raw_df.columns)} columns")

        # ── Scope filter ─────────────────────────────────────────────────────
        in_scope_df, excluded_df, exclusion_labels = filter_scope(raw_df, detected_config)
        log(f"✓ Scope filter — {len(in_scope_df)} in scope, {len(excluded_df)} excluded")

        # ── Classify ─────────────────────────────────────────────────────────
        log(f"⟳ Classifying {len(in_scope_df)} in-scope lines...")
        if not in_scope_df.empty:
            classified_df = classify_dataframe(in_scope_df, material_map, rules)
        else:
            classified_df = in_scope_df.copy()
            for col in ("Level", "Classification_Reason", "Data_Quality_Flag"):
                classified_df[col] = ""

        level_series = classified_df.get("Level", pd.Series([], dtype=str)).astype(str)
        l1 = int((level_series == "Level 1").sum())
        l2 = int((level_series == "Level 2").sum())
        l3 = int((level_series == "Level 3").sum())
        dq_series = classified_df.get("Data_Quality_Flag", pd.Series([], dtype=str)).astype(str)
        missing = int(dq_series.str.startswith("MISSING").sum())
        log(f"✓ Classification — I: {l1} | II: {l2} | III: {l3} | Missing data: {missing}")

        # ── CN assignment (Level 1 only) ─────────────────────────────────────
        log("⟳ Assigning Calculation Numbers to Level I lines...")
        classified_df, cn_proposals = assign_cns(classified_df, rules, material_map)

        auto   = sum(1 for p in cn_proposals if p["review_flag"] == "AUTO-CONFIRMED")
        large  = sum(1 for p in cn_proposals if p["review_flag"] == "REVIEW-LARGE-CN")
        stand  = sum(1 for p in cn_proposals if p["review_flag"] == "REVIEW-STANDALONE")
        log(f"✓ CN assignment — {len(cn_proposals)} CNs ({auto} auto, {large} large, {stand} standalone)")

        # ── Build enriched_df: raw_df with 5 new columns merged in ───────────
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

        # ── Write enriched Excel ─────────────────────────────────────────────
        log("⟳ Writing enriched output file...")
        write_enriched_output(
            input_path=input_path,
            output_path=output_path,
            detected_config=detected_config,
            enriched_df=enriched_df,
            cn_proposals=cn_proposals,
        )
        log(f"✓ Output ready — {os.path.basename(output_path)}")

        # ── Build result payload for UI ──────────────────────────────────────
        total      = len(raw_df)
        excl_count = len(excluded_df)
        scope_count = len(in_scope_df)

        classifications = []
        for _, row in classified_df.iterrows():
            classifications.append({
                "line_number":        _s(row.get("line_number")),
                "size":               _s(row.get("size")),
                "material":           _s(row.get("material")),
                "design_temperature": _s(row.get("design_temperature")),
                "level":              _s(row.get("Level")),
                "reason":             _s(row.get("Classification_Reason")),
                "data_quality":       _s(row.get("Data_Quality_Flag")),
                "from_equipment":     _s(row.get("from_equipment")),
                "to_equipment":       _s(row.get("to_equipment")),
                "cn_number":          _s(row.get("CN_Number")),
                "cn_review_flag":     _s(row.get("CN_Review_Flag")),
                "fluid_service":      _s(row.get("fluid_service")),
            })

        cns = []
        for i, p in enumerate(cn_proposals):
            min_t = p.get("min_temperature"); max_t = p.get("max_temperature")
            temp_range = (f"{min_t:.0f}°C – {max_t:.0f}°C"
                          if (min_t is not None and max_t is not None) else "—")
            cns.append({
                "cn_number":       p["cn_number"],
                "line_count":      p["line_count"],
                "review_flag":     p["review_flag"],
                "grouping_reason": p.get("grouping_reason", ""),
                "line_numbers":    p.get("line_numbers", []),
                "equipment_tags":  list(p.get("equipment_tags", [])),
                "temp_range":      temp_range,
                "delta_t":         round(p["delta_t"], 1) if p.get("delta_t") is not None else None,
                "color":           CN_COLORS[i % len(CN_COLORS)],
            })

        jobs[job_id]["result"] = {
            "output_filename": os.path.basename(output_path),
            "stats": {
                "total":      total,
                "excluded":   excl_count,
                "in_scope":   scope_count,
                "level1":     l1,
                "level2":     l2,
                "level3":     l3,
                "missing":    missing,
                "cn_count":   len(cn_proposals),
                "cn_auto":    auto,
                "cn_large":   large,
                "cn_standalone": stand,
            },
            "classifications": classifications,
            "cns":             cns,
            "generated_at":    datetime.now().isoformat(),
            "project_code":    project_code,
        }
        jobs[job_id]["status"] = "done"
        log("✓ Analysis complete")

    except Exception as exc:
        jobs[job_id]["status"] = "error"
        jobs[job_id]["error"]  = str(exc)
        msgs.append(f"✗ Error: {exc}")
        msgs.append(traceback.format_exc())


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _build_detection_badges(cfg: dict) -> list[dict]:
    badges = []

    header_excel = cfg.get("header_row", 0) + 1
    badges.append({"label": f"Row {header_excel}", "icon": "✓", "type": "ok",
                   "title": f"Header detected at Excel row {header_excel}"})

    size_unit = cfg.get("size_unit", "inches")
    if size_unit == "mm":
        badges.append({"label": "mm→NPS", "icon": "✓", "type": "ok",
                       "title": "Size column in mm — will be converted to NPS inches"})
    else:
        badges.append({"label": "NPS inches", "icon": "✓", "type": "ok",
                       "title": "Size column already in NPS inches"})

    scope_mode = cfg.get("scope_mode", "assume_all_in_scope")
    scope_labels = {
        "include_values":        "Scope column",
        "text_keywords":         "Keyword scope",
        "column_exclude_values": "Exclude values",
        "assume_all_in_scope":   "⚠ No scope col",
    }
    scope_type = "warning" if scope_mode == "assume_all_in_scope" else "ok"
    badges.append({"label": scope_labels.get(scope_mode, scope_mode),
                   "icon": "✓" if scope_type == "ok" else "⚠",
                   "type": scope_type, "title": f"Scope mode: {scope_mode}"})

    equip_mode = cfg.get("equipment_mode", "tag_prefix")
    badges.append({"label": "Keyword equip" if equip_mode == "keyword" else "Tag prefix",
                   "icon": "✓", "type": "ok",
                   "title": f"Equipment detection: {equip_mode}"})

    mapped   = len(cfg.get("column_mappings", {}))
    missing  = len(cfg.get("not_found_columns", []))
    total    = mapped + missing
    col_type = "warning" if missing > 5 else "ok"
    badges.append({"label": f"{mapped}/{total} cols",
                   "icon": "✓" if col_type == "ok" else "⚠",
                   "type": col_type,
                   "title": f"{mapped} columns mapped, {missing} not found"})

    row_count = cfg.get("row_count", 0)
    badges.append({"label": f"{row_count} lines", "icon": "✓", "type": "ok",
                   "title": f"{row_count} data rows detected"})

    return badges


def _guess_project_code(filename: str) -> str:
    import re
    m = re.search(r"\b([A-Z]{1,2}\d{4,6})\b", filename, re.IGNORECASE)
    return m.group(1).upper() if m else ""


def _s(val) -> str:
    if val is None:
        return ""
    try:
        import math
        if isinstance(val, float) and math.isnan(val):
            return ""
    except Exception:
        pass
    return str(val)


# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print()
    print("  +------------------------------------------+")
    print("  |   SCLL Tool  --  Web Interface           |")
    print("  |   Format-Agnostic Line List Classifier   |")
    print("  +------------------------------------------+")
    print("  |   http://localhost:5000                  |")
    print("  +------------------------------------------+")
    print()
    app.run(debug=True, port=5000, threaded=True, use_reloader=False)

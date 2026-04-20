# CLAUDE.md — Behavioral Guide for Claude Code Sessions

## Project Identity
**SCLL Tool** — Stress Critical Line List automated classifier for piping stress engineers.
Deterministic, rule-based, format-agnostic Python tool. No AI inference. No guessing.

## Core Philosophy (format-agnostic, not file-specific)
The tool understands the engineering LOGIC and applies it to ANY line list from ANY
project/client. `format_detector.py` reads the workbook and produces a `detected_config`
that drives the rest of the pipeline — no per-project rules files.

If a column the engineering logic expects is missing from the input, the rule that
depends on it is **skipped** and the affected rows get a `Data_Quality_Flag`. The tool
NEVER rejects a file for "unknown format".

## Architecture — Which File Does What
| File | Responsibility |
|------|---------------|
| `scll_tool.py`          | CLI entry point — orchestration only |
| `app.py`                | Flask web server — 3-step flow, SSE streaming |
| `format_detector.py`    | Auto-detects header row, units row, column mappings, size unit, scope mode, equipment mode. Writes into `detected_config`. |
| `parser.py`             | Reads Excel via `detected_config`, converts mm→NPS, filters scope. No classification. |
| `classifier.py`         | 5-step classification waterfall. Emits `Level`, `Classification_Reason`, `Data_Quality_Flag`. |
| `cn_assigner.py`        | CN grouping on Level 1 lines (6-rule spec). Emits `CN_Number`, `CN_Review_Flag`. |
| `output.py`             | Opens engineer's original workbook and appends **5** new columns + SUMMARY + CN PROPOSALS sheets. |
| `rules.yaml`            | Engineering data only: charts, thresholds, exception flags, CN settings. No column names, no project knowledge. |
| `mapping.yaml`          | Column-name pattern matcher used by `format_detector.py`. Engineer-editable. |
| `material_mapping.yaml` | Pipe-class code → material group (CS / SS / FRP). |
| `templates/index.html`  | Single-page web UI (Upload&Detect → Progress → Results with inline download). |

## Non-Negotiable Rules

### 1. No hardcoded numeric values in `.py` files
Temperature thresholds, size limits, pressure limits live in `rules.yaml`.
Never write `if temp > 250` in Python — read from `rules["chart_1"]`.

### 2. No hardcoded pipe-class codes in `.py` files
All pipe class → material group mappings live in `material_mapping.yaml`.
Never write `if material == "BA1"` in Python.

### 3. No hardcoded column names in `.py` files
Column patterns live in `mapping.yaml`. `format_detector.py` matches them against the
workbook; `parser.py` renames columns to internal field names. After parsing, all
Python code uses the internal names (`size`, `design_temperature`, `material`, etc.).

### 4. No project-specific rules files
The tool must work on ANY line list using `rules.yaml` + `mapping.yaml` only. If a
project needs different patterns, update `mapping.yaml` (patterns list) — do NOT create
`rules_<project>.yaml`.

### 5. Classification is a strict 5-step waterfall — never reorder
Step 1 → 2 → 3 → 4 → 5. Each step returns immediately once a level is assigned.
Later steps must NEVER override an earlier one.

### 6. Every classified row must have a non-empty `Classification_Reason`
The reason names which rule fired and which Step.

### 7. Missing data → `Data_Quality_Flag`, never assume
If a required field is blank:
- Set `Data_Quality_Flag = "MISSING: <field_list>"` (or `"AMBIGUOUS: ..."`)
- Leave `Level = ""`
- Do not guess, interpolate, or infer a value
- If a whole rule depends on a column that doesn't exist in the file, skip the rule
  silently (don't flag every row — only flag rows where the specific cell is empty)

### 8. Excluded rows are kept — not deleted
Scope-excluded lines (vendor/client/licensor) stay in the output, greyed out, with
`Data_Quality_Flag = "SCOPE: VENDOR"` (or CLIENT, LICENSOR, etc.). Never drop them.

## Output Philosophy — Enrich, Don't Replace
The output Excel **IS the engineer's original file**, enriched. Five new columns are
appended AFTER all original columns, preserving every original sheet and column:

| # | Column                | Content                                                      |
|---|-----------------------|--------------------------------------------------------------|
| 1 | CRITICALITY LEVEL     | `I` / `II` / `III` (Roman), color-coded                      |
| 2 | CLASSIFICATION REASON | Full text of which rule fired                                |
| 3 | CN NUMBER             | `CN-001`, `CN-002`, … (Level I only)                         |
| 4 | CN REVIEW FLAG        | `AUTO-CONFIRMED` / `REVIEW-LARGE-CN` / `REVIEW-STANDALONE`   |
| 5 | DATA QUALITY FLAG     | `""` / `SCOPE: VENDOR` / `MISSING: <fields>` / `AMBIGUOUS: …` |

SCOPE is **folded into DATA QUALITY FLAG** — there is no separate SCOPE column.

Two new sheets are appended: **SUMMARY** and **CN PROPOSALS**.

## CN Review Flags
`AUTO-CONFIRMED` / `REVIEW-LARGE-CN` / `REVIEW-STANDALONE` — no square brackets.

## Web App (3-step flow)
Run: `python app.py` then open http://localhost:5000
1. **UPLOAD & DETECT** — drop `.xlsx`, auto-detect format, show badges, confirm project code
2. **ANALYSIS** — SSE-streamed progress log
3. **RESULTS** — stats bar, classified lines table (filter/search/paginate), CN cards, inline download button

No rules-file selector. No mode selector. No separate download step.

## Testing
After any change to `classifier.py`, `parser.py`, `cn_assigner.py`, `format_detector.py`,
`rules.yaml`, `mapping.yaml`, or `material_mapping.yaml`:

```
python scll_tool.py --input "Q37027-02-A00-PE-LST-00010 (1) (1).xlsx" --output /tmp/scll_out.xlsx
```
Expected (2026-04-20): Total=907, Excluded=0, L1=441, L2=105, L3=326, Missing=35, CNs=139
(63 auto, 5 large, 71 standalone).

If counts change unexpectedly, investigate before committing.

## What NOT To Do
- No AI classification — logic must be 100% deterministic from rules.yaml
- No silent decisions — every output must be traceable to a named rule
- No assumptions on missing data — flag it via `Data_Quality_Flag`
- No project-specific rules files (`rules_<project>.yaml` is forbidden)
- No reintroducing `output_jesa.py` or any second output path — `output.py` is the only writer
- No reintroducing the SCOPE column — SCOPE is folded into DATA QUALITY FLAG
- No creating additional helper files unless the user asks for them

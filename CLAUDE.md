# CLAUDE.md — Behavioral Guide for Claude Code Sessions

## Project Identity
**SCLL Tool** — Stress Critical Line List automated classifier for JESA piping stress engineers.
This is a deterministic, rule-based Python CLI tool. No AI inference. No guessing. No assumptions.

## Architecture — Which File Does What
| File | Responsibility |
|------|---------------|
| `scll_tool.py` | CLI entry point — orchestration only, no domain logic |
| `parser.py` | Reads Excel input, validates columns, filters scope, mm→NPS conversion — no classification |
| `classifier.py` | 5-step classification engine — no file I/O, no Excel operations |
| `output.py` | Writes color-coded Excel output — no classification logic |
| `rules.yaml` | Generic rules template — column mappings, chart tables, thresholds |
| `rules_q83010.yaml` | Project-specific rules for Q83010 (multi-row header, mm sizes, keyword detection) |
| `material_mapping.yaml` | Pipe class code → material group — editable by engineers |

## Non-Negotiable Rules

### 1. Never hardcode numeric values in .py files
All temperature thresholds, size limits, and pressure values must live in `rules.yaml`.
Never write `if temp > 250` in Python. Always read from `rules["chart_1"]` etc.

### 2. Never hardcode pipe class codes in .py files
All pipe class → material group mappings live in `material_mapping.yaml`.
Never write `if material == "BA1"` in Python.

### 3. Classification is a strict 5-step waterfall — never reorder
Steps 1 → 2 → 3 → 4 → 5. Each step returns immediately if a level is assigned.
A later step must NEVER override an earlier one.

### 4. Never classify without a written reason
Every classified row must have a non-empty `Classification_Reason` column.
The reason must name which rule fired and which Step.

### 5. Missing data → NEEDS REVIEW, never assume
If `size`, `design_temperature`, or `material` is blank:
- Set `Review_Flag = "NEEDS REVIEW"`, `Level = ""`
- Write a reason naming the missing field(s)
- Do not attempt to guess, interpolate, or infer a value

### 6. Column names come from rules.yaml — never hardcoded in Python
The Excel header for every field is defined in `rules.yaml['column_mappings']`.
`parser.py` renames columns during load. After that, all Python code uses internal names.

### 7. Excluded rows are kept — not deleted
Lines with Scope != JESA are greyed out in output Sheet 1 at the bottom.
They appear in the excluded count on Sheet 2. Never drop them.

## Testing Requirement
After ANY change to `classifier.py`, `parser.py`, `rules.yaml`, or `material_mapping.yaml`, run both:

**Generic test suite (prefix detection mode):**
```
python scll_tool.py --input test_linelist.xlsx --output test_output.xlsx
```
Expected: Total=22, Excluded=2, L1=13, L2=3, L3=2, Review=2

**Real project file (keyword detection mode):**
```
python scll_tool.py --input "Q83010-00-00-PR-LST-00003.xlsx" --output Q83010_scll_output.xlsx --rules rules_q83010.yaml
```
Expected: Total=179, Excluded=78, In-scope=101, L1=47, L2=38, L3=3, Review=13

If counts change unexpectedly, investigate before committing.

## Project-Specific Rules Files
Each project with a non-standard Excel layout gets its own `rules_<project>.yaml`.
The generic `rules.yaml` is the template — never modify it for project-specific needs.

Key settings that vary by project:
- `excel_config` — sheet name, header row, skip rows (for multi-row headers)
- `size_config.unit` — `"mm"` or `"inches"` (with `mm_to_nps` lookup if mm)
- `scope.mode` — `"include_values"` (default) or `"text_keywords"`
- `strain_sensitive_equipment.detection_mode` — `"prefix"` (default) or `"keyword"`

## What NOT To Do
- No AI classification — logic must be 100% deterministic from rules.yaml
- No silent decisions — every output must be traceable to a named rule
- No assumptions on missing data — flag it
- No schema changes to rules.yaml without updating CONTEXT.md and test_linelist.xlsx
- No cleaning up / reformatting rules.yaml without explicit instruction from the user
- No creating additional helper files unless the user asks for them

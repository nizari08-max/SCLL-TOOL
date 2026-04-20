# CONTEXT.md — Engineering Knowledge Base

## Specification Reference
- Document: QW2507-00-PE-SPC-00007, Rev A, 31/10/2025
- Title: General Specification for Critical Line List
- Project: DESALINATION WAVE II EAST EXTENSION – JORF LASFAR (QW2507)
- Customer: OCP SA / JESA

The engineering rules below come from this specification. The tool itself is
format-agnostic and works on line lists from any project, not just QW2507.

## Classification Level Definitions
| Level | Name             | Analysis Method                                                              |
|-------|------------------|------------------------------------------------------------------------------|
| 1     | Rigorous         | Computer analysis (Caesar II or equivalent) — required for critical lines    |
| 2     | Normal           | Manual calculations — guided cantilever, Tube Turns, JESA charts             |
| 3     | Visual / Approx  | Visual inspection or approximate methods by Piping Design Dept               |

## Classification Logic Summary (5-Step Waterfall)

### Step 1 — FRP / GRE check
If material group == FRP → **Level 1** always.
Reason: GRE/FRP is a special case evaluated by the manufacturer.

### Step 2 — Exception flags → Level 1 (spec §3.2.5)
All 17+ boolean flag checks. See `rules.yaml['exception_flags']` for the full list.
Key exceptions:
- Relief lines with inlet pressure > 10 barg
- Cement-brick / refractory / glass-lined pipe
- Vertical tower connections ≥ 4"
- Jacketed piping (differential temp)
- Expansion joints, vibration, settlement, vacuum
- PED Category III ≥ 4"
- Underground T ≥ 100°C
- Category M service (ASME B31.3)
- Schedule 160 / OD/WT < 10
- Client request

**If a flag column is absent from the input file**, the rule is silently skipped
(no flag is fired for missing optional columns — only specific rows with genuinely
blank required fields get `Data_Quality_Flag`).

### Step 3 — Strain-sensitive equipment override (Chart 3)
Triggered if FROM or TO tag matches a strain-sensitive equipment prefix/keyword.
Decision matrix (applied to the most critical equipment type found):
- Centrifugal / reciprocating pump → **Level 1** (any size, any temp)
- Any strain-sensitive, T ≥ 100°C → **Level 1**
- Any strain-sensitive, size ≤ 6", T ≥ 50°C → **Level 2**
- Any strain-sensitive, size > 6", T < 100°C → **Level 2**
- Any strain-sensitive, size ≤ 6", T < 50°C → **Level 3**

Strain-sensitive equipment types: rotating equipment (centrifugal pumps, compressors,
turbines), reciprocating pumps/compressors, heaters/boilers, air-cooled heat exchangers,
aluminum/bronze/cast iron equipment, brick/glass/refractory-lined thin-wall vessels,
shell-and-tube exchangers with shell bellows, plate-type / printed-circuit heat
exchangers, vendors specifying low nozzle loading.

### Step 4 — Material chart lookup
Chart selection:
- FROM or TO has non-strain-sensitive equipment tag → use **chart_4** (relaxed)
- CS / Low-Alloy Steel → **chart_1**
- SS / Duplex / Other alloy → **chart_2**

Each chart is a step-function of (diameter, temperature). See `rules.yaml` for thresholds.

### Step 5 — Missing data → Data Quality Flag
If `size`, `design_temperature`, or `material` is blank on a specific row:
- `Data_Quality_Flag = "MISSING: <fields>"`
- `Level = ""`
- Classification reason names the missing field(s)

## Material Groups (see `material_mapping.yaml`)

### CS — Carbon Steel / Low-Alloy Steel (chart_1)
BA1, BA2, BB1, BB1U, BB2, BB3, BB4, BB5, BB6, BP1, CB1, CB2, CB3, CJ1, CP1, **CS**,
FB1, FJ1. Also: MOC codes **CSG** (galvanized), **CS300**, **RLCS** (rubber-lined).
- CJ1 / FJ1 are Low Alloy Steel but use the CS chart (similar thermal properties)
- BP1 is CS but if also jacketed (`jacketed=True`), the jacketed exception fires in Step 2

### SS — Stainless Steel / Duplex / Alloy Metals (chart_2)
BD1, BD2, BD2U, BG1, BG2, BS1, BS2, BS3, CD1, CD2, CG1, CG2, CS1, ES2, GC1, GD1, GD2,
KD1. Also: MOC code **PTFECS** (PTFE-lined CS — treat as other metallic).

### FRP — Fiber-Reinforced Plastic / GRE (always Level 1)
BK1, BK2, BK3, BK4, BV1, CK3. Also: **HDPE**, **FRP** (MOC codes).

## Format Detection (`format_detector.py`)
Replaces all per-project rules files. Reads ANY `.xlsx` and emits a `detected_config`:

| Field                    | Strategy                                                                    |
|--------------------------|-----------------------------------------------------------------------------|
| `sheet_name`             | First sheet that contains a header-row match                                 |
| `header_row`             | Scans rows 0–15; scores each row's cells against `mapping.yaml` patterns     |
| `skip_rows`              | Row after header if ≥50% cells match unit-regex (°C, barg, mm, in, NPS…)     |
| `column_mappings`        | Greedy score matrix between `mapping.yaml` patterns and actual header cells  |
| `size_unit`              | Reads first 20 size values; if any ≥15 and no NPS fractions → `"mm"`         |
| `scope_mode`             | Tries: dedicated scope col → notes col → material col → `assume_all_in_scope`|
| `equipment_mode`         | `"keyword"` if FROM cells contain multi-word names; else `"tag_prefix"`      |
| `not_found_columns`      | Internal-field names that had no match — rules depending on these are skipped|

`apply_detection_to_rules(rules, detected_config, mapping)` injects the detection
results into a deep-copy of `rules` (size unit, equipment mode, patterns) before
the pipeline runs. `rules.yaml` itself is never mutated on disk.

## Scope Filter
All four modes flow through `parser.filter_scope(df, detected_config)`, returning
`(in_scope_df, excluded_df, exclusion_labels)`. Excluded rows get
`Data_Quality_Flag = "SCOPE: <LABEL>"` (e.g. `SCOPE: VENDOR`) and are greyed in the
output.

| Mode                    | When selected                                        | Behavior                                      |
|-------------------------|------------------------------------------------------|-----------------------------------------------|
| `include_values`        | Dedicated scope column with a known set of values    | Keep rows whose value is in the include set   |
| `text_keywords`         | Free-text NOTES / REMARKS column                     | Exclude rows containing Vendor/Client keywords|
| `column_exclude_values` | Scope flagged by a specific value in another column  | Exclude rows where column value matches       |
| `assume_all_in_scope`   | No scope signal detected                             | All rows in scope                             |

## Equipment Tag Detection
Two modes — set in `detected_config.equipment_mode` (auto-detected by
`format_detector.py`):

### `tag_prefix` — short-code tags
Equipment identified by tag prefix. Configured in `rules.yaml → strain_sensitive_equipment → tag_patterns`.

| Prefix          | Equipment Type                  |
|-----------------|---------------------------------|
| P- / CP-        | Centrifugal pump                |
| RP-             | Reciprocating pump              |
| C- / K-         | Compressor                      |
| ST- / GT-       | Turbine                         |
| H-              | Fired heater                    |
| B-              | Boiler                          |
| E-A / AC- / EC- | Air-cooled heat exchanger       |
| E- (not E-A)    | Shell-and-tube (non-strain-sensitive) |
| V-              | Vessel (non-strain-sensitive)   |
| T- / TW-        | Tower (check `vertical_tower` flag) |

### `keyword` — descriptive equipment names
FROM/TO cells contain full names like `"MOLTEN SULFUR TRANSFER PUMP 02AP02"`.
- Strain-sensitive detected by substring match against `keyword_patterns` (evaluated in order — "RECIPROCATING PUMP" before "PUMP")
- Non-strain-sensitive: `non_strain_sensitive_keywords` (TANK, VESSEL, DRUM, TOWER…)
- Pipe line numbers (e.g. `"50-SU-02-002-BP1-HC-SJ"`) match no keyword → not equipment

In keyword mode, if FROM/TO contains neither kind of keyword, the line falls back
to pure material-chart classification (no chart_4 override).

## Output Color Legend
| Color  | Meaning                                     |
|--------|---------------------------------------------|
| Red    | Level I — rigorous analysis required        |
| Orange | Level II — normal analysis required         |
| Green  | Level III — visual / approximate check      |
| Grey   | Excluded (SCOPE: VENDOR / CLIENT / LICENSOR)|
| Yellow | MISSING / AMBIGUOUS — Data Quality Flag set |

## Output Columns (5 appended after all original columns)
| # | Column                | Content                                                      |
|---|-----------------------|--------------------------------------------------------------|
| 1 | CRITICALITY LEVEL     | `I` / `II` / `III` (Roman), color-coded                      |
| 2 | CLASSIFICATION REASON | Full text of which rule fired                                |
| 3 | CN NUMBER             | `CN-001`, `CN-002`, … (Level I only)                         |
| 4 | CN REVIEW FLAG        | `AUTO-CONFIRMED` / `REVIEW-LARGE-CN` / `REVIEW-STANDALONE`   |
| 5 | DATA QUALITY FLAG     | `""` / `SCOPE: <LBL>` / `MISSING: <fields>` / `AMBIGUOUS: …` |

Plus two appended sheets: **SUMMARY** and **CN PROPOSALS**.

## CN (Calculation Number) Assignment

### What a CN Is
A CN is a Caesar II model boundary: a group of piping lines that must be analyzed
together because thermal forces transfer between them. The tool **proposes**
groupings; the engineer is the final decision-maker.

### Scope
- **Level 1 lines only.** Level 2/3 lines do not receive CNs.
- Every CN includes a written `grouping_reason` and a `CN_Review_Flag`.

### 6-Rule Grouping Spec (`cn_assigner.py`)
1. Same fluid service + same material + ΔT ≤ 30°C → edge in connectivity graph
2. Within each connected component, keep edges only if (pump-shared) OR (fluid+material+temp consistent) — refines groups so incompatible lines split apart
3. Sequential CN numbering: CN-001, CN-002, …
4. Large-CN soft flag: `REVIEW-LARGE-CN` when component size exceeds `max_lines_per_cn`
5. Train bucketing: if a `train` signal is present, lines are bucketed before the graph is built
6. Standalone: lines that connect to nothing get `CN-900+` with `REVIEW-STANDALONE`

### CN Number Ranges
| Range     | Type                                        |
|-----------|---------------------------------------------|
| 001–899   | All assigned CNs in sequential order        |
| 900+      | Standalone / missing-connectivity lines     |

### Engineer Review Flags
| Flag                | Meaning                                                            |
|---------------------|--------------------------------------------------------------------|
| `AUTO-CONFIRMED`    | Rules applied cleanly, no ambiguity                                |
| `REVIEW-LARGE-CN`   | Component exceeds `max_lines_per_cn` — engineer should consider splitting |
| `REVIEW-STANDALONE` | Single line, no connectivity evidence OR missing FROM/TO data      |

## Compound Temperature Parsing
Design temperature cells may contain compound values like `"60/-20"` or `"186/-20"`.
These represent MAX/MIN — the classifier splits on `/` and takes `max()`.

## Reference Runs

### Q37027 — AHF Plant (format-agnostic pipeline)
Input: `Q37027-02-A00-PE-LST-00010 (1) (1).xlsx`
Run: `python scll_tool.py --input "Q37027-02-A00-PE-LST-00010 (1) (1).xlsx" --output /tmp/scll_out.xlsx`

Detection output: header=Excel row 10, skip=row 11 (units), size=mm, equipment=keyword, scope=assume_all (no scope column detected), 19/31 columns mapped.

| Metric                | Value (2026-04-20)                                             |
|-----------------------|----------------------------------------------------------------|
| Total rows            | 907                                                            |
| Excluded              | 0 (no scope column)                                            |
| Level I               | 441                                                            |
| Level II              | 105                                                            |
| Level III             | 326                                                            |
| Missing data          | 35                                                             |
| CNs proposed          | 139 (63 auto-confirmed, 5 large, 71 standalone)                |
| Output sheets         | Coversheet + Line List (907+5 cols) + Missing no + Sheet1 + SUMMARY + CN PROPOSALS |

## File Structure
```
scll_tool.py           — CLI entry point
app.py                 — Flask web app (3-step UI)
format_detector.py     — Format-agnostic layout detection; builds detected_config
parser.py              — Reads Excel via detected_config; scope filter; mm→NPS
classifier.py          — 5-step classification engine
cn_assigner.py         — CN assignment (6-rule spec, sequential CN-001…)
output.py              — Single writer: opens original file + appends 5 cols + SUMMARY + CN PROPOSALS
rules.yaml             — Engineering data: charts, thresholds, exception flags, CN settings
mapping.yaml           — Column-name patterns (engineer-editable)
material_mapping.yaml  — Pipe-class → material group
templates/index.html   — Single-page web UI
test_linelist.xlsx     — Small synthetic test input
CLAUDE.md              — Claude Code behavioral guide
CONTEXT.md             — This file
README.md              — User instructions
requirements.txt       — Python dependencies
```

## Open Items
- Giant-CN splitting: very large HDPE vent/header networks still end up as a single
  CN (flagged `REVIEW-LARGE-CN`); engineer must split manually on P&ID
- `ped_category_3` and similar boolean flags must be pre-filled by process/piping
  engineer — the tool cannot derive them automatically
- Chart 2 (SS) boundary values for sizes 8–12" should be cross-validated against
  the original spec charts

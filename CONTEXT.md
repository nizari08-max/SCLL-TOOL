# CONTEXT.md — Engineering Knowledge Base

## Specification Reference
- Document: QW2507-00-PE-SPC-00007, Rev A, 31/10/2025
- Title: General Specification for Critical Line List
- Project: DESALINATION WAVE II EAST EXTENSION – JORF LASFAR (QW2507)
- Customer: OCP SA / JESA

## Classification Level Definitions
| Level | Name | Analysis Method |
|-------|------|-----------------|
| Level 1 | Rigorous Analysis | Computer analysis (Caesar II or equivalent) — required for critical lines |
| Level 2 | Normal Analysis | Manual calculations — guided cantilever, Tube Turns, JESA charts |
| Level 3 | Visual/Approx | Visual inspection or approximate methods by Piping Design Dept |

## Classification Logic Summary (5-Step Waterfall)

### Step 1 — FRP/GRE check
If material group == FRP → **Level 1** always.
Reason: GRE/FRP is a special case evaluated by the manufacturer. Always rigorous.

### Step 2 — Exception flags → Level 1 (19 exceptions from spec §3.2.5)
All 17 boolean flag checks. See `rules.yaml['exception_flags']` for full list.
Key exceptions:
- Relief lines with inlet pressure > 10 barg
- Cement-brick / refractory / glass-lined pipe
- Vertical tower connections >= 4"
- Jacketed piping (differential temp)
- Expansion joints, vibration, settlement, vacuum
- PED Category III >= 4"
- Underground T >= 100°C
- Category M service (ASME B31.3)
- Schedule 160 / OD/WT < 10
- Client request

### Step 3 — Strain-sensitive equipment override (Chart 3)
Triggered if FROM or TO tag matches a strain-sensitive equipment prefix.
Decision matrix (applied to the most critical equipment type found):
- Centrifugal/reciprocating pump → **Level 1** (any size, any temp)
- Any strain-sensitive, T >= 100°C → **Level 1**
- Any strain-sensitive, size <= 6", T >= 50°C → **Level 2**
- Any strain-sensitive, size > 6", T < 100°C → **Level 2**
- Any strain-sensitive, size <= 6", T < 50°C → **Level 3**

Strain-sensitive equipment types: rotating equipment (centrifugal pumps, compressors, turbines),
reciprocating pumps/compressors, heaters/boilers, air-cooled heat exchangers,
aluminum/bronze/cast iron equipment, brick/glass/refractory-lined thin-wall vessels,
shell-and-tube exchangers with shell bellows, plate-type/printed-circuit heat exchangers,
vendors specifying low nozzle loading.

### Step 4 — Material chart lookup
Chart selection:
- FROM or TO has non-strain-sensitive equipment tag → use **chart_4** (relaxed)
- CS / Low-Alloy Steel → **chart_1**
- SS / Duplex / Other alloy → **chart_2**

Each chart is a step-function of (diameter, temperature). See `rules.yaml` for thresholds.

### Step 5 — Missing data → NEEDS REVIEW
If size, design_temperature, or material is blank → flag, don't classify.

## Material Groups

### CS — Carbon Steel / Low-Alloy Steel (chart_1)
BA1, BA2, BB1, BB1U, BB2, BB3, BB4, BB5, BB6, BP1, CB1, CB2, CB3, CJ1, CP1, **CS**, FB1, FJ1

**Special notes:**
- **CS** (generic code) is used in some project files (e.g. Q83010) as a plain carbon steel designator
- **CJ1 and FJ1** are Low Alloy Steel but use CS chart (chart_1) — similar thermal properties
- **BP1** is CS but if also jacketed (`jacketed=True`), the jacketed exception fires in Step 2

### SS — Stainless Steel / Duplex / Alloy Metals (chart_2)
BD1, BD2, BD2U, BG1, BG2, BS1, BS2, BS3, CD1, CD2, CG1, CG2, CS1, ES2, GC1, GD1, GD2, KD1

### FRP — Fiber-Reinforced Plastic / GRE (always Level 1)
BK1, BK2, BK3, BK4, BV1, CK3

## Scope Filter
Two modes (set per project in `rules.yaml['scope']['mode']`):

**`include_values`** (default — generic template):
- **Keep**: Scope == `"JESA"`
- **Exclude** (grey in output): `"Vendor"`, `"Client"`, any other value

**`text_keywords`** (Q83010 and files with free-text NOTES column):
- Blank / null NOTES cell → in scope
- Cell containing `"Client"` or `"Vendor"` (case-insensitive) → EXCLUDED
- Pipe line numbers and comments not matching keywords → in scope

## Equipment Tag Convention (FROM/TO columns)

Two detection modes (set per project in `rules.yaml['strain_sensitive_equipment']['detection_mode']`):

### `prefix` mode (default — generic template)
Equipment identified by tag prefix. Configured in `tag_patterns`.

| Prefix | Equipment Type |
|--------|---------------|
| P- / CP- | Centrifugal pump |
| RP- | Reciprocating pump |
| C- / K- | Compressor |
| ST- / GT- | Turbine |
| H- | Fired heater |
| B- | Boiler |
| E-A / AC- / EC- | Air-cooled heat exchanger |
| E- (not E-A) | Shell-and-tube heat exchanger (non-strain-sensitive) |
| V- | Vessel (non-strain-sensitive) |
| T- / TW- | Tower (check vertical_tower flag separately) |

In prefix mode, any non-empty FROM/TO tag is considered an equipment connection (triggers chart_4 if not strain-sensitive).

### `keyword` mode (Q83010 — full descriptive names in FROM/TO)
FROM/TO cells contain full names like `"MOLTEN SULFUR TRANSFER PUMP 02AP02"`.
- Strain-sensitive detected by substring match against `keyword_patterns` (evaluated in order; "RECIPROCATING PUMP" before "PUMP")
- Non-strain-sensitive detected by `non_strain_sensitive_keywords` (TANK, VESSEL, DRUM, TOWER, etc.)
- Pipe line numbers (e.g. `"50-SU-02-002-BP1-HC-SJ"`) do NOT match any keyword → NOT treated as equipment connection

**Important:** In keyword mode, if FROM/TO contains neither a strain-sensitive nor non-strain-sensitive keyword, `_has_equipment_connection()` returns False → falls back to pure material chart (no chart_4 override).

## Output Color Legend
| Color | Meaning |
|-------|---------|
| Red | Level 1 — rigorous analysis required |
| Orange | Level 2 — normal analysis required |
| Green | Level 3 — visual/approx check |
| Grey | Excluded (Vendor/Client scope) |
| Yellow | Needs Review — missing data |

## File Structure
```
scll_tool.py              — CLI entry point
parser.py                 — Input reader, scope filter, mm→NPS conversion
classifier.py             — 5-step classification engine
output.py                 — Excel writer
rules.yaml                — Generic rules template (prefix mode, inches, include_values scope)
rules_q83010.yaml         — Project Q83010 rules (keyword mode, mm sizes, text_keywords scope)
material_mapping.yaml     — Pipe class → material group (shared across all projects)
test_linelist.xlsx        — 22 test rows for generic rules.yaml validation
CLAUDE.md                 — Claude Code behavioral guide
CONTEXT.md                — This file
README.md                 — User instructions
requirements.txt          — Python dependencies
```

## Projects Processed

### Q83010 — TSP S – SAFI (Molten Sulfur / Phosphoric Acid / Coating Oil Amine)
- **Input file:** `Q83010-00-00-PR-LST-00003.xlsx` (sheet "Line List", header at Excel row 10)
- **Rules file:** `rules_q83010.yaml`
- **Last run results (2026-04-19):** Total=179, Excluded=78, In-scope=101, L1=47, L2=38, L3=3, Review=13
- **NEEDS REVIEW cause:** 13 lines have size `"XX"` — process engineer must fill in actual pipe sizes
- **Key flags available:** `relief_line` (BLOWDOWN OR RELIEF LINE Y/N), `vibration` (SLUG FLOW Y/N)
- **Flags absent in this file:** vacuum, settlement, expansion_joint, jacketed, cement_lined, vertical_tower, ped_category_3, nozzle_load_limit, heavy_wall, differential_settlement, cyclic_service, category_m, client_request, underground, schedule_160 — all treated as False

## CN (Calculation Number) Assignment Logic

### What a CN Is
A CN is a Caesar II model boundary: a group of piping lines that must be analyzed together
because thermal forces transfer between them. The tool **proposes** groupings automatically.
The engineer is always the final decision maker and must change CN_Status from "PROPOSED"
to "CONFIRMED" or "REVISED".

### CN Assignment Rules — Confirmed
- **Level 1 lines only.** Level 2 and Level 3 lines do NOT receive CN assignments.
- **Level 2 does not get CNs** — Level 2 lines are analyzed by manual/simplified methods only.
- Every CN proposal includes a written reason and an engineer review flag.

### Boundary Rules (applied in strict priority order)
| Priority | Boundary Type | Action |
|----------|--------------|--------|
| 1 | Missing FROM or TO equipment tag | Hard isolate → REVIEW-MISSING-DATA → CN-9XX range |
| 2 | Expansion joint or flexible connector flag | Hard isolate → standalone CN → 300+ range |
| 3 | Centrifugal / reciprocating pump tag | Pump-first pass: each unique pump tag → dedicated CN → 001-099 range |
| 4 | Graph connected components (shared equipment tags) | BFS grouping of remaining lines |
| 4a | Different area codes within same component | Auto-split by area → REVIEW-AREA-CONFLICT |

Soft flags (no auto-split, flag only):
- Temperature delta > temp_delta_flag (default 80°C) between connected lines → REVIEW-TEMPERATURE
- Group size > max_lines_per_cn (default 12) → REVIEW-MODEL-SIZE
- Single line with no confirmed connectivity → REVIEW-MANUAL

### Network Topology Algorithm
1. **Extract Level 1 lines** from the classified DataFrame.
2. **Build line records**: parse FROM/TO equipment tags, extract area code, detect equipment types.
3. **Boundary 1** — Isolate lines with blank FROM or TO: assign to CN-9XX, flag REVIEW-MISSING-DATA.
4. **Boundary 2** — Isolate lines with expansion_joint=True: assign standalone CN in 300+ range.
5. **Boundary 3 (pump-first)** — For each unique pump tag (centrifugal or reciprocating) found in FROM/TO of remaining lines, collect all lines that directly reference that pump. Assign to a pump CN (001-099). Enforce: different pump tags → different CNs, always.
6. **Boundary 4 (graph)** — For remaining lines: build adjacency (two lines share any equipment tag → connected). Run BFS to find connected components.
7. **Area split** — For each component, split by area code. Cross-area groups → REVIEW-AREA-CONFLICT.
8. **Number CNs** using the configured ranges. Apply soft-flag checks (temp delta, model size).
9. **Write back** Proposed_CN, CN_Reason, CN_Review_Flag, CN_Status to the DataFrame.

### CN Number Ranges
| Range | Type |
|-------|------|
| 001–099 | Centrifugal / reciprocating pump CNs |
| 100–199 | Compressor / turbine CNs |
| 200–299 | Equipment-to-equipment system CNs |
| 300–399 | Area-grouped CNs (including expansion joint isolated lines) |
| 900–999 | Unassigned — missing equipment data |

### Engineer Review Flags
| Flag | Meaning |
|------|---------|
| [AUTO-CONFIRMED] | All rules applied cleanly, no ambiguity |
| [REVIEW-TEMPERATURE] | Delta T > threshold between connected lines |
| [REVIEW-MODEL-SIZE] | CN exceeds max_lines_per_cn |
| [REVIEW-MISSING-DATA] | FROM or TO equipment tag missing |
| [REVIEW-AREA-CONFLICT] | Lines from different area codes appear connected |
| [REVIEW-MANUAL] | Single line, no connectivity evidence — check P&ID |

### Output Additions (Level 1 lines in main sheet)
| Column | Content |
|--------|---------|
| Proposed_CN | CN number (e.g., QW2507-CN-001) |
| CN_Reason | Full explanation of boundary rules applied |
| CN_Review_Flag | One of the flags above |
| CN_Status | "PROPOSED" (engineer changes to CONFIRMED or REVISED) |

Additional sheets: **CN Proposals** (one row per CN) and **Dashboard** (summary statistics).

### Configuration (cn_settings in rules.yaml)
Key settings: `project_code`, `max_lines_per_cn` (default 12), `temp_delta_flag` (default 80°C),
`cn_number_format`, `area_code_extraction`, `equipment_type_prefixes`, `cn_number_ranges`,
`column_mapping.area_code` (point to a dedicated area code column or null to extract from line number).

## Project Q37027 — AHF Plant (ISBL)

### Real File Structure (Q37027-02-A00-PE-LST-00010.xlsx)
- **Sheet layout:** "Coversheet" + "Line List" + "Missing no" + "Sheet1"
- **Header row:** Excel row 10 (0-indexed: 9); row 11 = units row → skip_rows: [10]
- **Total rows:** 907 data rows (row 12 onwards)
- **Key columns (0-based index):**

| Index | Header (stripped) | Internal name |
|-------|-------------------|---------------|
| 1 | REVISION No. | — (revision, not scope) |
| 3 | LINE SIZE | — (original mm size, not used for classification) |
| 4 | Max Line Size | `size` (mm → NPS) |
| 5 | FLUID SERVICE CODE | `fluid_service_code` |
| 6 | PIPING MATERIAL CLASS | `pipe_class` |
| 7 | MOC | `material` (CS/CSG/CS300/SS/PTFECS/HDPE/FRP/RLCS) |
| 8 | AREA | `area` |
| 9 | UNIT | `unit` (used as area code for CN grouping) |
| 12 | NEW Line No. | `line_number` |
| 13 | FROM | `from_equipment` |
| 14 | TO | `to_equipment` |
| 18 | BLOWDOWN OR RELIEF LINE (Y / N) | `relief_line` |
| 19 | SLUG FLOW (Y/N) | `vibration` |
| 43 | DESIGN PRESSURE - INTERNAL | `inlet_pressure` |
| 45 | DESIGN TEMPERATURE - MAX | `design_temperature` |
| 80 | STRESS CRITICALITY (ANALYSIS LEVELS) | `stress_criticality` |
| 81 | CALCULATION NUMBER | `calculation_number` |
| 84 | NOTES/REMARKS | — (not used for scope) |

### MOC → Material Group Mapping (Q37027)
| MOC value | Material Group | Rationale |
|-----------|---------------|-----------|
| CS | CS | Plain carbon steel |
| CSG | CS | Galvanized CS — same thermal chart |
| CS300 | CS | CS-300 designation — treated as CS |
| RLCS | CS | Rubber-lined CS — substrate is CS |
| SS | SS | Stainless steel |
| PTFECS | SS | PTFE-lined CS — treat as other metallic (SS chart) |
| HDPE | FRP | Non-metallic — always Level 1 |
| FRP | FRP | Non-metallic — always Level 1 |
| Not in Piping Scope | EXCLUDED | Scope filter removes these rows |

### Scope Filter (Q37027)
- Mode: `column_exclude_values` on `material` column
- Exclude: `["Not in Piping Scope"]`
- Result: 903 in scope, 4 excluded ("Not in Piping Scope")

### CN Grouping Rules A–F Learned from Real File

**Rule A — Same fluid service + same material + same temperature zone:**
Lines with the same FLUID SERVICE CODE that are physically connected through FROM/TO AND share the same material AND operate within ±30°C belong in the same CN.
Example: CN-052 groups all FSA3+FSA2a+FSA1 HDPE lines in the Contactor/Decanter circuit at ~75°C.

**Rule B — Parallel identical trains = separate CNs:**
When the same piping system repeats across trains (detectable from repeating patterns in line numbers and equipment tags with different numerical suffixes), each train gets its own CN.
Example: CN-087 = Filter Train 1 (A11PM452–455, A11FT450A), CN-089 = Filter Train 2 (A11PM552–555, A11FT550B), CN-091 = Filter Train 3 (A11PM652–655, A11FT650B).
**Tool behavior:** pump-first pass creates separate pump CNs per unique pump tag, which implicitly separates trains. Non-pump lines from all trains may merge into one large CN via graph algorithm — engineer must manually split.

**Rule C — Pump + its suction and discharge = same CN:**
A pump and its directly connected suction/discharge lines belong in the same CN.
All Q37027 pump tags start with "A11PM" — configured in `equipment_type_prefixes.centrifugal_pump`.

**Rule D — Same header system = same CN:**
Lines all connecting to the same header belong in the same CN.
Example: CN-001 groups 3 lines all TO/FROM the same WCO header.

**Rule E — FRP/HDPE vent collection headers = own CN:**
FRP vent lines collecting into the same vent header belong in one CN.
Example: original CN-111 has 29 lines all TO VP-A11-xxx headers.
**Tool behavior:** Lines are split into sub-CNs by shared sub-header tag (CN-214, CN-215, CN-216, etc.) rather than one big CN. Flag: REVIEW-MODEL-SIZE if > max_lines_per_cn.

**Rule F — Single-line CN for ungroupable lines:**
Lines that cannot be grouped → standalone CN. Flag: [REVIEW-MANUAL].

### Compound Temperature Parsing
Design temperature cells may contain compound values like "60/-20" or "186/-20".
These represent MAX/MIN design temperatures. The classifier extracts the maximum value:
- "60/-20" → 60°C (design max)
- "186/-20" → 186°C (design max)
Fixed in `classifier.py`: `_to_float()` splits on "/" and takes `max()`.

### Validation Results (2026-04-19)
Run: `python scll_tool.py --input "Q37027-02-A00-PE-LST-00010 (1) (1).xlsx" --output Q37027_scll_output.xlsx --rules rules_q37027.yaml`

| Metric | Original file | Tool output | Match? |
|--------|--------------|-------------|--------|
| Total rows | 907 | 907 | ✓ |
| Excluded | 4 (Not in Piping Scope) | 4 | ✓ |
| Level I | 363 | 442 | Diff: +79 |
| Level II | 246 | 106 | Diff: -140 |
| Level III | 230 | 324 | Diff: +94 |
| Not assessed (-) | 67 | 31 (NEEDS REVIEW) | See notes |
| All rows have Classification Reason | N/A | ✓ 0 blank | ✓ |
| Proposed CNs | 102 | 122 | More (pump-split) |

**Level I discrepancies (28 lines):**
- Original Level I → tool Level II: PTFECS/SS lines at moderate temps connected to non-strain-sensitive equipment (engineer classified entire CN as Level I; tool classifies each line independently)
- Original Level I → tool Level III: CS lines at 60–68°C where CN contained a pump but the specific line didn't directly connect to it
- Original Level I → tool NEEDS REVIEW: 2 lines with blank design temperature

**Root cause of Level II gap (246 → 106):**
Most missing Level II lines are PTFECS lines that our tool classifies as Level II (correct per spec chart_2). The original 246 Level II included many lines that our tool also classifies as Level II, but some were reclassified to Level I (pump connections in same CN) or NEEDS REVIEW. The sum still totals 903 ✓.

**CN grouping notes:**
- CN-200: 181 lines — large connected component via shared pipe-number tags; flagged [REVIEW-TEMPERATURE]; engineer must split on P&ID
- Pump CNs: each unique pump tag gets its own CN (e.g., A11PM452 → CN-015, A11PM453 → CN-016) which means original CN-087 (11 lines, 4 pumps) is now 4 separate pump CNs + non-pump lines in CN-200
- Vent collection: split into 5–6 sub-CNs by sub-header (CN-214 through CN-218) instead of original single CN-111 (29 lines)

### New Files Added (2026-04-19)
| File | Purpose |
|------|---------|
| `rules_q37027.yaml` | Q37027-specific rules (JESA SCLL format, mm sizes, column_exclude_values scope) |
| `output_jesa.py` | JESA format Excel writer (preserves all cols, adds 4 new cols, Roman numerals) |

### Changes to Existing Files (2026-04-19)
| File | Change |
|------|--------|
| `material_mapping.yaml` | Added MOC codes: CSG, CS300, RLCS → CS; SS, PTFECS → SS; HDPE, FRP → FRP |
| `parser.py` | Added `column_exclude_values` scope mode (checks any named column for exclusion values) |
| `classifier.py` | `_to_float()` now handles compound temps like "60/-20" → takes max value |
| `cn_assigner.py` | `_parse_tags()` now splits on `\n` (newline) in addition to / & , |
| `scll_tool.py` | Added `--mode` flag (classify/cn/full); auto-detects JESA format; prevents overwriting input |

## Current Project Status
**Last updated:** 2026-04-19

### Completed
- Phase 1 MVP: full Python project created and tested end-to-end
- **Phase 3A: Q37027 Real File Support** — JESA SCLL format output writer, rules_q37027.yaml, MOC-based scope filter, compound temp parsing
- All 5 classification steps implemented and tested (22-row test suite)
- Generic `rules.yaml` template — prefix mode, inches, include_values scope
- `rules_q83010.yaml` — project-specific adaptation (keyword mode, mm→NPS, text_keywords scope)
- Multi-row Excel header support (`excel_config` in rules file)
- mm→NPS size conversion (`size_config` in rules file)
- Text-keyword scope filtering (`scope.mode: text_keywords`)
- Keyword-based strain-sensitive equipment detection (`detection_mode: keyword`)
- Non-strain-sensitive equipment keyword list (triggers chart_4 in keyword mode)
- Real project file `Q83010-00-00-PR-LST-00003.xlsx` processed successfully
- **Phase 2: CN Assignment Engine** (`cn_assigner.py`) — fully implemented and tested

### Phase 2 Test Results (2026-04-19)
Test file: `test_linelist.xlsx` (44 rows, 42 in-scope, 35 Level 1)

| CN | Lines | Flag | Scenario validated |
|----|-------|------|--------------------|
| QW2507-CN-001 | L-005, L-023 | REVIEW-TEMPERATURE | P-101 pump (suction + discharge grouped) — temp delta 215°C flagged |
| QW2507-CN-002 | L-024, L-025 | AUTO-CONFIRMED | P-102 pump (separate CN — different pump tag) |
| QW2507-CN-200 | L-026, L-027 | AUTO-CONFIRMED | Equipment-system: V-201→E-301→V-301 |
| QW2507-CN-201 | 13 lines | REVIEW-MODEL-SIZE | Hub group exceeds max_lines_per_cn=12 |
| QW2507-CN-300 | L-007 | AUTO-CONFIRMED | Expansion joint hard boundary — isolated |
| QW2507-CN-312/313/314 | L-028/029/030 | REVIEW-MANUAL | Three standalone area-03 lines, no shared equipment |
| QW2507-CN-900 | L-031 | REVIEW-MISSING-DATA | Missing To_Equipment → CN-9XX range |

**Validated: no two different pump tags share a CN.**

### Pending (Future Phases)
- Phase 3: P&ID PDF color highlighting
- Phase 4: Streamlit web UI
- Phase 5: AI-assisted review for NEEDS REVIEW lines only

### Open Questions / Edge Cases
- 13 NEEDS REVIEW lines in Q83010 — all have size "XX"; engineer must supply actual sizes
- "COATING OIL AMINE HEATER 02AE50" is classified as strain-sensitive via HEATER keyword — engineer should confirm whether this is truly strain-sensitive or a shell-and-tube exchanger
- "COATING DRUM" is detected as non-strain-sensitive via DRUM keyword → chart_4 — engineer should confirm
- Chart 2 (SS) boundary values for sizes 8–12" should be validated against original spec charts
- `ped_category_3` flag must be pre-filled by process/piping engineer; tool cannot derive it automatically
- CN assignment on Q83010 real file not yet run — FROM/TO columns present in that file but need validation against P&IDs before trusting CN proposals

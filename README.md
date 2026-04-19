# SCLL Tool — Stress Critical Line List Classifier

Automates the preparation of the Stress Critical Line List (SCLL) per JESA specification
QW2507-00-PE-SPC-00007 for project QW2507 (Desalination Wave II East Extension – Jorf Lasfar).

Given a project line list Excel file, the tool classifies each piping line as:
- **Level 1** — Rigorous computer analysis (Caesar II)
- **Level 2** — Normal manual analysis
- **Level 3** — Visual/approximate check

Output is a color-coded Excel file with a summary dashboard.

---

## Quick Start

### 1. Install Python (if not already installed)
Download Python 3.10 or later from https://python.org. During install, check "Add Python to PATH".

### 2. Install dependencies
Open a terminal in this folder and run:
```
pip install -r requirements.txt
```

### 3. Run the tool
```
python scll_tool.py --input your_linelist.xlsx --output scll_output.xlsx
```

The output file will be created in the same folder.

---

## Input File Requirements

Your line list Excel file must have these column headers (exact names matter):

| Column | Required | Description |
|--------|----------|-------------|
| Line Number | Yes | Unique line identifier |
| Size (inches) | Yes | Nominal pipe size as a number (e.g. 4, 6, 10) |
| Design Temperature (°C) | Yes | Design temperature as a number |
| Material | Yes | Pipe class code (e.g. BA1, BD2, BK1) |
| Scope | Yes | Must be "JESA" to be classified; "Vendor" or "Client" = excluded |
| FROM | No | Equipment tag at line origin (e.g. P-101, V-201) |
| TO | No | Equipment tag at line destination |
| Inlet Pressure (barg) | No | Required only for relief line check |
| Fluid/Service | No | Carried through to output unchanged |

### Boolean flag columns (optional)
If these columns exist, values of `Yes`, `TRUE`, `1`, or `X` are treated as True.
If the column is absent, it is treated as all-False.

`vacuum`, `vibration`, `settlement`, `expansion_joint`, `jacketed`, `relief_line`,
`cement_lined`, `vertical_tower`, `ped_category_3`, `nozzle_load_limit`, `heavy_wall`,
`differential_settlement`, `cyclic_service`, `category_m`, `client_request`,
`underground`, `schedule_160`

---

## Output File

The tool creates an Excel file with two sheets:

### Sheet 1 — Classified Lines
All original columns plus:
- **Level** — Level 1 / Level 2 / Level 3
- **Classification_Reason** — Text explanation of which rule fired
- **Review_Flag** — blank / NEEDS REVIEW / EXCLUDED
- **Scope_Filter_Result** — In Scope / Excluded

Row colors:
- Red = Level 1
- Orange = Level 2
- Green = Level 3
- Grey = Excluded (Vendor/Client)
- Yellow = Needs Review (missing data)

### Sheet 2 — Summary
Count of lines per level, excluded, and flagged for review.

---

## Customising Rules

All classification thresholds are in `rules.yaml`. You can edit this file with
Notepad or any text editor to:
- Update temperature/size thresholds for each chart
- Change which scope values are included/excluded
- Add new equipment tag prefixes for strain-sensitive equipment
- Update column header names to match your Excel template

**Do not edit the Python .py files** unless you are a developer.

---

## Full Command Options

```
python scll_tool.py --input FILE --output FILE [--rules FILE] [--material-map FILE]

Options:
  --input          Path to input line list .xlsx file (required)
  --output         Path to write classified output .xlsx file (required)
  --rules          Path to rules.yaml (default: rules.yaml in current folder)
  --material-map   Path to material_mapping.yaml (default: material_mapping.yaml)
```

---

## Adding New Pipe Class Codes

1. Open `material_mapping.yaml`
2. Add the code under the correct group (CS, SS, or FRP)
3. Save the file
4. Re-run the tool

---

## Troubleshooting

**"Required column(s) missing from input file"**
The column headers in your Excel file don't match what's configured in `rules.yaml`.
Open `rules.yaml`, find the `column_mappings` section, and update the right-hand values
to match your Excel headers exactly.

**"Material code 'XX' not found in material_mapping.yaml"**
The pipe class code in your line list is not in `material_mapping.yaml`.
Add it to the appropriate group and re-run.

**"Cannot write to output file"**
The output file is open in Excel. Close it and re-run the tool.

**Lines showing as "NEEDS REVIEW"**
One or more required fields (Size, Design Temperature, Material) are blank for that line.
Fill in the missing data in your line list and re-run.

---

## Running the Test

```
python scll_tool.py --input test_linelist.xlsx --output test_output.xlsx
```

Expected summary: Total=22, Excluded=2, Level 1=11, Level 2=4, Level 3=3, Needs Review=2.
Check the "Expected Level" column in `test_linelist.xlsx` against the "Level" column in the output.

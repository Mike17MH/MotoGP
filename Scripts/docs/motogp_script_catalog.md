# MotoGP Python Script Catalog

> Scanned: `c:\VS WorkSpace\MotoGP` — all subfolders included.

---

## Folder: `Scripts/` *(Current / Production)*

These are the most up-to-date scripts used for the full 2026 season workflow.

---

### 1. [motogp_parser_generic.py](file:///c:/VS%20WorkSpace/MotoGP/Scripts/motogp_parser_generic.py)

**Purpose:** Generic PDF parser — reads any MotoGP *Chronological Analysis of Performances* PDF and extracts all riders and lap-by-lap data into a structured JSON file. Automatically detects the race name, circuit, date, and all rider blocks without hardcoded rider lists.

**What it produces:** A `.json` file with metadata + one entry per rider containing laps with: lap number, lap time (text + seconds), sector times T1–T4, speed, cumulative seconds, cancelled/pit-lane/partial-timing flags.

| Parameter | Flag | Required | Default | Description |
|---|---|---|---|---|
| `pdf` | positional | ✅ | — | Path to the MotoGP Analysis PDF |
| `--output` / `-o` | optional | ❌ | `motogp_race.json` | Output JSON file path |

**Example:**
```bash
python motogp_parser_generic.py "Mandalika_Analysis.pdf" -o motogp_race_mandalika.json
```

---

### 2. [motogp_excel_builder_generic.py](file:///c:/VS%20WorkSpace/MotoGP/Scripts/motogp_excel_builder_generic.py)

**Purpose:** Builds a full Excel workbook from a JSON file produced by `motogp_parser_generic.py`. Creates one sheet per rider with lap data, a combined **Race Data** flat table, and an interactive **Race Progress** dashboard with a lap-time line chart and per-rider checkboxes.

**Tabs created:** One tab per rider + `Race Data` + `Race Progress`

| Parameter | Flag | Required | Default | Description |
|---|---|---|---|---|
| `json` | positional | ✅ | — | JSON produced by `motogp_parser_generic.py` |
| `--output` / `-o` | optional | ❌ | `motogp_race.xlsx` | Output Excel file path |

**Example:**
```bash
python motogp_excel_builder_generic.py motogp_race_mandalika.json -o Mandalika_Race.xlsx
```

---

### 3. [motogp_parser_all_riders.py](file:///c:/VS%20WorkSpace/MotoGP/Scripts/motogp_parser_all_riders.py)

**Purpose:** A **hardcoded** PDF parser specifically for the Mandalika 2026 race. Extracts all 21 riders using pre-defined page/column coordinates (RIDER_META + SEGMENTS dictionaries). Less flexible than `motogp_parser_generic.py` but very precise for that specific race's PDF layout.

**What it produces:** Same JSON structure as the generic parser.

| Parameter | Flag | Required | Default | Description |
|---|---|---|---|---|
| `pdf` | positional | ✅ | — | Path to the 6-page Mandalika Analysis PDF |
| `--output` / `-o` | optional | ❌ | `motogp_all_riders.json` | Output JSON file path |

**Example:**
```bash
python motogp_parser_all_riders.py "Mandalika_Analysis.pdf" -o motogp_all_riders.json
```

---

### 4. [motogp_excel_builder_all_riders.py](file:///c:/VS%20WorkSpace/MotoGP/Scripts/motogp_excel_builder_all_riders.py)

**Purpose:** Excel builder companion to `motogp_parser_all_riders.py`. Functionally similar to the generic builder but paired with the all-riders JSON output. Creates individual rider sheets, a flat **Race Data** table, and a **Race Progress** dashboard with a lap-time chart.

| Parameter | Flag | Required | Default | Description |
|---|---|---|---|---|
| `json` | positional | ✅ | — | JSON produced by `motogp_parser_all_riders.py` |
| `--output` / `-o` | optional | ❌ | `motogp_all_riders.xlsx` | Output Excel file path |

**Example:**
```bash
python motogp_excel_builder_all_riders.py motogp_all_riders.json -o Mandalika_AllRiders.xlsx
```

---

### 5. [motogp_probability_and_scenarios.py](file:///c:/VS%20WorkSpace/MotoGP/Scripts/motogp_probability_and_scenarios.py)

**Purpose:** Season-level analysis tool. Reads **all** race JSON files in a folder (produced by `motogp_parser_generic.py`), calculates probabilistic rider statistics (win%, podium%, pace delta), and builds a multi-sheet Excel workbook with:
- **Model Data** — raw statistical table per rider
- **Probabilistic Charts** — editable model controls + win/title probability charts
- **Championship Solver** — interactive scenario solver (editable current points, GP/Sprint finish selectors, math-alive checks)

| Parameter | Flag | Required | Default | Description |
|---|---|---|---|---|
| `input_folder` | positional | ✅ | — | Folder containing season race JSON files |
| `--output` / `-o` | optional | ❌ | `motogp_probability_scenarios.xlsx` | Output Excel file path |
| `--recent-window` | optional | ❌ | `5` | Number of recent races used for form calculation |

**Example:**
```bash
python motogp_probability_and_scenarios.py "C:\VS WorkSpace\MotoGP\Season_JSON" -o "Season_Analysis_2026\motogp_probability_scenario.xlsx" --recent-window 5
```

---

### 6. [motogp_season_bump_chart.py](file:///c:/VS%20WorkSpace/MotoGP/Scripts/motogp_season_bump_chart.py)

**Purpose:** Season-level Excel workbook with a **finishing-position bump chart** across all rounds. Reads all race JSON files in a folder, sorts them by event date, and produces:
- **Race Calendar** — list of all rounds with dates
- **Season Results** — normalized flat table of every rider's result per round
- **Position Matrix** — grid of riders vs rounds with color-coded positions (podium=green, top10=yellow, outside=red)
- **Season Progress** — interactive bump chart with per-rider checkboxes

| Parameter | Flag | Required | Default | Description |
|---|---|---|---|---|
| `input_folder` | positional | ✅ | — | Folder containing season race JSON files |
| `--output` / `-o` | optional | ❌ | `motogp_season_bump_chart.xlsx` | Output Excel file path |

**Example:**
```bash
python motogp_season_bump_chart.py "C:\VS WorkSpace\MotoGP\Season_JSON" -o "Season_Analysis_2026\motogp_season_analysis.xlsx"
```

---

## Folder: `Scripts/Scripts V1/` *(Legacy — Proof of Concept)*

Early POC scripts hardcoded for the top 3 riders of the Mandalika race only.
> ⚠️ `Scripts/Scripts_V1_Hardcored/` contains identical copies of these same 5 files.

---

### 7. [motogp_parser_poc.py](file:///c:/VS%20WorkSpace/MotoGP/Scripts/Scripts%20V1/motogp_parser_poc.py)

**Purpose:** Original POC PDF parser. Extracts lap data for only the **top 3 riders** (Bezzecchi, R. Fernandez, Martin) from a 2-page section of the Mandalika PDF using hardcoded crop coordinates. Validates that exactly 27 laps are found per rider.

| Parameter | Flag | Required | Default | Description |
|---|---|---|---|---|
| `pdf` | positional | ✅ | — | Path to Analysis.pdf |
| `--output` / `-o` | optional | ❌ | `motogp_poc.json` | Output JSON file path |

**Example:**
```bash
python motogp_parser_poc.py "Mandalika_Analysis.pdf" -o motogp_poc.json
```

---

### 8. [motogp_excel_builder_poc.py](file:///c:/VS%20WorkSpace/MotoGP/Scripts/Scripts%20V1/motogp_excel_builder_poc.py)

**Purpose:** Original POC Excel builder. Reads the 3-rider POC JSON and creates individual rider sheets, a flat Race Data table, and an interactive lap-pace chart. Y-axis is in raw seconds (no `m:ss` time formatting).

| Parameter | Flag | Required | Default | Description |
|---|---|---|---|---|
| `json` | positional | ✅ | — | JSON from `motogp_parser_poc.py` |
| `--output` / `-o` | optional | ❌ | `motogp_poc.xlsx` | Output Excel file path |

**Example:**
```bash
python motogp_excel_builder_poc.py motogp_poc.json -o motogp_poc.xlsx
```

---

### 9. [motogp_excel_builder_poc_v2.py](file:///c:/VS%20WorkSpace/MotoGP/Scripts/Scripts%20V1/motogp_excel_builder_poc_v2.py)

**Purpose:** Upgraded version of the POC builder. Adds an **Excel Time** column (column J) storing lap time as an Excel serial fraction so the chart Y-axis shows times in `m:ss.000` format instead of raw seconds. Otherwise identical scope to v1 (3-rider POC only).

| Parameter | Flag | Required | Default | Description |
|---|---|---|---|---|
| `json` | positional | ✅ | — | JSON from `motogp_parser_poc.py` |
| `--output` / `-o` | optional | ❌ | `motogp_poc.xlsx` | Output Excel file path |

**Example:**
```bash
python motogp_excel_builder_poc_v2.py motogp_poc.json -o motogp_poc_v2.xlsx
```

---

### 10. [motogp_parser_all_riders.py](file:///c:/VS%20WorkSpace/MotoGP/Scripts/Scripts%20V1/motogp_parser_all_riders.py) *(V1 copy)*

Same as **Script #3** above. This is the version stored under `Scripts V1/`.

---

### 11. [motogp_excel_builder_all_riders.py](file:///c:/VS%20WorkSpace/MotoGP/Scripts/Scripts%20V1/motogp_excel_builder_all_riders.py) *(V1 copy)*

Same as **Script #4** above. This is the version stored under `Scripts V1/`.

---

## Typical End-to-End Workflows

### Single Race (new race)
```
1. motogp_parser_generic.py  <Analysis.pdf>  →  race.json
2. motogp_excel_builder_generic.py  race.json  →  race.xlsx
```

### Full Season Analysis
```
1. Run motogp_parser_generic.py for each race  →  Season_JSON/*.json
2. motogp_probability_and_scenarios.py  Season_JSON/  →  probability_scenarios.xlsx
3. motogp_season_bump_chart.py           Season_JSON/  →  season_analysis.xlsx
```

---

## Dependencies

| Library | Used by |
|---|---|
| `pdfplumber` | All parser scripts |
| `xlsxwriter` | All Excel builder scripts |
| `colorsys` | `motogp_excel_builder_generic`, `motogp_probability_and_scenarios`, `motogp_season_bump_chart` |
| `argparse`, `json`, `re`, `pathlib` | All scripts (stdlib) |

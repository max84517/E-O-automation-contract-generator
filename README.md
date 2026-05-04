# E&O Settlement Contract Generator

Automatically fills Word templates with Excel data to batch-generate E&O (Excess & Obsolete) Settlement Agreement documents (.docx).

## Project Structure

```
project_root/
├── data/
│   └── E&O summary table.xlsx        ← Input Excel (place here manually)
├── template/
│   ├── Settlement Form_Draft.docx                 ← General template
│   └── Settlement Form_Draft_Elan & SYNA.docx     ← Elan/Synaptics template
├── output/                                        ← Generated Word contracts
├── src/
│   ├── main.py
│   ├── core/
│   │   ├── loader.py      — InfoLoader
│   │   ├── processor.py   — DataProcessor
│   │   └── writer.py      — WordWriter
│   ├── ui/
│   │   └── app.py         — CustomTkinter GUI (dark mode)
│   └── config/
│       └── paths.py       — Centralised path management
├── pyproject.toml
└── requirements.txt
```

## Installation & Usage

### Using Poetry (recommended)

```bash
poetry install
poetry run eo-generator
```

### Using pip

```bash
pip install -r requirements.txt
python src/main.py
```

## Excel Requirements

The workbook must contain two sheets:

### Summary Sheet (required columns)

| Column | Description |
|--------|-------------|
| GTK Supplier | Supplier code (merge key) |
| Platform | Platform name |
| GBU | Product category (e.g. cNB, bDT) — used for Signer matching |
| Actual Payment | Settlement amount |
| Sub-Category | Component category |

### Info Sheet (required columns)

| Column | Description |
|--------|-------------|
| Supplier | Supplier code (merge key) |
| GBU | NB / DT — differentiates Signers when a supplier has multiple rows |
| Supplier name | Full legal supplier name |
| Master Agreement | Contract number |
| Effective Date | Contract effective date |
| Signer | Signatory name |
| Signer title | Signatory title |

> **Signer matching logic**: If a supplier has multiple Info rows differentiated by GBU
> (e.g. Chicony NB / DT), the system normalises Summary.GBU to its last two characters
> (cNB → NB, bDT → DT) and matches against Info.GBU to select the correct Signer automatically.

## Word Placeholder Mapping

| Placeholder | Source |
|-------------|--------|
| `<Master Agreement>` | Info.Master Agreement |
| `<Supplier Name>` | Info.Supplier name |
| `<Effective Date>` | Info.Effective Date (formatted: Month DD, YYYY) |
| `<Sub-Category>` | Summary.Sub-Category |
| `<Platform>` | Summary.Platform |
| `<Actual Payment>` | Summary.Actual Payment (formatted: $X,XXX.XX) |
| `<Capital Money Letter>` | Summary.Actual Payment (uppercase English words) |
| `<Signer>` | Info.Signer |
| `<Title>` | Info.Signer title |

> All 9 fields are **required**. Any missing value disables generation for that row.
> The log panel shows exactly which fields are missing on failure.

## Required Fields for Contract Generation

All 9 fields below must be present for a row to be eligible for generation.
Hovering over the **⚠ Missing data** status label in the UI shows which fields are missing.

| Field | Source Sheet | Description |
|-------|-------------|-------------|
| GTK Supplier | Summary | Supplier code (merge key) |
| Platform | Summary | Platform name |
| Actual Payment | Summary | Settlement amount |
| Effective Date | Summary | Contract effective date |
| Sub-Category | Summary | Component category |
| Supplier name | Info | Full legal supplier name |
| Master Agreement | Info | Contract number |
| Signer | Info | Signatory name |
| Signer title | Info | Signatory title |

## GUI Overview

- **Three row groups**: Pending (ready to generate) → Missing Data (incomplete) → Generated (already exists)
- **Pending** rows are checked by default; **Missing Data** rows have a disabled checkbox
- **Select All / Deselect All**: batch-select eligible rows
- **Refresh**: re-scans the `output/` folder and updates row statuses
- **Generate Selected**: runs generation in a background thread with a progress bar and live log
- Failure log example: `❌ LiteOn Holmes: Missing fields — Effective Date, Master Agreement`

## Template Selection Logic

- GTK Supplier contains **Elan** or **Synaptics** (case-insensitive) → `Settlement Form_Draft_Elan & SYNA.docx`
- All other suppliers → `Settlement Form_Draft.docx`

## Output Filename Format

```
Settlement Form_<GTK Supplier> <Platform>.docx
```

## Packaging as an Executable (optional)

```bash
poetry run pyinstaller --noconfirm --onedir --windowed \
    --name "EO_Contract_Generator" \
    src/main.py
```

Output is placed in `dist/EO_Contract_Generator/`. Deploy the entire folder together with `data/`, `template/`, and `output/`.



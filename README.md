# Timing-Scripts
SAFe timing python scripts

## Prerequisites

Install the required Python packages:

```bash
pip install pandas python-pptx openpyxl
```

## Usage

```bash
python3 "SAFe Timing Sheet Script.py" --input-folder /path/to/pptx/folder
```

### Required Argument

| Argument | Description |
|---|---|
| `--input-folder` | Path to a folder containing `.pptx` files |

### Optional Arguments

| Argument | Default | Description |
|---|---|---|
| `--output` | `TimingSheet.xlsx` | Output Excel file name/path |
| `--mins-per-slide` | `2.0` | Minutes allocated per slide |
| `--day-window` | `8:30-17:30` | Day window: `8:30-17:30` or `9:00-18:00` |
| `--no-open` | *(flag)* | Skip auto-opening the file on macOS |
| `--verbose` | *(flag)* | Enable verbose output |

### Examples

Basic usage:
```bash
python3 "SAFe Timing Sheet Script.py" --input-folder ./decks
```

With a 9:00–18:00 day window and custom output file:
```bash
python3 "SAFe Timing Sheet Script.py" \
  --input-folder ./decks \
  --output MyTiming.xlsx \
  --day-window 9:00-18:00
```

With a custom slide duration:
```bash
python3 "SAFe Timing Sheet Script.py" \
  --input-folder ./decks \
  --mins-per-slide 2.5
```

## Output

The script produces an Excel file (`TimingSheet.xlsx` by default) with:
- A **Timing** sheet listing all lessons, breaks, lunch, and close rows with Excel formulas for Start/End times and durations
- A **Settings** sheet storing `MinsPerSlide`, `StartOfDay`, and `EndOfDay` values used by the formulas

On macOS the file is automatically opened after generation (use `--no-open` to suppress this).

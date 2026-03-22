# Timing-Scripts
SAFe timing python scripts

## Prerequisites

Install the required Python packages:

```bash
pip install python-pptx openpyxl
```

## Usage

```bash
python3 "SAFe Timing Sheet Script.py" --input-folder /path/to/pptx/folder
```

When the script runs it will prompt you for configuration before generating the Excel file. Press **Enter** to accept defaults, or type your choices. Pass `--no-config` to skip the prompts entirely and rely on CLI flag defaults.

### Interactive Configuration

Each run shows a two-step configuration prompt:

**1 – Day Timing**
```
── Day Timing ──
  1. 9:00 – 18:00  (lunch 13:00)  [default]
  2. 8:30 – 18:00  (lunch 12:30)
  3. Custom start time
Choice [1]:
```
Choosing option 3 asks for a start time (e.g. `08:00`); lunch and hourly breaks are derived automatically.

**2 – Lesson Balancing**
```
── Lesson Balancing ──
  Detected modules: 1, 2, 3, 4, 5, 6, 7

  Distribution mode:
  1. Weighted – natural overflow by content  [default]
  2. Even     – split modules as evenly as possible across N days
  3. Custom   – specify which module starts each day
Choice [1]:
```
- **Weighted**: lessons overflow to the next day when the day is full (original behaviour)
- **Even**: enter a number of days; modules are split as evenly as possible
- **Custom**: enter `MODULE:DAY` pairs, e.g. `7:4` forces module 7 to start on day 4

### Required Argument

| Argument | Description |
|---|---|
| `--input-folder` | Path to a folder containing `.pptx` files |

### Optional Arguments

| Argument | Default | Description |
|---|---|---|
| `--output` | `TimingSheet.xlsx` | Output Excel file name/path |
| `--mins-per-slide` | `2.0` | Minutes allocated per slide |
| `--day-window` | `9:00-18:00` | Pre-set day window (`8:30-18:00` or `9:00-18:00`) — only used with `--no-config` |
| `--no-config` | *(flag)* | Skip interactive prompts and use CLI flag defaults |
| `--no-open` | *(flag)* | Skip auto-opening the file on macOS |
| `--verbose` | *(flag)* | Enable verbose output |

### Examples

Basic usage (interactive config):
```bash
python3 "SAFe Timing Sheet Script.py" --input-folder ./decks
```

4-day class, modules 1–6 spread across days 1–3, module 7 pinned to day 4:
```
Choice [1]: 3   ← Custom balancing
  > 7:4
```

Non-interactive (CI / scripting):
```bash
python3 "SAFe Timing Sheet Script.py" \
  --input-folder ./decks \
  --output MyTiming.xlsx \
  --day-window 9:00-18:00 \
  --no-config
```

## Output

The script produces an Excel file (`TimingSheet.xlsx` by default) with:
- A **Timing** sheet listing all lessons, breaks, lunch, and close rows with Excel formulas for Start/End times and durations
- A **Settings** sheet storing `MinsPerSlide`, `StartOfDay`, and `EndOfDay` values used by the formulas

On macOS the file is automatically opened after generation (use `--no-open` to suppress this).

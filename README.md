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

When the script runs it will prompt you for configuration before generating the Excel file. Press **Enter** to accept defaults. Pass `--no-config` to skip all prompts and use CLI flag defaults.

### Interactive Configuration

Each run shows a three-step configuration prompt:

**1 – Training Mode**
```
── Training Mode ──
  1. Online      (10-min breaks every hour)           [default]
  2. In-Person   (15-min morning + afternoon breaks)
Choice [1]:
```
- **Online**: 10-minute break every hour throughout the day
- **In-Person**: two 15-minute breaks — morning at start + 1h45m, afternoon at start + 6h45m (e.g. 10:45 & 15:45 for a 9:00 start)

**2 – Day Timing**
```
── Day Timing ──
  1. 9:00 – 18:00  (lunch 13:00)  [default]
  2. 8:30 – 18:00  (lunch 12:30)
  3. Custom start time
Choice [1]:
```
Choosing option 3 asks for a start time (e.g. `08:00`); lunch and hourly breaks are derived automatically.

**3 – Lesson Balancing**

First, anchor any modules that must start on a specific day:
```
  Anchor any module to a specific day? (e.g. 7:4 pins module 7 to day 4)
  Press Enter with no input when done.
  > 7:4
  >
```

Then enter the total number of days (defaults to the highest anchored day):
```
  Number of days [4]: 4
```

Finally, choose how the remaining modules are distributed:
```
  Distribute remaining modules (1, 2, 3, 4, 5, 6) across days 1–3:
    1. Weighted – natural content overflow  [default]
    2. Balanced – equalise total time per day
  Choice [1]:
```

- **Weighted**: lessons fill each day by content length until the day ends naturally
- **Balanced**: finds cut points between modules that equalise total minutes per day; earlier days are slightly favoured when days are unequal; complete modules are never split

The summary confirms what was configured:
```
── Summary ──
  Start time : 09:00
  Day window : 09:00 – 18:00  (lunch 13:00)
  Modules 1, 2, 3, 4, 5, 6 → days 1–3  (weighted)
  Module 7 → day 4  (anchored)
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
| `--day-window` | `9:00-18:00` | Pre-set day window (`8:30-18:00` or `9:00-18:00`) — only used with `--no-config` |
| `--mode` | `online` | Break pattern: `online` (10-min/hour) or `in-person` (15-min morning+afternoon) — only used with `--no-config` |
| `--no-config` | *(flag)* | Skip interactive prompts and use CLI flag defaults |
| `--no-open` | *(flag)* | Skip auto-opening the file on macOS |
| `--verbose` | *(flag)* | Enable verbose output |

### Examples

Basic usage — 4-day class, modules 1–6 weighted across days 1–3, module 7 on day 4:
```
python3 "SAFe Timing Sheet Script.py" --input-folder ./decks

  > 7:4          ← anchor module 7 to day 4
  Number of days [4]: ← press Enter to accept 4
  Choice [1]:    ← press Enter for Weighted
```

Non-interactive (scripting / CI):
```bash
python3 "SAFe Timing Sheet Script.py" \
  --input-folder ./decks \
  --output MyTiming.xlsx \
  --day-window 9:00-18:00 \
  --no-config

# In-person mode:
python3 "SAFe Timing Sheet Script.py" \
  --input-folder ./decks \
  --output MyTiming.xlsx \
  --day-window 9:00-18:00 \
  --mode in-person \
  --no-config
```

## Output

The script produces an Excel file (`TimingSheet.xlsx` by default) with:
- A **Timing** sheet listing all lessons, breaks, lunch, and close rows with Excel formulas for Start/End times and durations
- A **Settings** sheet storing `MinsPerSlide`, `StartOfDay`, and `EndOfDay` values used by the formulas

On macOS the file is automatically opened after generation (use `--no-open` to suppress this).

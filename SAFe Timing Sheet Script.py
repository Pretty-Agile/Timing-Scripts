#!/usr/bin/env python3
"""
Pretty Agile — SAFe Timing v4.5
- Hard daily cut-off: 9:00→18:00 (default); 8:30→17:30 (optional)
- Visible 'Start of Day' row (0 mins, grey) resets time each day
- Do NOT skip breaks/lunch: if a Break/Lunch would overflow today, end day and place it first thing next day
- Final day: if time remains, add 'Close – Photo, Feedback & Exam' to fill remaining time (formula)
- 12-hour h:mm AM/PM formatting; grey shading for Break/Lunch/Start of Day
- Excel formulas for Minutes/Total/Start/End; no Day/Deck columns
- Fixed breaks (10m): 8:30→ 09:30,10:30,11:30,14:30,15:30,16:30; 9:00→ 10:00,11:00,12:00,15:00,16:00,17:00
- Lunch target 12:30 (8:30) or 13:00 (9:00), earliest = target−15m; before/after closest boundary; no splitting lessons

NEW:
- Day 1: 30-minute "Intros" row immediately after "Start of Day"
- Days 2+: no Intros row
- Intros consume the day window (less time for lessons/Close)
"""
import argparse, re, os, sys, platform, time
from pathlib import Path
from typing import List, Tuple, Optional, Dict, Any
from concurrent.futures import ProcessPoolExecutor, as_completed
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from datetime import timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill

# ---------- regex ----------
ACTIVITY_PREFIX = re.compile(r'^\s*(activity|discussion|video|action\s*plan)\b', re.I)
VIDEO_PREFIX    = re.compile(r'^\s*video\b', re.I)
PI_PLANNING_DECK = re.compile(r'leading\s+safe|safe\s+scrum\s+master|implementing\s+safe', re.I)
PI_PLANNING_SLIDE = re.compile(r'pi\s+planning', re.I)
SUBHEADER_RE = re.compile(r'^\s*\d+\.\d+\b(?!\.)')
DUR_PATTERNS = [
    (re.compile(r'(\d{1,3})\s*-\s*(\d{1,3})\s*m(?:in(?:s|utes)?)?\b', re.I),
     lambda m: max(int(m.group(1)), int(m.group(2)))),
    (re.compile(r'(\d{1,2})\s*h(?:ours?)?\s*(\d{1,3})\s*m(?:in(?:s|utes)?)?\b', re.I),
     lambda m: int(m.group(1)) * 60 + int(m.group(2))),
    (re.compile(r'(\d{1,2})\s*h(?:ours?)?\b', re.I),
     lambda m: int(m.group(1)) * 60),
    (re.compile(r'(\d{1,3})\s*m(?:in(?:s|utes)?)?\b', re.I),
     lambda m: int(m.group(1))),
    (re.compile(r'(?:^|\D)(\d{1,3})(?!\s*[\.\d])(?:\D|$)'),
     lambda m: int(m.group(1))),
]

# ---------- utils ----------
def clean_one_line(s: str) -> str:
    return " ".join((s or "").split())

def strip_leading_two_digits(name: str) -> str:
    return re.sub(r'^\s*\d{2}\s+', '', name).strip()

def deck_label_from_filename(fname: str) -> str:
    base = Path(fname).stem
    base = re.sub(r'\s*\(Converted\)\s*$', '', base, flags=re.I)
    base = re.sub(r'\s*\(\d+(?:\.\d+)*\)\s*$', '', base)
    return base.strip()

def slide_title(slide) -> str:
    for shp in slide.shapes:
        try:
            if shp.is_placeholder and shp.placeholder_format.type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
                if shp.has_text_frame:
                    t = clean_one_line(shp.text or "")
                    if t:
                        return t
        except Exception:
            pass
    for shp in slide.shapes:
        if getattr(shp, "has_text_frame", False):
            t = clean_one_line(shp.text or "")
            if t:
                return t
    return ""

def slide_all_text(slide) -> str:
    parts = []
    for shp in slide.shapes:
        if getattr(shp, "has_text_frame", False):
            if shp.text:
                parts.append(shp.text)
    try:
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            parts.append(slide.notes_slide.notes_text_frame.text or "")
    except Exception:
        pass
    return "\n".join(p.strip() for p in parts if p and p.strip())

def is_subheader_title(title: str) -> bool:
    return bool(SUBHEADER_RE.match(title or ""))

def normalized_sublesson_name(slide, header_title: str, precomputed_body: str = None) -> str:
    title = header_title or ""
    code_m = re.match(r'^\s*(\d+\.\d+)\b', title)
    code = code_m.group(1) if code_m else ""
    m = re.match(r'^\s*\d+\.\d+\s*[:\-–]?\s*(.+)$', title)
    if m and m.group(1).strip():
        return f"{code} {m.group(1).strip()}"
    body = precomputed_body if precomputed_body is not None else slide_all_text(slide)
    for line in (l.strip() for l in body.splitlines() if l.strip()):
        m2 = re.match(r'^\s*\d+\.\d+\s*[:\-–]?\s*(.+)$', line)
        if m2 and m2.group(1).strip():
            return f"{code} {m2.group(1).strip()}"
    lines = [l.strip() for l in body.splitlines() if l.strip()]
    longest = max(lines, key=len) if lines else "Sub-lesson"
    return f"{code} {longest}"

def is_activity_slide(title: str) -> bool:
    return bool(ACTIVITY_PREFIX.match(title or ""))

def is_video_slide(title: str) -> bool:
    return bool(VIDEO_PREFIX.match(title or ""))

def extract_minutes(text: str) -> Optional[int]:
    s = (text or "").replace("\u2013", "-").replace("\u2014", "-")
    for pat, func in DUR_PATTERNS:
        m = pat.search(s)
        if m:
            try:
                return max(0, func(m))
            except Exception:
                continue
    return None

# ---------- parse deck ----------
def parse_deck(path: Path, deck_label: str) -> List[Tuple[str, int, int, str]]:
    prs = Presentation(str(path))
    slides = list(prs.slides)
    titles = [slide_title(s) for s in slides]
    texts = [slide_all_text(s) for s in slides]
    combined = [(titles[i] + "\n" + texts[i]).strip() for i in range(len(slides))]

    header_idxs, header_names = [], {}
    for i, t in enumerate(titles):
        if is_subheader_title(t):
            header_idxs.append(i)
            header_names[i] = normalized_sublesson_name(slides[i], t, texts[i])

    deck_clean_no_lead = strip_leading_two_digits(deck_label)
    intro_label = "1.0 " + deck_clean_no_lead
    if header_idxs:
        first_name = header_names[header_idxs[0]] or titles[header_idxs[0]]
        m = re.match(r'^\s*(\d+)\.\d+\b', first_name or "")
        if m:
            intro_label = f"{m.group(1)}.0 {deck_clean_no_lead}"
    else:
        m = re.match(r'^\s*(\d+)', deck_label)
        if m:
            intro_label = f"{int(m.group(1))}.0 {deck_clean_no_lead}"

    is_pi_deck = bool(PI_PLANNING_DECK.search(deck_label))

    def sum_activity(start: int, end: int) -> int:
        total = 0
        for i in range(start, end):
            if is_activity_slide(titles[i]):
                mins = extract_minutes(combined[i])
                if mins is not None:
                    if is_pi_deck and mins == 50 and PI_PLANNING_SLIDE.search(combined[i]):
                        mins = 30
                    total += mins
        return total

    def count_video(start: int, end: int) -> int:
        return sum(1 for i in range(start, end) if is_video_slide(titles[i]))

    def adjusted_row(name, start, end):
        """Return (name, slide_count, activity_mins) with video slides contributing 1 min each."""
        vid = count_video(start, end)
        count = max(0, end - start) - vid
        act = sum_activity(start, end) + vid  # 1 min per video slide replaces mins_per_slide
        return [name, count, act, "lesson"]

    rows: List[Tuple[str, int, int, str]] = []
    if not header_idxs:
        rows.append(adjusted_row(intro_label, 0, len(slides)))
        return rows

    first = header_idxs[0]
    if first > 0:
        rows.append(adjusted_row(intro_label, 0, first))

    for k, start in enumerate(header_idxs):
        end = header_idxs[k + 1] if k + 1 < len(header_idxs) else len(slides)
        name = header_names.get(start) or f"Sub-lesson {k+1}"
        rows.append(adjusted_row(name, start, end))

    return rows

# ---------- time helpers ----------
def parse_hhmm(s: str) -> timedelta:
    h, m = s.split(":")
    return timedelta(hours=int(h), minutes=int(m))

def fixed_break_targets(day_window: str) -> List[timedelta]:
    if day_window == "8:30-17:30":
        return [
            timedelta(hours=9, minutes=30),
            timedelta(hours=10, minutes=30),
            timedelta(hours=11, minutes=30),
            timedelta(hours=14, minutes=30),
            timedelta(hours=15, minutes=30),
            timedelta(hours=16, minutes=30),
        ]
    else:
        return [
            timedelta(hours=10, minutes=0),
            timedelta(hours=11, minutes=0),
            timedelta(hours=12, minutes=0),
            timedelta(hours=15, minutes=0),
            timedelta(hours=16, minutes=0),
            timedelta(hours=17, minutes=0),
        ]

# ---------- sequence with hard cut-off ----------
def build_sequence(
    rows_by_deck: List[Dict[str, Any]],
    day_start: str,
    day_end: str,
    lunch_time: str,
    day_window: str,
    mins_per_slide: float,
) -> List[Dict[str, Any]]:
    """
    Returns ordered rows (lessons + Start of Day + Intros + breaks + lunch + final close),
    respecting daily cut-offs and 'do not skip' for breaks/lunch.
    """
    start_td = parse_hhmm(day_start)
    end_td = parse_hhmm(day_end)
    lunch_td = parse_hhmm(lunch_time)
    lunch_earliest = lunch_td - timedelta(minutes=15)
    targets = fixed_break_targets(day_window)

    out: List[Dict[str, Any]] = []
    day = 1
    t = start_td
    lunch_done_today = False
    next_break_idx = 0
    first_lesson_scheduled_today = False

    def start_day_row():
        nonlocal t, lunch_done_today, next_break_idx, first_lesson_scheduled_today
        t = start_td
        lunch_done_today = False
        next_break_idx = 0
        first_lesson_scheduled_today = False

        # Visible start-of-day marker (0 mins, grey in Excel)
        out.append(
            {
                "Sub-lesson": "Start of Day",
                "Slides": 0,
                "Activity Minutes": 0,
                "kind": "sod",
            }
        )

        # Intros row: 30 mins on Day 1 only
        if day == 1:
            out.append(
                {
                    "Sub-lesson": "Intros",
                    "Slides": 0,
                    "Activity Minutes": 30,
                    "kind": "intros",
                }
            )
            t = t + timedelta(minutes=30)

    def end_day_and_start_next():
        nonlocal day
        day += 1
        start_day_row()

    # Begin Day 1
    start_day_row()

    # Helper to append Break/Lunch respecting overflow (move to next day if needed)
    def append_block(kind: str, minutes: int):
        nonlocal t
        # Don't place a break immediately after another break
        if kind == "break" and out and out[-1]["kind"] == "break":
            return
        if t + timedelta(minutes=minutes) > end_td:
            end_day_and_start_next()
        out.append(
            {
                "Sub-lesson": "Break" if kind == "break" else "Lunch",
                "Slides": 0,
                "Activity Minutes": minutes,
                "kind": kind,
            }
        )
        t = t + timedelta(minutes=minutes)

    i = 0
    while i < len(rows_by_deck):
        row = rows_by_deck[i]
        slides = int(row["Slides"])
        minutes = int(round(slides * mins_per_slide))
        total = minutes + int(row["Activity Minutes"])

        # LUNCH BEFORE block if at/after earliest and closer before
        if not lunch_done_today and t >= lunch_earliest:
            before_delta = abs((t - lunch_td).total_seconds())
            after_delta = abs(((t + timedelta(minutes=total)) - lunch_td).total_seconds())
            if before_delta <= after_delta:
                append_block("lunch", 45)
                lunch_done_today = True
                # skip the first post-lunch break target
                if day_window == "8:30-17:30":
                    skip_target = timedelta(hours=13, minutes=30)
                else:
                    skip_target = timedelta(hours=14, minutes=0)
                if next_break_idx < len(targets) and targets[next_break_idx] == skip_target:
                    next_break_idx += 1

        # BREAK BEFORE if target falls inside this lesson window and before boundary is closer
        if next_break_idx < len(targets):
            target = targets[next_break_idx]
            if t > target:
                # Catch-up break – only insert if lesson still fits in today after the break
                if first_lesson_scheduled_today:
                    if t + timedelta(minutes=10 + total) <= end_td:
                        append_block("break", 10)
                    next_break_idx += 1
            elif t + timedelta(minutes=total) > target:
                dist_before = abs((t - target).total_seconds())
                dist_after = abs(((t + timedelta(minutes=total)) - target).total_seconds())
                if dist_before <= dist_after and first_lesson_scheduled_today:
                    if t + timedelta(minutes=10 + total) <= end_td:
                        append_block("break", 10)
                    next_break_idx += 1

        # If the lesson itself won't fit, push to next day
        if t + timedelta(minutes=total) > end_td:
            end_day_and_start_next()
            # Lunch check will happen naturally on the new day loop

        # Schedule lesson
        out.append(
            {
                "Sub-lesson": row["Sub-lesson"],
                "Slides": slides,
                "Activity Minutes": int(row["Activity Minutes"]),
                "kind": "lesson",
            }
        )
        t = t + timedelta(minutes=total)
        first_lesson_scheduled_today = True

        # LUNCH AFTER if closer and allowed
        if not lunch_done_today and t >= lunch_earliest:
            before_delta = abs(((t - timedelta(minutes=total)) - lunch_td).total_seconds())
            after_delta = abs((t - lunch_td).total_seconds())
            if after_delta < before_delta:
                append_block("lunch", 45)
                lunch_done_today = True
                # skip immediate post-lunch target
                if day_window == "8:30-17:30":
                    skip_target = timedelta(hours=13, minutes=30)
                else:
                    skip_target = timedelta(hours=14, minutes=0)
                if next_break_idx < len(targets) and targets[next_break_idx] == skip_target:
                    next_break_idx += 1

        # BREAK AFTER if target crossed and after boundary closer
        if next_break_idx < len(targets):
            target = targets[next_break_idx]
            if t > target:
                append_block("break", 10)
                next_break_idx += 1

        i += 1

    # After last lesson sequence, add final Close row if time remains in the current day
    # (Excel will compute the minutes via formula, see writer below)
    out.append(
        {
            "Sub-lesson": "Close – Photo, Feedback & Exam",
            "Slides": 0,
            "Activity Minutes": 0,
            "kind": "close",
        }
    )

    return out

# ---------- parallel helper (top-level so it's picklable) ----------
def _parse_deck_task(args):
    path_str, deck_label = args
    rows = parse_deck(Path(path_str), deck_label)
    return path_str, deck_label, rows

# ---------- runner ----------
def main():
    ap = argparse.ArgumentParser(description="SAFe timing → Excel with hard cut-offs and start-of-day rows.")
    ap.add_argument("--input-folder", required=True)
    ap.add_argument("--output", default="TimingSheet.xlsx")
    ap.add_argument("--mins-per-slide", type=float, default=2.0)
    ap.add_argument("--day-window", choices=["8:30-17:30", "9:00-18:00"], default="9:00-18:00")
    ap.add_argument("--no-open", action="store_true")
    ap.add_argument("--verbose", action="store_true")
    args = ap.parse_args()

    if args.day_window == "8:30-17:30":
        day_start, day_end, lunch_time = "08:30", "17:30", "12:30"
    else:
        day_start, day_end, lunch_time = "09:00", "18:00", "13:00"

    in_dir = Path(args.input_folder).expanduser().resolve()
    if not in_dir.exists():
        print(f"✗ Folder not found: {in_dir}")
        sys.exit(1)

    pptx_files = sorted([p for p in in_dir.glob("*.pptx") if p.is_file()])
    if not pptx_files:
        print(f"✗ No .pptx files found in {in_dir}")
        sys.exit(1)

    # Parse decks (parallel when multiple files)
    tasks = [(str(f), deck_label_from_filename(f.name)) for f in pptx_files]
    deck_results: Dict[str, Tuple] = {}

    if len(pptx_files) > 1:
        with ProcessPoolExecutor() as executor:
            futures = {executor.submit(_parse_deck_task, t): t for t in tasks}
            for future in as_completed(futures):
                try:
                    path_str, deck_label, rows = future.result()
                    deck_results[deck_label] = (rows, Path(path_str).name)
                    print(f"✓ {Path(path_str).name}  → {len(rows)} rows", flush=True)
                except Exception as e:
                    _, deck_label = futures[future]
                    print(f"✗ {deck_label}  FAILED: {e}", flush=True)
    else:
        for path_str, deck_label in tasks:
            try:
                _, _, rows = _parse_deck_task((path_str, deck_label))
                deck_results[deck_label] = (rows, Path(path_str).name)
                print(f"✓ {Path(path_str).name}  → {len(rows)} rows", flush=True)
            except Exception as e:
                print(f"✗ {Path(path_str).name}  FAILED: {e}", flush=True)

    # Flatten results preserving original file order
    parsed_rows = []
    for f in pptx_files:
        deck_label = deck_label_from_filename(f.name)
        if deck_label not in deck_results:
            continue
        rows, _ = deck_results[deck_label]
        for name, slides, act, kind in rows:
            parsed_rows.append(
                {
                    "Deck": deck_label,
                    "Sub-lesson": name,
                    "Slides": int(slides or 0),
                    "Activity Minutes": int(act or 0),
                    "kind": kind,
                }
            )

    if not parsed_rows:
        print("✗ No data parsed.")
        sys.exit(1)

    # Sort by deck then sublesson code
    def subcode(s: str) -> str:
        m = re.match(r"^\s*(\d+\.\d+)", s or "")
        return m.group(1) if m else ""

    parsed_rows.sort(key=lambda r: (r["Deck"], subcode(r["Sub-lesson"]), r["Sub-lesson"]))

    # Build sequence with hard cut-offs (now includes Intros rows)
    sequence = build_sequence(
        parsed_rows,
        day_start,
        day_end,
        lunch_time,
        args.day_window,
        mins_per_slide=args.mins_per_slide,
    )

    out_path = Path(args.output).expanduser().resolve()
    tmp_path = out_path.with_suffix(".tmp.xlsx")

    # Build workbook directly with openpyxl (no pandas needed)
    wb = Workbook()
    ws = wb.active
    ws.title = "Timing"
    ws.append(["Sub-lesson", "Slides", "Minutes", "Activity Minutes", "Total Minutes", "Start", "End"])
    for r in sequence:
        ws.append([r["Sub-lesson"], r["Slides"], None, r["Activity Minutes"], None, None, None])

    # Settings sheet
    settings = wb["Settings"] if "Settings" in wb.sheetnames else wb.create_sheet("Settings")
    settings["A1"] = "MinsPerSlide"
    settings["B1"] = float(args.mins_per_slide)
    settings["A2"] = "StartOfDay"
    settings["B2"] = day_start
    settings["A3"] = "EndOfDay"
    settings["B3"] = day_end

    # Apply formulas & formats
    first_data_row = 2
    last_row = ws.max_row
    time_fmt = "h:mm AM/PM"
    grey = PatternFill(fill_type="solid", start_color="EEEEEE", end_color="EEEEEE")

    # Helper: detect row label
    def label(r):
        return (ws[f"A{r}"].value or "").strip()

    for r in range(first_data_row, last_row + 1):
        lbl = label(r)

        # Minutes = Slides * Settings!$B$1 (for all ordinary rows)
        ws[f"C{r}"] = f'=IF(B{r}="","",B{r}*Settings!$B$1)'

        # Total Minutes:
        if lbl == "Close – Photo, Feedback & Exam":
            # Fill remaining time to EndOfDay (never negative)
            ws[f"E{r}"] = f"=MAX(0,(TIMEVALUE(Settings!$B$3)-F{r})*1440)"
        else:
            ws[f"E{r}"] = f'=IF(C{r}="","",C{r}+D{r})'

        # Start / End formulas:
        if r == first_data_row:
            # First row gets StartOfDay
            ws[f"F{r}"] = "=TIMEVALUE(Settings!$B$2)"
        else:
            # If this is a 'Start of Day' marker, reset Start to StartOfDay; otherwise chain from previous End
            if lbl == "Start of Day":
                ws[f"F{r}"] = "=TIMEVALUE(Settings!$B$2)"
            else:
                ws[f"F{r}"] = f"=G{r-1}"

        ws[f"G{r}"] = f"=F{r}+E{r}/1440"

        # Display format for Start/End
        ws[f"F{r}"].number_format = time_fmt
        ws[f"G{r}"].number_format = time_fmt

        # Grey fill for Break/Lunch/Start of Day (Intros stays normal)
        if lbl in ("Break", "Lunch", "Start of Day"):
            for col in ("A", "B", "C", "D", "E", "F", "G"):
                ws[f"{col}{r}"].fill = grey

    # Column widths
    for col, w in {"A": 60, "B": 8, "C": 10, "D": 16, "E": 14, "F": 12, "G": 12}.items():
        ws.column_dimensions[col].width = w

    wb.save(tmp_path)

    # Replace original file
    try:
        if out_path.exists():
            out_path.unlink()
        Path(tmp_path).rename(out_path)
        final_path = out_path
    except Exception:
        ts_path = out_path.with_name(out_path.stem + f".{int(time.time())}.xlsx")
        Path(tmp_path).rename(ts_path)
        final_path = ts_path
        print(f"⚠️ Could not replace {out_path.name} (maybe open). Kept: {final_path}")

    print(f"\n✅ Wrote {last_row-1} rows")
    print(f"📄 File: {final_path}")

    if (not args.no_open) and platform.system() == "Darwin":
        os.system(f'open "{final_path}"')

if __name__ == "__main__":
    main()

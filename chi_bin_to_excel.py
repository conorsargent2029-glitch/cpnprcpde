"""
chi_bin_to_excel.py

A separate, standalone program from chi_swv_to_excel.py -- this one does NOT
watch a folder or run continuously. It's a one-off reduction tool: point it
at a folder of CHI .bin saves that never got a matching .txt export (e.g.
CHI's auto-text-save setting was off for an unattended overnight run), and
it writes everything straight into an Excel sheet in a single pass, then
exits.

Why this needs to exist at all
-------------------------------
Normally CHI auto-saves a readable .txt file for every run, and
chi_swv_to_excel.py just parses those. But .txt export is a separate
software setting from the .bin save itself -- if it was off, you're left
with only the named .bin files, which look empty at a glance (no plain text
in them). They're not actually empty: CHI's .bin format stores the full raw
current curve, just packed as binary floats instead of a text table. This
script reads that binary layout directly (reverse-engineered from a matched
.bin/.txt pair -- see the BIN_* constants and parse_swv_bin_file below for
exactly how).

The one thing CHI does NOT store in the .bin is its own already-computed
"ip" peak value (the number that would show up in the .txt's "Results:"
section) -- confirmed by exhaustively searching a known .bin for its paired
.txt's exact reported values and finding nothing, in any float format. CHI
must calculate that only when it exports to text. So every ip value this
script produces is an estimate, derived from the raw curve by
estimate_peak_from_curve() (a straight-line baseline fit through the edges
of the curve, then the point of largest deviation from that line -- the
standard by-hand way of reading peak height off a voltammogram). Cross-
checked against a known CHI-reported result, this landed within 0.3-1.9% on
ip and exact on peak potential -- good, but not identical to CHI's own
algorithm, so treat these numbers as approximate, not lab-final.

How labeling works
-------------------
Files are paired into cycles by the number CHI appends to the filename
(e.g. "10hz_blank_5.bin" pairs with "60hz_blank_5.bin" -- both are cycle
5), NOT by file timestamp. Timestamps turned out to be unreliable for
this: if a folder gets reorganized/moved after the fact (copying files on
Windows resets creation time) or synced by OneDrive, the save-time order
timestamps imply can end up completely scrambled, badly mispairing rows.
CHI doesn't suffix the very first save of a given name, only repeats get
_2, _3, ... -- so a filename with no trailing number is treated as cycle 1.
Cycles are then written out in ascending cycle-number order: the first is
labeled FIRST_LABEL ("initial reading" by default), every one after in
STEP_MINUTES increments (20 mins, 40 mins, ...) -- one row per cycle, no
filler rows for a cycle where one frequency's file is missing entirely.

DAILY USE
---------
Double-click run_bin_to_excel.bat (in the same folder as this script), or
run it from a terminal:

    python chi_bin_to_excel.py

It'll ask for the folder of .bin files, the Excel file to write to, and the
minutes between readings (default 20), then processes everything in one
pass and exits -- there's no "leave the window open," nothing to watch.
"""

import os
import re
import struct
import sys
import time

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

# ----------------------- CONFIG -----------------------
EXCEL_DIR = r"C:\Users\gao22\Documents"    # folder the workbook is saved into
DEFAULT_EXCEL_NAME = "Overnight_SWV_Results.xlsx"
FREQUENCIES = [10, 60]     # the SWV frequencies to look for, in column order left to right
DEFAULT_STEP_MINUTES = 20  # minutes between readings
DEFAULT_FIRST_LABEL = "initial reading"
SETTINGS_DIR = r"C:\Users\gao22\Documents"  # where "last used" choices are remembered
# --------------------------------------------------------

NUM_FMT = "0.00E+00"
ESTIMATED_FONT = Font(italic=True)

EXCEL_PATH = None  # set at startup by prompt_for_excel_path()


def prompt_for_excel_path():
    global EXCEL_PATH
    last_used = _load_last_used("chibin_last_excel_name")
    default_name = last_used or DEFAULT_EXCEL_NAME
    typed = input(f"Excel file name to write to [{default_name}]: ").strip()
    name = typed or default_name
    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    EXCEL_PATH = os.path.join(EXCEL_DIR, name)
    _save_last_used("chibin_last_excel_name", name)
    print(f"Writing to: {EXCEL_PATH}\n")


def prompt_for_bin_folder():
    last_used = _load_last_used("chibin_last_folder")
    default_folder = last_used or ""
    while True:
        typed = input(f"Folder of .bin files to read{f' [{default_folder}]' if default_folder else ''}: ").strip().strip('"')
        folder = typed or default_folder
        if folder and os.path.isdir(folder):
            _save_last_used("chibin_last_folder", folder)
            print(f"Reading: {folder}\n")
            return folder
        print(f"[!] That folder doesn't exist: {folder}\n")


def prompt_for_step_minutes():
    typed = input(f"Minutes between readings [{DEFAULT_STEP_MINUTES}]: ").strip()
    if not typed:
        return DEFAULT_STEP_MINUTES
    try:
        return int(typed)
    except ValueError:
        print(f"[!] Not a number, using default of {DEFAULT_STEP_MINUTES}.")
        return DEFAULT_STEP_MINUTES


def _last_used_marker_path(key):
    return os.path.join(SETTINGS_DIR, f".{key}.txt")


def _load_last_used(key):
    try:
        with open(_last_used_marker_path(key), "r") as f:
            return f.read().strip() or None
    except OSError:
        return None


def _save_last_used(key, value):
    try:
        os.makedirs(SETTINGS_DIR, exist_ok=True)
        with open(_last_used_marker_path(key), "w") as f:
            f.write(value)
    except OSError:
        pass


# ------------------------------------------------------------------
# Best-effort peak height from a raw current-vs-potential curve
# ------------------------------------------------------------------
def estimate_peak_from_curve(curve, potentials, edge_points=8):
    """Fits a straight-line baseline through the first/last edge_points of
    the curve, then returns the curve value with the largest deviation
    from that line -- the standard by-hand way of reading peak height off
    a voltammogram. Not CHI's own peak-picking algorithm (see the module
    docstring for how close this comes). Returns None if curve is empty."""
    n = len(curve)
    if n == 0:
        return None
    edge = min(edge_points, max(1, n // 2))
    xs = potentials[:edge] + potentials[-edge:]
    ys = curve[:edge] + curve[-edge:]
    mean_x = sum(xs) / len(xs)
    mean_y = sum(ys) / len(ys)
    denom = sum((x - mean_x) ** 2 for x in xs)
    slope = 0.0 if denom == 0 else sum((x - mean_x) * (y - mean_y) for x, y in zip(xs, ys)) / denom
    intercept = mean_y - slope * mean_x
    residuals = [y - (slope * x + intercept) for x, y in zip(potentials, curve)]
    return max(residuals, key=abs)


# ------------------------------------------------------------------
# Reading CHI's .bin save files directly
# ------------------------------------------------------------------
# CHI's binary layout, reverse-engineered from a matched .bin/.txt pair
# (byte-for-byte cross-checked against the .txt's raw data table) rather
# than from any documentation -- CHI doesn't publish this format. Every
# value below is a little-endian 32-bit float unless noted, and none of it
# changed between a 10Hz and a 60Hz save, so it's assumed stable across
# frequencies:
#
#   bytes 0-1690      fixed-size header (technique name, run parameters --
#                      Ei/Ef/Increment/Amplitude/Frequency/Quiet Time/
#                      Sensitivity all live in here at fixed offsets)
#   bytes 1691+        raw per-point curve data, n points where
#                      n = (filesize - 1691) / 48 -- a run with zero data
#                      points (e.g. aborted before it produced anything)
#                      is exactly 1691 bytes with no data section at all
#     for each point k (0-indexed):
#       channel 1's (i1d, i1f, i1r) triplet at 1691 + 12*k
#       channels 2-4's 9 floats (i2d,i2f,i2r, i3d,i3f,i3r, i4d,i4f,i4r)
#         at 1691 + 12*n + 36*k
BIN_SIGNATURE = b"Square Wave Voltammetry"
BIN_HEADER_SIZE = 1691
BIN_FREQ_OFFSET = 1159
BIN_EI_OFFSET = 1091
BIN_EF_OFFSET = 1095
BIN_INCRE_OFFSET = 1111
BIN_CH1_POINT_BYTES = 12   # one (i1d, i1f, i1r) float32 triplet per data point
BIN_REST_POINT_BYTES = 36  # one (i2d,i2f,i2r, i3d,i3f,i3r, i4d,i4f,i4r) block per data point
BIN_BYTES_PER_POINT = BIN_CH1_POINT_BYTES + BIN_REST_POINT_BYTES  # 48


def parse_swv_bin_file(filepath):
    """Reads a CHI .bin save file directly. Returns (frequency_hz,
    {channel_num: ip_value}) or (None, {}) if this doesn't look like a CHI
    SWV .bin, or it has zero data points (e.g. an aborted run). Every ip
    value is an estimate -- see the module docstring."""
    with open(filepath, "rb") as f:
        data = f.read()

    if BIN_SIGNATURE not in data[:200] or len(data) < BIN_HEADER_SIZE:
        return None, {}

    n, remainder = divmod(len(data) - BIN_HEADER_SIZE, BIN_BYTES_PER_POINT)
    if remainder != 0 or n <= 0:
        return None, {}

    def f32(offset):
        return struct.unpack_from("<f", data, offset)[0]

    frequency = f32(BIN_FREQ_OFFSET)
    ei = f32(BIN_EI_OFFSET)
    ef = f32(BIN_EF_OFFSET)
    incre = f32(BIN_INCRE_OFFSET)
    step = incre if ef >= ei else -incre
    potentials = [ei + step * (k + 1) for k in range(n)]

    curves = {1: [], 2: [], 3: [], 4: []}
    for k in range(n):
        curves[1].append(f32(BIN_HEADER_SIZE + BIN_CH1_POINT_BYTES * k))
    base_rest = BIN_HEADER_SIZE + BIN_CH1_POINT_BYTES * n
    for k in range(n):
        row = base_rest + BIN_REST_POINT_BYTES * k
        for i, ch in enumerate((2, 3, 4)):
            curves[ch].append(f32(row + i * 12))

    ip_values = {ch: estimate_peak_from_curve(curves[ch], potentials) for ch in (1, 2, 3, 4)}
    return frequency, ip_values


# ------------------------------------------------------------------
# Excel template -- one 5-column block (Run, Ch1-4) per frequency in
# FREQUENCIES, laid out left to right in that order.
# ------------------------------------------------------------------
def block_start_col(freq_index):
    return 1 + freq_index * 5


def build_template():
    wb = Workbook()
    ws = wb.active
    ws.title = "SWV Data"
    for i, freq in enumerate(FREQUENCIES):
        col = block_start_col(i)
        ws.cell(row=1, column=col, value=f"{freq} Hz")
        ws.cell(row=2, column=col, value="Run")
        ws.cell(row=2, column=col + 1, value="Ch1")
        ws.cell(row=2, column=col + 2, value="Ch2")
        ws.cell(row=2, column=col + 3, value="Ch3")
        ws.cell(row=2, column=col + 4, value="Ch4")
    return wb


def find_next_empty_row(ws):
    row = 3
    while ws.cell(row=row, column=1).value not in (None, ""):
        row += 1
    return row


def write_ip_row(ws, row, col_start, ip_values):
    ip_values = ip_values or {}
    for offset, ch in enumerate((1, 2, 3, 4)):
        cell = ws.cell(row=row, column=col_start + offset, value=ip_values.get(ch))
        cell.number_format = NUM_FMT
        cell.font = ESTIMATED_FONT  # every value here is an estimate -- see module docstring


def fill_next_row(label, freq_data, freq_names):
    """freq_data: {freq: ip_values_dict} -- a frequency missing from this
    dict (or with None ip_values) is left blank. freq_names: {freq:
    source_filename} for the console summary line."""
    if os.path.exists(EXCEL_PATH):
        wb = load_workbook(EXCEL_PATH)
        ws = wb["SWV Data"] if "SWV Data" in wb.sheetnames else wb.active
    else:
        wb = build_template()
        ws = wb["SWV Data"]

    row = find_next_empty_row(ws)
    for i, freq in enumerate(FREQUENCIES):
        col = block_start_col(i)
        ws.cell(row=row, column=col, value=label)
        write_ip_row(ws, row, col_start=col + 1, ip_values=freq_data.get(freq))

    warned = False
    while True:
        try:
            wb.save(EXCEL_PATH)
            break
        except PermissionError:
            if not warned:
                print(f"[!] Couldn't save {EXCEL_PATH} -- it's open in Excel right now. "
                      f"Close it there; this row will save automatically as soon as you do.")
                warned = True
            time.sleep(3.0)

    names_summary = " | ".join(f"{freq}Hz: {freq_names.get(freq) or '(none)'}" for freq in FREQUENCIES)
    print(f'Filled row {row} ("{label}") -- {names_summary}')


# ------------------------------------------------------------------
CYCLE_NUMBER_RE = re.compile(r"_(\d+)$")


def extract_cycle_number(fname):
    """'10hz_blank_5.bin' -> 5. CHI doesn't suffix the very first save of a
    given name -- only repeats get _2, _3, ... -- so a filename with no
    trailing _N is treated as cycle 1."""
    stem = fname[:-4] if fname.lower().endswith(".bin") else fname
    m = CYCLE_NUMBER_RE.search(stem)
    return int(m.group(1)) if m else 1


def run_bin_folder(folder, step_minutes, first_label):
    """Reads every .bin in folder and pairs same-cycle 10Hz/60Hz saves by
    the number CHI appends to the filename (see extract_cycle_number) --
    NOT by file timestamp, which can get scrambled by moving/reorganizing
    files or OneDrive sync. Writes one row per cycle number, in ascending
    order: the first cycle labeled first_label, each one after in
    step_minutes increments."""
    groups = {}  # cycle_number -> {"freq_data": {...}, "freq_names": {...}}
    for fname in sorted(os.listdir(folder)):
        if not fname.lower().endswith(".bin"):
            continue
        fpath = os.path.join(folder, fname)
        frequency, ip_values = parse_swv_bin_file(fpath)
        if frequency is None:
            print(f"[!] {fname}: couldn't read this as a CHI SWV .bin (or it has no data points) -- skipped.")
            continue
        freq = int(round(frequency))
        if freq not in FREQUENCIES:
            print(f"[!] {fname}: {freq}Hz isn't in FREQUENCIES ({FREQUENCIES}) -- ignored.")
            continue
        cycle_num = extract_cycle_number(fname)
        g = groups.setdefault(cycle_num, {"freq_data": {}, "freq_names": {}})
        if freq in g["freq_data"]:
            print(f"[!] {fname}: another {freq}Hz file already claimed cycle number {cycle_num} "
                  f"({g['freq_names'][freq]}) -- keeping that one, skipping this one.")
            continue
        g["freq_data"][freq] = ip_values
        g["freq_names"][freq] = fname

    written = 0
    for cycle_num in sorted(groups):
        g = groups[cycle_num]
        if not any(g["freq_data"].values()):
            names_str = ", ".join(f"{f}Hz: {n}" for f, n in g["freq_names"].items())
            print(f"Skipped empty run (no data) -- {names_str}")
            continue
        label = first_label if written == 0 else f"{written * step_minutes} mins"
        fill_next_row(label, g["freq_data"], g["freq_names"])
        written += 1

    print(f"\nDone -- wrote {written} row(s) to {EXCEL_PATH}")


if __name__ == "__main__":
    if len(sys.argv) >= 2:
        folder = sys.argv[1]
        step = int(sys.argv[2]) if len(sys.argv) > 2 else DEFAULT_STEP_MINUTES
        first_label = sys.argv[3] if len(sys.argv) > 3 else DEFAULT_FIRST_LABEL
        if not os.path.isdir(folder):
            print(f"That folder doesn't exist: {folder}")
            sys.exit(1)
        prompt_for_excel_path()
        run_bin_folder(folder, step_minutes=step, first_label=first_label)
    else:
        folder = prompt_for_bin_folder()
        prompt_for_excel_path()
        step = prompt_for_step_minutes()
        run_bin_folder(folder, step_minutes=step, first_label=DEFAULT_FIRST_LABEL)

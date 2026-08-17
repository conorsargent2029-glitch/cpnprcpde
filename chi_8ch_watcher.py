"""
chi_8ch_watcher.py

(Renamed from chi_10hz_8ch_watcher.py -- it now reads both 10Hz and 60Hz,
not just 10Hz, so the old name was no longer accurate.)

A continuous folder watcher for CHI's .bin save files -- point it at a
folder and leave it running during a live overnight/24-hour SWV run over
8 channels, split across two sensors and two frequencies. It ONLY reads
files matching CHI's auto-named per-run save pattern, e.g.
"20260813_172423_SWV.bin" -- CHI writes one of these automatically for
every single run, regardless of whether the macro's "save:" command also
wrote a separately-labeled file. Any other file in the folder (.txt,
custom-labeled .bin saves, anything else) is ignored entirely, never
parsed, never written -- this project's earlier attempt to also read
custom-labeled saves broke when some runs turned out not to have gotten
one (the custom label didn't fire for every cycle), so the auto-named
save is the one reliable source: it's the only file guaranteed to exist
for every single run.

Because auto-named files carry no label text (just a timestamp), there is
no round/stage/spike detection here -- rows are just numbered in fixed
20-minute increments in chronological order: "initial reading", "20
mins", "40 mins", ...

It checks for new .bin files every few seconds, and as each one appears it:

  1. Reads CHI's binary save format directly (see the BIN_* constants for
     the byte layout, reverse-engineered from a matched 4-channel .bin/.txt
     pair and extrapolated to 8 channels -- confirmed against 5 real
     8-channel .bin files, see below). CHI's own peak-picked ip is NOT
     stored in a .bin at all, only the raw curve -- confirmed by
     exhaustively searching a known .bin for its paired .txt's exact ip
     values and finding nothing, in any float format (see
     chi_bin_to_excel.py's docstring for the full story). So every ip
     value here is an estimate, derived from the raw curve via a
     baseline-corrected peak read (estimate_peak_from_curve) -- not
     CHI's own number. Marked italic in Excel as a reminder.
  2. Keeps only files at a frequency in FREQUENCIES (10Hz and 60Hz by
     default, read straight out of the binary header). Any other
     frequency is skipped entirely and never written.
  3. Splits the 8 channels into two sensor groups, and writes one block
     per (frequency, sensor) combination -- 4 blocks total in the same
     row:
       - Ch1-4 -> Agarose Sensor 1
       - Ch5-8 -> Agarose Sensor 2
     e.g. "10 Hz - Agarose Sensor 1", "10 Hz - Agarose Sensor 2",
          "60 Hz - Agarose Sensor 1", "60 Hz - Agarose Sensor 2"
  4. Pairs a 10Hz file with the 60Hz file that landed within
     GROUP_GAP_SECONDS of it (parsed straight out of each auto-named
     filename's own embedded timestamp -- reliable regardless of OS file
     metadata, since CHI bakes it into the name at save time) into one
     row. If a file's pair never shows up before either a same-frequency
     file arrives (meaning the next cycle has started) or GROUP_GAP_SECONDS
     of real time passes with nothing new, the row is written anyway with
     that block left blank rather than waiting forever.

DAILY USE
---------
Double-click run_8ch_watcher.bat (in the same folder as this script), or
run:

    python chi_8ch_watcher.py

It will ask for the folder to watch and the Excel file to write to, then
run continuously -- checking for new .bin files every POLL_SECONDS -- until
you close the window (Ctrl+C).

BINARY LAYOUT CONFIDENCE
--------------------------
The byte layout used to decode a .bin (see the BIN_* constants) was
reverse-engineered from a matched 4-channel .bin/.txt pair, then
extrapolated to 8 channels by assuming the same per-channel pattern
continues. Confirmed against 5 real 8-channel .bin files: the header
decodes to sane parameters, the point count computed purely from file size
matches (Ef-Ei)/Increment exactly, and the decoded channel curves are
smooth with no NaN/garbage values.
"""

import os
import re
import struct
import sys
import time
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

# ----------------------- CONFIG -----------------------
EXCEL_DIR = r"C:\Users\gao22\Documents"
DEFAULT_EXCEL_NAME = "8Channel_Results.xlsx"
SETTINGS_DIR = r"C:\Users\gao22\Documents"
FREQUENCIES = [10, 60]            # only files at these frequencies are kept, in column order
STEP_MINUTES = 20                 # minutes between readings
FIRST_LABEL = "initial reading"
POLL_SECONDS = 5                  # how often to check the folder for new files
GROUP_GAP_SECONDS = 200.0         # max span between a cycle's 10Hz/60Hz saves to still count as one row
SENSOR_GROUPS = [
    ("Agarose Sensor 1", (1, 2, 3, 4)),
    ("Agarose Sensor 2", (5, 6, 7, 8)),
]
# --------------------------------------------------------

NUM_FMT = "0.00E+00"
EXCEL_PATH = None  # set at startup by prompt_for_excel_path()


def prompt_for_excel_path():
    global EXCEL_PATH
    last_used = _load_last_used("chi8ch_last_excel_name")
    default_name = last_used or DEFAULT_EXCEL_NAME
    typed = input(f"Excel file name to write to [{default_name}]: ").strip()
    name = typed or default_name
    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    EXCEL_PATH = os.path.join(EXCEL_DIR, name)
    _save_last_used("chi8ch_last_excel_name", name)
    print(f"Writing to: {EXCEL_PATH}\n")


def prompt_for_bin_folder():
    last_used = _load_last_used("chi8ch_last_folder")
    default_folder = last_used or ""
    while True:
        typed = input(f"Folder to watch for .bin files{f' [{default_folder}]' if default_folder else ''}: ").strip().strip('"')
        folder = typed or default_folder
        if folder and os.path.isdir(folder):
            _save_last_used("chi8ch_last_folder", folder)
            print(f"Watching: {folder}\n")
            return folder
        print(f"[!] That folder doesn't exist: {folder}\n")


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


ALL_CHANNELS = tuple(ch for _, channels in SENSOR_GROUPS for ch in channels)


def estimate_peak_from_curve(curve, potentials, edge_points=8):
    """Fits a straight-line baseline through the first/last edge_points of
    the curve, then returns the curve value with the largest deviation
    from that line -- the standard by-hand way of reading peak height off
    a voltammogram. Not CHI's own peak-picking algorithm, but cross-checked
    elsewhere in this project's other scripts it lands within 0.3-1.9% on
    ip and exact on peak potential. Returns None if curve is empty."""
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
# Reading CHI's .bin save files directly -- see module docstring
# ------------------------------------------------------------------
#   bytes 0-1690     fixed-size header (Ei/Ef/Increment/Frequency/etc at
#                     fixed offsets below) -- assumed identical regardless
#                     of channel count, since it's unrelated to how many
#                     channels are wired up.
#   bytes 1691+       raw per-point curve data, n points where
#                     n = (filesize - 1691) / (12 * channel_count)
#     for each point k (0-indexed):
#       channel 1's (d, f, r) triplet at 1691 + 12*k
#       every other channel's (d, f, r) triplet, in channel order, at
#         1691 + 12*n + 12*(channel_count-1)*k + 12*(position in that list)
BIN_SIGNATURE = b"Square Wave Voltammetry"
BIN_HEADER_SIZE = 1691
BIN_FREQ_OFFSET = 1159
BIN_EI_OFFSET = 1091
BIN_EF_OFFSET = 1095
BIN_INCRE_OFFSET = 1111


def parse_swv_bin_file(filepath):
    """Reads a CHI .bin save file directly. Returns (frequency_hz,
    {channel_num: ip_value}, {estimated channel numbers}) -- every channel
    comes back estimated, since CHI's own peak-picked ip isn't stored in
    the .bin at all (see module docstring). Returns (None, {}, set()) if
    this doesn't look like a CHI SWV .bin, or it has zero data points
    (e.g. an aborted run)."""
    with open(filepath, "rb") as f:
        data = f.read()

    if BIN_SIGNATURE not in data[:200] or len(data) < BIN_HEADER_SIZE:
        return None, {}, set()

    num_channels = len(ALL_CHANNELS)
    bytes_per_point = 12 * num_channels
    n, remainder = divmod(len(data) - BIN_HEADER_SIZE, bytes_per_point)
    if remainder != 0 or n <= 0:
        return None, {}, set()

    def f32(offset):
        return struct.unpack_from("<f", data, offset)[0]

    frequency = f32(BIN_FREQ_OFFSET)
    ei = f32(BIN_EI_OFFSET)
    ef = f32(BIN_EF_OFFSET)
    incre = f32(BIN_INCRE_OFFSET)
    step = incre if ef >= ei else -incre
    potentials = [ei + step * (k + 1) for k in range(n)]

    curves = {ch: [] for ch in ALL_CHANNELS}
    for k in range(n):
        curves[ALL_CHANNELS[0]].append(f32(BIN_HEADER_SIZE + 12 * k))
    rest_channels = ALL_CHANNELS[1:]
    base_rest = BIN_HEADER_SIZE + 12 * n
    rest_stride = 12 * len(rest_channels)
    for k in range(n):
        row = base_rest + rest_stride * k
        for i, ch in enumerate(rest_channels):
            curves[ch].append(f32(row + i * 12))

    ip_values = {ch: estimate_peak_from_curve(curves[ch], potentials) for ch in ALL_CHANNELS}
    return frequency, ip_values, set(ALL_CHANNELS)


# ------------------------------------------------------------------
# Excel template -- one block per (frequency, sensor) combination,
# in FREQUENCIES x SENSOR_GROUPS order, left to right.
# ------------------------------------------------------------------
BLOCKS = [(freq, name, channels) for freq in FREQUENCIES for name, channels in SENSOR_GROUPS]


def block_start_col(block_index):
    return 1 + block_index * 5


def build_template():
    wb = Workbook()
    ws = wb.active
    ws.title = "8Ch Data"

    for i, (freq, sensor_name, channels) in enumerate(BLOCKS):
        col = block_start_col(i)
        ws.cell(row=1, column=col, value=f"{freq} Hz - {sensor_name}")
        ws.cell(row=2, column=col, value="Run")
        for offset, ch in enumerate(channels):
            ws.cell(row=2, column=col + 1 + offset, value=f"Ch{ch}")

    return wb


def find_next_empty_row(ws):
    row = 3
    while ws.cell(row=row, column=1).value not in (None, ""):
        row += 1
    return row


ESTIMATED_FONT = Font(italic=True)


def write_ip_row(ws, row, col_start, ip_values, estimated_channels, channels):
    for offset, ch in enumerate(channels):
        cell = ws.cell(row=row, column=col_start + offset, value=ip_values.get(ch))
        cell.number_format = NUM_FMT
        if ch in estimated_channels:
            cell.font = ESTIMATED_FONT


def fill_next_row(label, freq_data):
    """freq_data: {freq: (ip_values_dict, estimated_channels_set)} -- a
    frequency missing from this dict leaves both of that frequency's
    blocks blank for this row."""
    if os.path.exists(EXCEL_PATH):
        wb = load_workbook(EXCEL_PATH)
        ws = wb["8Ch Data"] if "8Ch Data" in wb.sheetnames else wb.active
    else:
        wb = build_template()
        ws = wb["8Ch Data"]

    row = find_next_empty_row(ws)
    for i, (freq, sensor_name, channels) in enumerate(BLOCKS):
        col = block_start_col(i)
        ws.cell(row=row, column=col, value=label)
        ip_values, estimated_channels = freq_data.get(freq) or ({}, set())
        write_ip_row(ws, row, col_start=col + 1, ip_values=ip_values,
                     estimated_channels=estimated_channels, channels=channels)

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

    have = ", ".join(f"{freq}Hz" for freq in FREQUENCIES if (freq_data.get(freq) or ({}, None))[0])
    missing = ", ".join(f"{freq}Hz" for freq in FREQUENCIES if not (freq_data.get(freq) or ({}, None))[0])
    note = f"  [missing: {missing}]" if missing else ""
    print(f'Filled row {row} ("{label}") -- have: {have or "(none)"}{note}')


# ------------------------------------------------------------------
# CHI's auto-named per-run save, e.g. "20260813_172423_SWV.bin" -- the one
# file guaranteed to exist for every run, whether or not the macro's
# "save:" command also wrote a custom-labeled file. The date/time is
# embedded straight in the filename, so it doesn't depend on OS file
# metadata (creation/modified time), which can get reordered by copying
# files around or OneDrive sync.
AUTO_NAME_RE = re.compile(r"^(\d{8})_(\d{6})_SWV\.bin$", re.IGNORECASE)


def parse_auto_timestamp(fname):
    """'20260813_172423_SWV.bin' -> epoch seconds. Returns None if fname
    doesn't match CHI's auto-named pattern."""
    m = AUTO_NAME_RE.match(fname)
    if not m:
        return None
    dt = datetime.strptime(m.group(1) + m.group(2), "%Y%m%d%H%M%S")
    return dt.timestamp()


def label_for_index(index):
    if index == 0:
        return FIRST_LABEL
    return f"{index * STEP_MINUTES} mins"


def watch_folder(folder):
    """Polls folder every POLL_SECONDS for new files matching CHI's
    auto-named save pattern (AUTO_NAME_RE) -- anything else (.txt,
    custom-labeled .bin saves, anything) is ignored entirely and never
    read. Keeps only files at a frequency in FREQUENCIES. A 10Hz file is
    paired with the 60Hz file whose embedded timestamp is within
    GROUP_GAP_SECONDS of it into one row -- if a same-frequency file
    arrives before its pair does (meaning the next cycle has started), or
    GROUP_GAP_SECONDS of real time passes with the group otherwise idle,
    the row is written with whichever frequencies it actually got, the
    rest left blank. Rows are numbered "initial reading", "20 mins", "40
    mins", ... in the order they're written -- there's no label text to
    read a round/stage/spike from here, only a timestamp."""
    seen = set()
    current_group = None  # {"freq_data": {freq: (ip,estimated)}, "freq_names": {freq: fname}, "last_time": ts}
    written_count = 0

    print(f"Watching {folder} for new auto-named .bin files ({'/'.join(str(f) for f in FREQUENCIES)} Hz only)... Ctrl+C to stop.\n")

    def flush_group():
        nonlocal current_group, written_count
        if current_group is None:
            return
        if not any(current_group["freq_data"].values()):
            names = ", ".join(f"{f}Hz: {n}" for f, n in current_group["freq_names"].items())
            print(f"Skipped empty run (no data) -- {names or '(no files)'}")
            current_group = None
            return
        fill_next_row(label_for_index(written_count), current_group["freq_data"])
        written_count += 1
        current_group = None

    while True:
        try:
            candidates = []
            for fname in os.listdir(folder):
                if fname in seen:
                    continue
                ts = parse_auto_timestamp(fname)
                if ts is None:
                    continue
                candidates.append((ts, fname))
            candidates.sort()

            for ts, fname in candidates:
                seen.add(fname)
                fpath = os.path.join(folder, fname)

                frequency, ip_values, estimated = parse_swv_bin_file(fpath)
                if frequency is None:
                    print(f"[!] {fname}: couldn't read this as a CHI SWV .bin (or it has no data) -- skipped.")
                    continue

                freq_rounded = int(round(frequency))
                if freq_rounded not in FREQUENCIES:
                    print(f"[-] {fname}: {freq_rounded} Hz (not in {FREQUENCIES}) -- ignored.")
                    continue

                if current_group is not None:
                    gap = ts - current_group["last_time"]
                    if freq_rounded in current_group["freq_data"] or gap > GROUP_GAP_SECONDS:
                        flush_group()

                if current_group is None:
                    current_group = {"freq_data": {}, "freq_names": {}, "last_time": ts}
                else:
                    current_group["last_time"] = ts
                current_group["freq_data"][freq_rounded] = (ip_values, estimated)
                current_group["freq_names"][freq_rounded] = fname
                print(f"[+] {fname}: {freq_rounded}Hz -- queued.")

            if current_group is not None and (time.time() - current_group["last_time"]) > GROUP_GAP_SECONDS:
                flush_group()

            time.sleep(POLL_SECONDS)

        except KeyboardInterrupt:
            flush_group()
            print("\nStopped watching.")
            break


if __name__ == "__main__":
    if len(sys.argv) >= 2:
        folder = sys.argv[1]
        if not os.path.isdir(folder):
            print(f"That folder doesn't exist: {folder}")
            sys.exit(1)
        prompt_for_excel_path()
        watch_folder(folder)
    else:
        folder = prompt_for_bin_folder()
        prompt_for_excel_path()
        watch_folder(folder)

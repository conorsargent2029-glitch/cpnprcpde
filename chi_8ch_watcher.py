"""
chi_8ch_watcher.py

(Renamed from chi_10hz_8ch_watcher.py -- it now reads both 10Hz and 60Hz,
not just 10Hz, so the old name was no longer accurate.)

A continuous folder watcher for CHI's plain-text SWV exports (.txt) -- point
it at a folder and leave it running during a live overnight/24-hour SWV run
over 8 channels, split across two sensor types and two frequencies. It
checks for new .txt files every few seconds, and as each one appears it:

  1. Reads the file as plain text (CHI's normal auto-exported format --
     no binary parsing needed, unlike a .bin save with text-export turned
     off; see chi_bin_to_excel.py for that separate situation).
  2. Pulls the real "ip" value straight out of each channel's "Results:"
     section (the Difference ip, e.g. "Channel 5: ... Difference: ...
     ip = -1.989e-6A"). This is CHI's own computed peak. If a channel has
     no Difference result at all (CHI's peak-picker gave up on it -- this
     happens more on a long unattended run than a supervised one), its ip
     is instead estimated from that channel's raw difference-current curve
     further down in the same file (same baseline-corrected method
     chi_swv_to_excel.py and chi_bin_to_excel.py use), rather than left
     blank -- marked italic in Excel so you can tell it apart from CHI's
     own number.
  3. Keeps only files at a frequency in FREQUENCIES (10Hz and 60Hz by
     default, checked via the "Frequency (Hz) = " line in the file header).
     Any other frequency is skipped entirely and never written.
  4. Splits the 8 channels into two sensor-type groups, and writes one
     block per (frequency, sensor type) combination -- 4 blocks total in
     the same row:
       - Ch1-4 -> Agarose Gel Sensor  (agarose-gel-coated)
       - Ch5-8 -> Normal Sensor       (uncoated / normal type)
     e.g. "10 Hz - Agarose Gel Sensor", "10 Hz - Normal Sensor",
          "60 Hz - Agarose Gel Sensor", "60 Hz - Normal Sensor"
  5. Labels rows in fixed 20-minute increments: "initial reading", then
     "20 mins", "40 mins", ... Rows are placed in that order using the
     cycle number CHI appends to the filename (e.g. "60hz_sweat_13.txt" =
     cycle 13) -- NOT file timestamps, which can get scrambled by copying
     files around or OneDrive sync reordering things after the fact.

     Since each row now needs BOTH a 10Hz file and a 60Hz file for the same
     cycle number, a cycle's row is written once either (a) both
     frequencies' files for that cycle number have arrived, or (b) a LATER
     cycle number has shown up for a frequency that's still missing this
     cycle -- meaning that frequency's file for this cycle isn't coming
     (e.g. the macro skipped it), so the row is written with that block
     left blank rather than waiting forever.
  6. Detects when you change the save label mid-run -- e.g. spiking the
     sample partway through, "sweat" -> "sweat 5nM" -- the same way the
     filename encodes a stage change (frequency prefix and trailing cycle
     number stripped, whatever's left is the stage: "60hz_sweat_13.txt" is
     stage "sweat" cycle 13; "60Hz_sweat 5nM_1.txt" is stage "sweat 5nM"
     cycle 1 -- note CHI restarts the _N counter from 1 for each distinct
     label, so cycle numbers are only unique WITHIN a stage, not across
     the whole run). The row where the stage first changes gets the new
     stage name as its label instead of "N mins" (e.g. "sweat 5nM"), and
     is highlighted gold in Excel so the spike point is easy to spot at a
     glance. Every row after that goes back to counting "N mins" from that
     point. All of a stage's cycles are written out (in ascending
     cycle-number order) before any row from the next stage, regardless of
     which order the underlying files happened to arrive in.

EXPECTED FILE FORMAT
---------------------
Standard CHI SWV .txt export, e.g.:

    Frequency (Hz) = 10
    ...
    Results:

    Channel 1:
    Difference:
    Ep = -0.230V
    ip = -2.614e-7A
    ...
    Channel 8:
    Difference:
    Ep = -0.238V
    ip = -1.946e-6A

If a file has fewer than 8 "Channel N:" sections, whichever channels are
missing are just left blank rather than causing an error.

DAILY USE
---------
Double-click run_8ch_watcher.bat (in the same folder as this script), or
run:

    python chi_8ch_watcher.py

It will ask for the folder to watch and the Excel file to write to, then
run continuously -- checking for new .txt/.bin files every POLL_SECONDS --
until you close the window (Ctrl+C).

READING .bin FILES DIRECTLY (no .txt export)
---------------------------------------------
If a stage is missing its .txt export (e.g. CHI's auto-text-save setting
was off), this watcher also reads the named .bin save directly -- same
approach as chi_bin_to_excel.py: CHI's binary format doesn't store its own
peak-picked ip, only the raw curve, so every channel from a .bin is always
an estimate (marked italic, like any other estimated channel here).

The byte layout used to decode a .bin (see the BIN_* constants) was
reverse-engineered from a 4-channel file and extrapolated to 8 channels by
assuming the same per-channel pattern continues -- confirmed against 5 real
8-channel .bin files: the header decodes to sane parameters, the point
count computed purely from file size matches (Ef-Ei)/Increment exactly,
and the decoded channel curves are smooth with no NaN/garbage values.
"""

import os
import re
import struct
import sys
import time

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

# ----------------------- CONFIG -----------------------
EXCEL_DIR = r"C:\Users\gao22\Documents"
DEFAULT_EXCEL_NAME = "8Channel_Results.xlsx"
SETTINGS_DIR = r"C:\Users\gao22\Documents"
FREQUENCIES = [10, 60]            # only files at these frequencies are kept, in column order
STEP_MINUTES = 20                 # minutes between readings
TOTAL_HOURS = 24                  # run length this labeling scheme is built for
TOTAL_ROWS = int(TOTAL_HOURS * 60 / STEP_MINUTES) + 1  # +1 for "initial reading" -> 73
FIRST_LABEL = "initial reading"
POLL_SECONDS = 5                  # how often to check the folder for new files
SENSOR_GROUPS = [
    ("Agarose Gel Sensor", (1, 2, 3, 4)),  # agarose-gel-coated sensors
    ("Normal Sensor", (5, 6, 7, 8)),       # normal (uncoated) sensors
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


def prompt_for_txt_folder():
    last_used = _load_last_used("chi8ch_last_folder")
    default_folder = last_used or ""
    while True:
        typed = input(f"Folder to watch for .txt files{f' [{default_folder}]' if default_folder else ''}: ").strip().strip('"')
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


# ------------------------------------------------------------------
# Parsing CHI's plain-text SWV export
# ------------------------------------------------------------------
FREQUENCY_RE = re.compile(r"Frequency\s*\(Hz\)\s*=\s*([\d.]+)")
DATA_TABLE_HEADER_RE = re.compile(r"^Potential/V.*$", re.MULTILINE)

# Matches each "Channel N:" section up through its first "Difference:"
# block's ip value, e.g.:
#   Channel 5:
#   Difference:
#   Ep = -0.238V
#   ip = -1.989e-6A
CHANNEL_IP_RE = re.compile(
    r"Channel\s+(\d+):\s*\n"
    r"Difference:\s*\n"
    r"Ep\s*=\s*[-\d.]+V\s*\n"
    r"ip\s*=\s*([-\d.eE]+)A"
)

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


def compute_difference_peak_from_raw(text, channel):
    """Fallback for when CHI's own peak-picker didn't report a Difference
    ip for this channel. The raw per-point difference-current curve (the
    "i{ch}d" column in the data table) is usually still there -- pull it
    out and run it through estimate_peak_from_curve. Returns None if the
    table or column is missing."""
    header_match = DATA_TABLE_HEADER_RE.search(text)
    if not header_match:
        return None
    col_index = 1 + (channel - 1) * 3  # columns: Potential, i1d,i1f,i1r, i2d,i2f,i2r, ...
    potentials, values = [], []
    for line in text[header_match.end():].splitlines():
        line = line.strip()
        if not line or "," not in line:
            continue
        parts = line.split(",")
        if len(parts) <= col_index:
            continue
        try:
            potential = float(parts[0])
            value = float(parts[col_index])
        except ValueError:
            continue
        potentials.append(potential)
        values.append(value)
    return estimate_peak_from_curve(values, potentials)


def parse_swv_txt_file(filepath):
    """Reads a CHI SWV .txt export. Returns (frequency_hz, {channel_num:
    ip_value}, {estimated channel numbers}) -- CHI's own computed ip from
    each channel's Difference section where available, falling back to
    compute_difference_peak_from_raw (and marked as estimated) for any
    channel CHI didn't report one for. Returns (None, {}, set()) if this
    doesn't look like a CHI SWV text file."""
    try:
        with open(filepath, "r", errors="ignore") as f:
            text = f.read()
    except OSError:
        return None, {}, set()

    if "Square Wave Voltammetry" not in text:
        return None, {}, set()

    freq_match = FREQUENCY_RE.search(text)
    if not freq_match:
        return None, {}, set()
    frequency = float(freq_match.group(1))

    ip_values = {}
    for ch_match in CHANNEL_IP_RE.finditer(text):
        ch_num = int(ch_match.group(1))
        ip_val = float(ch_match.group(2))
        ip_values[ch_num] = ip_val

    estimated = set()
    for ch in ALL_CHANNELS:
        if ch not in ip_values:
            fallback = compute_difference_peak_from_raw(text, ch)
            if fallback is not None:
                ip_values[ch] = fallback
                estimated.add(ch)

    if not ip_values:
        return None, {}, set()

    return frequency, ip_values, estimated


# ------------------------------------------------------------------
# Reading CHI's .bin save files directly (no .txt export available)
# ------------------------------------------------------------------
# Layout reverse-engineered from a matched 4-channel .bin/.txt pair (see
# chi_bin_to_excel.py's docstring for how), extrapolated to 8 channels by
# assuming the pattern continues: channel 1 gets its own dedicated block
# of (d,f,r) triplets, then every other channel follows in the same
# row-major layout, one (d,f,r) triplet per channel per data point.
# Confirmed against 5 real 8-channel .bin files -- see the module
# docstring. Every value is a little-endian 32-bit float.
#
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
    {channel_num: ip_value}, {estimated channel numbers}) in the same shape
    as parse_swv_txt_file, or (None, {}, set()) if this doesn't look like a
    CHI SWV .bin, or it has zero data points (e.g. an aborted run). Every
    channel comes back estimated -- CHI's own peak-picked ip isn't stored
    in the .bin at all, only the raw curve (see chi_bin_to_excel.py's
    docstring for how that was confirmed)."""
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
# Excel template -- one block per (frequency, sensor type) combination,
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


def fill_next_row(label, freq_data, spike=False):
    """freq_data: {freq: (ip_values_dict, estimated_channels_set)} -- a
    frequency missing from this dict leaves both of that frequency's
    blocks blank for this row. spike=True highlights the row's Run cells
    gold -- used for the first row of a new stage (see extract_stage_and_cycle),
    e.g. the point a sample was spiked."""
    if os.path.exists(EXCEL_PATH):
        wb = load_workbook(EXCEL_PATH)
        ws = wb["8Ch Data"] if "8Ch Data" in wb.sheetnames else wb.active
    else:
        wb = build_template()
        ws = wb["8Ch Data"]

    row = find_next_empty_row(ws)
    for i, (freq, sensor_name, channels) in enumerate(BLOCKS):
        col = block_start_col(i)
        label_cell = ws.cell(row=row, column=col, value=label)
        if spike:
            label_cell.fill = SPIKE_FILL
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
    spike_note = "  [SPIKE POINT]" if spike else ""
    print(f'Filled row {row} ("{label}") -- have: {have or "(none)"}{note}{spike_note}')


# ------------------------------------------------------------------
HZ_PREFIX_RE = re.compile(r"^\s*\d+\s*hz_?\s*", re.IGNORECASE)
CYCLE_NUMBER_RE = re.compile(r"_(\d+)$")

SPIKE_FILL = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")


def extract_stage_and_cycle(fname):
    """'60hz_sweat_13.txt' -> ('sweat', 13). '60Hz_sweat 5nM_1.txt' ->
    ('sweat 5nM', 1). No trailing _N means cycle 1 (CHI doesn't suffix the
    very first save of a given name). Cycle numbers are only unique WITHIN
    a stage -- CHI restarts the _N counter from 1 every time the save label
    itself changes -- so callers must key on (stage, cycle), never cycle
    alone."""
    stem = fname[:-4] if fname.lower().endswith(".txt") else fname
    stem = HZ_PREFIX_RE.sub("", stem, count=1)
    m = CYCLE_NUMBER_RE.search(stem)
    if m:
        return stem[:m.start()].strip(), int(m.group(1))
    return stem.strip(), 1


def watch_folder(folder):
    """Polls folder every POLL_SECONDS for new .txt/.bin files. Keeps only
    files at a frequency in FREQUENCIES (anything else is skipped and never
    written). Files are grouped by (stage, cycle_number) -- see
    extract_stage_and_cycle -- and written out stage by stage, ascending
    cycle number within a stage. A cycle is "settled" and gets flushed once
    it's either got both frequencies, or a LATER cycle in the SAME stage
    (or any cycle in a chronologically LATER stage) has shown up for a
    frequency still missing it here -- meaning that frequency's file for
    this cycle isn't coming, so the row is written with that block left
    blank rather than waiting forever. The first row of a stage after the
    very first one overall is labeled with the stage name itself and
    highlighted gold (see fill_next_row) -- this is what marks a
    spike/relabel partway through the run.

    Stages are ordered by each stage's EARLIEST file creation time, not by
    the order os.listdir() happens to return them in -- directory order is
    alphabetical, and e.g. "10Hz_sweat 5nM_1.txt" sorts before
    "10hz_sweat.txt" (capital H < lowercase h), which would register the
    spike stage as if it came first. Creation time is trustworthy here
    specifically because this is a live folder being watched as CHI writes
    to it, not a folder that was reorganized/copied after the fact (that's
    the scenario chi_bin_to_excel.py had to stop trusting timestamps for)."""
    seen = set()
    pending = {}            # stage -> {cycle_num: {freq: (ip_values, estimated_channels)}}
    max_seen_cycle = {}     # stage -> {freq: highest cycle number seen for that stage+freq}
    stage_first_ctime = {}  # stage -> earliest file creation time seen for that stage
    freq_latest_stage = {freq: None for freq in FREQUENCIES}  # each freq's chronologically furthest-along stage
    state = {"current_stage": None, "stage_written_count": 0, "any_written": False}

    print(f"Watching {folder} for new files ({'/'.join(str(f) for f in FREQUENCIES)} Hz only)... Ctrl+C to stop.\n")

    def stage_order():
        return sorted(stage_first_ctime, key=stage_first_ctime.get)

    def try_flush():
        order = stage_order()
        if state["current_stage"] is None:
            if not order:
                return
            state["current_stage"] = order[0]

        while True:
            stage = state["current_stage"]
            stage_pending = pending.get(stage, {})
            if not stage_pending:
                # Nothing queued for this stage right now. Only safe to move on
                # once a chronologically LATER stage is already known to exist --
                # otherwise this might just be the current stage still in progress.
                order = stage_order()
                idx = order.index(stage)
                if idx + 1 < len(order):
                    state["current_stage"] = order[idx + 1]
                    state["stage_written_count"] = 0
                    continue
                return

            cycle_num = min(stage_pending)
            entry = stage_pending[cycle_num]

            def freq_settled(freq):
                if freq in entry:
                    return True
                if max_seen_cycle.get(stage, {}).get(freq, 0) > cycle_num:
                    return True
                latest = freq_latest_stage[freq]
                return latest is not None and stage_first_ctime[latest] > stage_first_ctime[stage]

            if not all(freq_settled(freq) for freq in FREQUENCIES):
                return

            del stage_pending[cycle_num]
            if not state["any_written"]:
                label, spike = FIRST_LABEL, False
            elif state["stage_written_count"] == 0:
                label, spike = stage, True
            else:
                label, spike = f"{state['stage_written_count'] * STEP_MINUTES} mins", False
            fill_next_row(label, entry, spike=spike)
            state["any_written"] = True
            state["stage_written_count"] += 1

    while True:
        try:
            for fname in sorted(os.listdir(folder)):
                lower = fname.lower()
                is_txt = lower.endswith(".txt")
                is_bin = lower.endswith(".bin")
                if not (is_txt or is_bin) or fname in seen:
                    continue
                seen.add(fname)
                fpath = os.path.join(folder, fname)

                if is_txt:
                    frequency, ip_values, estimated = parse_swv_txt_file(fpath)
                else:
                    frequency, ip_values, estimated = parse_swv_bin_file(fpath)
                if frequency is None:
                    print(f"[!] {fname}: couldn't read this as a CHI SWV file (or it has no data) -- skipped.")
                    continue

                freq_rounded = int(round(frequency))
                if freq_rounded not in FREQUENCIES:
                    print(f"[-] {fname}: {freq_rounded} Hz (not in {FREQUENCIES}) -- ignored.")
                    continue

                stage, cycle_num = extract_stage_and_cycle(fname)
                try:
                    ctime = os.path.getctime(fpath)
                except OSError:
                    ctime = time.time()
                if stage not in stage_first_ctime or ctime < stage_first_ctime[stage]:
                    stage_first_ctime[stage] = ctime
                latest = freq_latest_stage[freq_rounded]
                if latest is None or stage_first_ctime[stage] > stage_first_ctime[latest]:
                    freq_latest_stage[freq_rounded] = stage

                entry = pending.setdefault(stage, {}).setdefault(cycle_num, {})
                if freq_rounded in entry:
                    print(f"[!] {fname}: {freq_rounded}Hz stage \"{stage}\" cycle {cycle_num} already queued -- keeping the first one seen, skipping this.")
                    continue
                entry[freq_rounded] = (ip_values, estimated)
                max_seen_cycle.setdefault(stage, {})[freq_rounded] = max(
                    max_seen_cycle.setdefault(stage, {}).get(freq_rounded, 0), cycle_num)
                still_missing = [ch for ch in ALL_CHANNELS if ch not in ip_values]
                est_note = f" (estimated Ch{sorted(estimated)})" if estimated else ""
                miss_note = f" (missing Ch{still_missing})" if still_missing else ""
                print(f"[+] {fname}: {freq_rounded}Hz, stage \"{stage}\" cycle {cycle_num} -- queued{est_note}{miss_note}.")

            try_flush()
            time.sleep(POLL_SECONDS)

        except KeyboardInterrupt:
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
        folder = prompt_for_txt_folder()
        prompt_for_excel_path()
        watch_folder(folder)

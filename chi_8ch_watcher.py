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
     ip = -1.989e-6A"). This is CHI's own computed peak, not an estimate --
     .txt exports are the one format where that number is actually present
     as text (chi_bin_to_excel.py's docstring explains that .bin saves do
     NOT store this value, which is why that script has to estimate it from
     the raw curve instead).
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
     cycle number CHI appends to the filename (e.g. "10hz_..._13.txt" =
     cycle 13) -- NOT file timestamps, which can get scrambled by copying
     files around or OneDrive sync reordering things after the fact.

     Since each row now needs BOTH a 10Hz file and a 60Hz file for the same
     cycle number, a cycle's row is written once either (a) both
     frequencies' files for that cycle number have arrived, or (b) a LATER
     cycle number has shown up for a frequency that's still missing this
     cycle -- meaning that frequency's file for this cycle isn't coming
     (e.g. the macro skipped it), so the row is written with that block
     left blank rather than waiting forever.

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
run continuously -- checking for new .txt files every POLL_SECONDS -- until
you close the window (Ctrl+C).
"""

import os
import re
import sys
import time

from openpyxl import Workbook, load_workbook

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


def parse_swv_txt_file(filepath):
    """Reads a CHI SWV .txt export. Returns (frequency_hz, {channel_num:
    ip_value}) using CHI's own computed ip from each channel's Difference
    section, or (None, {}) if this doesn't look like a CHI SWV text file."""
    try:
        with open(filepath, "r", errors="ignore") as f:
            text = f.read()
    except OSError:
        return None, {}

    if "Square Wave Voltammetry" not in text:
        return None, {}

    freq_match = FREQUENCY_RE.search(text)
    if not freq_match:
        return None, {}
    frequency = float(freq_match.group(1))

    ip_values = {}
    for ch_match in CHANNEL_IP_RE.finditer(text):
        ch_num = int(ch_match.group(1))
        ip_val = float(ch_match.group(2))
        ip_values[ch_num] = ip_val

    if not ip_values:
        return None, {}

    return frequency, ip_values


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


def write_ip_row(ws, row, col_start, ip_values, channels):
    for offset, ch in enumerate(channels):
        cell = ws.cell(row=row, column=col_start + offset, value=ip_values.get(ch))
        cell.number_format = NUM_FMT
        # Note: no "estimated" marker needed here -- unlike chi_bin_to_excel.py,
        # these ip values come straight from CHI's own Results section, not
        # from a curve-based estimate.


def fill_next_row(label, freq_data):
    """freq_data: {freq: ip_values_dict} -- a frequency missing from this
    dict (or with an empty ip_values) leaves both of that frequency's
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
        ip_values = freq_data.get(freq) or {}
        write_ip_row(ws, row, col_start=col + 1, ip_values=ip_values, channels=channels)

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

    have = ", ".join(f"{freq}Hz" for freq in FREQUENCIES if freq_data.get(freq))
    missing = ", ".join(f"{freq}Hz" for freq in FREQUENCIES if not freq_data.get(freq))
    note = f"  [missing: {missing}]" if missing else ""
    print(f'Filled row {row} ("{label}") -- have: {have or "(none)"}{note}')


# ------------------------------------------------------------------
CYCLE_NUMBER_RE = re.compile(r"_(\d+)$")


def extract_cycle_number(fname):
    """'10hz_exercise_sweat_13.txt' -> 13. No trailing _N means cycle 1
    (CHI doesn't suffix the very first save of a given name)."""
    stem = fname[:-4] if fname.lower().endswith(".txt") else fname
    m = CYCLE_NUMBER_RE.search(stem)
    return int(m.group(1)) if m else 1


def label_for_index(index):
    """index 0 -> 'initial reading', index 1 -> '20 mins', index 2 ->
    '40 mins', ... continues past 1440 mins (24h) if more files show up."""
    if index == 0:
        return FIRST_LABEL
    return f"{index * STEP_MINUTES} mins"


def watch_folder(folder):
    """Polls folder every POLL_SECONDS for new .txt files. Keeps only files
    at a frequency in FREQUENCIES (anything else is skipped and never
    written). New files are queued up per cycle number, per frequency, and
    written out in ascending cycle-number order -- but NOT assuming cycle
    numbers are small consecutive integers starting near 1: CHI's per-name
    counter carries over from however many times that save name has ever
    been used before (we've seen it start in the thousands), so this always
    processes the SMALLEST cycle number actually present in pending, not
    "whatever number comes right after the last one written." A cycle is
    "settled" and gets flushed once it's either got both frequencies, or a
    LATER cycle number has shown up for a frequency still missing it (see
    the module docstring's point 5) -- so a gap in the real numbering never
    produces a fake blank row, only a genuinely skipped file does."""
    seen = set()
    pending = {}  # cycle_number -> {freq: ip_values}, for cycles not yet written
    max_seen_cycle = {freq: 0 for freq in FREQUENCIES}  # how far each frequency's stream has gotten
    written_count = 0  # how many rows have been written so far -- drives the mins label, NOT the cycle number

    print(f"Watching {folder} for new .txt files ({'/'.join(str(f) for f in FREQUENCIES)} Hz only)... Ctrl+C to stop.\n")

    def try_flush():
        nonlocal written_count
        while pending:
            cycle_num = min(pending)
            entry = pending[cycle_num]
            ready = all(freq in entry or max_seen_cycle[freq] > cycle_num for freq in FREQUENCIES)
            if not ready:
                return
            pending.pop(cycle_num, None)
            label = label_for_index(written_count)
            fill_next_row(label, entry)
            written_count += 1

    while True:
        try:
            for fname in sorted(os.listdir(folder)):
                if not fname.lower().endswith(".txt") or fname in seen:
                    continue
                seen.add(fname)
                fpath = os.path.join(folder, fname)

                frequency, ip_values = parse_swv_txt_file(fpath)
                if frequency is None:
                    print(f"[!] {fname}: couldn't read as a CHI SWV .txt (or no Results found) -- skipped.")
                    continue

                freq_rounded = int(round(frequency))
                if freq_rounded not in FREQUENCIES:
                    print(f"[-] {fname}: {freq_rounded} Hz (not in {FREQUENCIES}) -- ignored.")
                    continue

                cycle_num = extract_cycle_number(fname)
                entry = pending.setdefault(cycle_num, {})
                if freq_rounded in entry:
                    print(f"[!] {fname}: {freq_rounded}Hz cycle {cycle_num} already queued -- keeping the first one seen, skipping this.")
                    continue
                entry[freq_rounded] = ip_values
                max_seen_cycle[freq_rounded] = max(max_seen_cycle[freq_rounded], cycle_num)
                missing = [ch for ch in range(1, 9) if ch not in ip_values]
                note = f" (missing Ch{missing})" if missing else ""
                print(f"[+] {fname}: {freq_rounded}Hz, cycle {cycle_num} -- queued{note}.")

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

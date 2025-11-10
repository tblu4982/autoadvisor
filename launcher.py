import subprocess
import threading
import re
import os
from datetime import datetime
import pandas as pd
from time import sleep as wait

TARGET_SCRIPT = "project.py"
PY313 = r"C:\Users\VSU Computer Science\AppData\Local\Programs\Python\Python313\python.exe"

def split_into_n(lst, n):
    q, r = divmod(len(lst), n)
    sizes = [q + 1 if i < r else q for i in range(n)]
    out, start = [], 0
    for s in sizes:
        out.append(lst[start:start + s])
        start += s
    return out

# --- Find latest CSV (campus-v2report-search-YYYY-MM-DD.csv) ---
pattern = re.compile(r"^campus-v2report-search-(\d{4})-(\d{2})-(\d{2})\.csv$")
latest_file, latest_date = None, None

for fname in os.listdir():
    m = pattern.match(fname)
    if m:
        y, mth, d = map(int, m.groups())
        file_date = datetime(y, mth, d)
        if latest_date is None or file_date > latest_date:
            latest_date, latest_file = file_date, fname

if latest_file is None:
    raise FileNotFoundError("No 'campus-v2report-search-YYYY-MM-DD.csv' file found in current directory.")

df = pd.read_csv(latest_file, skiprows=2)
vnums = df["Student ID"].dropna().tolist()
buckets = split_into_n(vnums, 4)  # 4 buckets

# --- Threaded stderr capture ---
def drain_stderr(proc: subprocess.Popen, label: str, sink: dict):
    # Read stderr lines as text; store in sink[label]
    lines = []
    for line in iter(proc.stderr.readline, ""):
        if not line:
            break
        lines.append(line.rstrip("\n"))
    sink[label] = "\n".join(lines)

processes = []
stderr_store = {}   # label -> aggregated stderr text
threads = []

# Helper to launch one process + its stderr thread
def launch_bucket(bucket_idx: int):
    label = f"T{bucket_idx+1}"
    args = [
        PY313,
        TARGET_SCRIPT,
        str(buckets[bucket_idx]),
        f"[{label}]"
    ]
    p = subprocess.Popen(
        args,
        stdout=subprocess.DEVNULL,            # we're only capturing errors
        stderr=subprocess.PIPE,               # pipe stderr
        text=True,                            # get strings, not bytes
        bufsize=1                             # line-buffered (helps on some setups)
    )
    t = threading.Thread(target=drain_stderr, args=(p, label, stderr_store), daemon=True)
    t.start()
    processes.append((label, p))
    threads.append(t)

# --- Launch all four, with your timed gaps ---
launch_bucket(0)
wait(10)
launch_bucket(1)
wait(10)
launch_bucket(2)
wait(10)
launch_bucket(3)

print("[MAIN] All subprocesses have been launched!")

# Do other work here while they run, if you want...

# --- Wait for children to exit and threads to finish ---
for label, p in processes:
    p.wait()              # don't block earlier; we only wait here at the end
for t in threads:
    t.join()              # make sure all stderr has been drained

# --- Print captured error output (if any) ---
for i in range(4):
    label = f"T{i+1}"
    print(f"\n{label} STDERR:")
    err = stderr_store.get(label, "")
    print(err if err else "<no error output>")

"""
Remote Services Collector
-------------------------
Paste a list of computer names, click Start, and the app pulls the Windows
service list from each machine (via PowerShell / CIM, WinRM first, DCOM
fallback) and writes everything to one CSV.

* Successful machines and failed machines are shown in separate boxes.
* Failed machines are automatically rechecked on a timer (default 60 min).
  Machines that succeed on a recheck move to the Successful box and their
  services are added to the CSV.

Requirements: Windows, Python 3.8+ (tkinter is included), no extra packages.
The scan runs as the user who launched the script, so run it as an account
that has admin rights on the remote machines.
"""

import csv
import json
import os
import queue
import re
import shutil
import subprocess
import tempfile
import threading
import tkinter as tk
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from tkinter import filedialog, messagebox, scrolledtext, ttk

DEFAULT_FOLDER = r"C:\temp"
DEFAULT_CSV = "services.csv"
DEFAULT_INTERVAL_MIN = 60
DEFAULT_THREADS = 10
PER_MACHINE_TIMEOUT_SEC = 120

CREATE_NO_WINDOW = 0x08000000 if os.name == "nt" else 0
NAME_RE = re.compile(r"^[A-Za-z0-9._-]+$")

CSV_FIELDS = [
    "ComputerName", "Name", "DisplayName", "State", "StartMode",
    "StartName", "ProcessId", "PathName", "CollectedAt",
]

PS_SCRIPT = r"""
param([Parameter(Mandatory=$true)][string]$Computer)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$session = $null
try {
    try {
        $session = New-CimSession -ComputerName $Computer -OperationTimeoutSec 45
        $svc = Get-CimInstance -CimSession $session -ClassName Win32_Service
    } catch {
        if ($session) { Remove-CimSession $session -ErrorAction SilentlyContinue }
        $opt = New-CimSessionOption -Protocol Dcom
        $session = New-CimSession -ComputerName $Computer -SessionOption $opt -OperationTimeoutSec 45
        $svc = Get-CimInstance -CimSession $session -ClassName Win32_Service
    }
    $svc | Select-Object Name, DisplayName, State, StartMode, StartName, ProcessId, PathName |
        ConvertTo-Json -Compress -Depth 2
} catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
} finally {
    if ($session) { Remove-CimSession $session -ErrorAction SilentlyContinue }
}
"""


def parse_names(text):
    """Split pasted text on whitespace/commas/semicolons, de-dupe, keep order."""
    seen, names = set(), []
    for tok in re.split(r"[\s,;]+", text):
        tok = tok.strip().strip("\"'").lstrip("\\")
        if tok and tok.lower() not in seen:
            seen.add(tok.lower())
            names.append(tok)
    return names


class ServiceCollector(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Remote Services Collector")
        self.geometry("1000x860")
        self.minsize(860, 680)

        self.msg_queue = queue.Queue()
        self.stop_event = threading.Event()
        self.scanning = False

        self.rows = {}        # computer -> list of service rows
        self.success = []     # ordered list of successful computers
        self.failed = {}      # computer -> last error message
        self.csv_path = None
        self.recheck_job = None
        self.next_recheck = None

        self._tmpdir = tempfile.mkdtemp(prefix="svc_collector_")
        self.script_path = os.path.join(self._tmpdir, "get_services.ps1")
        with open(self.script_path, "w", encoding="utf-8") as f:
            f.write(PS_SCRIPT)

        self._build_ui()
        self._update_buttons()
        self.after(100, self._process_queue)
        self.after(1000, self._tick)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        pad = dict(padx=8, pady=4)

        # Input names
        top = ttk.LabelFrame(
            self, text="Computer names (paste one per line, or separate with commas/spaces)")
        top.pack(fill="x", **pad)
        self.names_text = scrolledtext.ScrolledText(top, height=8, wrap="word")
        self.names_text.pack(fill="both", expand=True, padx=6, pady=6)

        # Output + options
        opts = ttk.LabelFrame(self, text="Output & options")
        opts.pack(fill="x", **pad)
        opts.columnconfigure(1, weight=1)

        ttk.Label(opts, text="Output folder:").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        self.folder_var = tk.StringVar(value=DEFAULT_FOLDER)
        ttk.Entry(opts, textvariable=self.folder_var).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(opts, text="Browse...", command=self._browse).grid(row=0, column=2, padx=6)

        ttk.Label(opts, text="CSV file name:").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        self.csv_var = tk.StringVar(value=DEFAULT_CSV)
        ttk.Entry(opts, textvariable=self.csv_var).grid(row=1, column=1, sticky="ew", padx=6)

        extra = ttk.Frame(opts)
        extra.grid(row=2, column=0, columnspan=3, sticky="w", padx=6, pady=4)
        ttk.Label(extra, text="Parallel connections:").pack(side="left")
        self.threads_var = tk.IntVar(value=DEFAULT_THREADS)
        ttk.Spinbox(extra, from_=1, to=50, width=5, textvariable=self.threads_var).pack(
            side="left", padx=(4, 20))
        self.auto_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(extra, text="Auto-recheck failed machines every",
                        variable=self.auto_var).pack(side="left")
        self.interval_var = tk.IntVar(value=DEFAULT_INTERVAL_MIN)
        ttk.Spinbox(extra, from_=1, to=1440, width=6, textvariable=self.interval_var).pack(
            side="left", padx=4)
        ttk.Label(extra, text="minutes").pack(side="left")

        # Buttons
        btns = ttk.Frame(self)
        btns.pack(fill="x", **pad)
        self.start_btn = ttk.Button(btns, text="Start scan", command=self.start_scan)
        self.start_btn.pack(side="left", padx=4)
        self.recheck_btn = ttk.Button(btns, text="Recheck failed now", command=self.recheck_now)
        self.recheck_btn.pack(side="left", padx=4)
        self.stop_btn = ttk.Button(btns, text="Stop", command=self.stop)
        self.stop_btn.pack(side="left", padx=4)
        ttk.Button(btns, text="Clear all", command=self.clear_all).pack(side="left", padx=4)

        # Success / failed
        res = ttk.Frame(self)
        res.pack(fill="both", expand=True, **pad)
        res.columnconfigure(0, weight=1, uniform="r")
        res.columnconfigure(1, weight=1, uniform="r")
        res.rowconfigure(0, weight=1)

        self.ok_frame = ttk.LabelFrame(res, text="Successful (0)")
        self.ok_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        self.ok_box = scrolledtext.ScrolledText(self.ok_frame, height=8, state="disabled")
        self.ok_box.pack(fill="both", expand=True, padx=6, pady=6)

        self.fail_frame = ttk.LabelFrame(res, text="Failed (0)")
        self.fail_frame.grid(row=0, column=1, sticky="nsew", padx=(4, 0))
        self.fail_box = scrolledtext.ScrolledText(self.fail_frame, height=8, state="disabled")
        self.fail_box.pack(fill="both", expand=True, padx=6, pady=6)

        # Output log
        logf = ttk.LabelFrame(self, text="Output")
        logf.pack(fill="both", expand=True, **pad)
        self.log_box = scrolledtext.ScrolledText(logf, height=10, state="disabled")
        self.log_box.pack(fill="both", expand=True, padx=6, pady=6)

        # Status bar
        self.status_var = tk.StringVar(value="Idle")
        ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w").pack(
            fill="x", side="bottom")

    def _browse(self):
        folder = filedialog.askdirectory(initialdir=self.folder_var.get() or DEFAULT_FOLDER)
        if folder:
            self.folder_var.set(os.path.normpath(folder))

    def _update_buttons(self):
        self.start_btn.config(state="disabled" if self.scanning else "normal")
        self.recheck_btn.config(
            state="normal" if (not self.scanning and self.failed) else "disabled")
        self.stop_btn.config(
            state="normal" if (self.scanning or self.recheck_job) else "disabled")

    # ------------------------------------------------------------ helpers
    def _log(self, text):
        stamp = datetime.now().strftime("%H:%M:%S")
        self.log_box.config(state="normal")
        self.log_box.insert("end", f"[{stamp}] {text}\n")
        self.log_box.see("end")
        self.log_box.config(state="disabled")

    @staticmethod
    def _set_box(box, lines):
        box.config(state="normal")
        box.delete("1.0", "end")
        box.insert("end", "\n".join(lines))
        box.config(state="disabled")

    def _refresh_lists(self):
        self._set_box(self.ok_box, self.success)
        self._set_box(self.fail_box, list(self.failed.keys()))
        self.ok_frame.config(text=f"Successful ({len(self.success)})")
        self.fail_frame.config(text=f"Failed ({len(self.failed)})")

    def _interval_minutes(self):
        try:
            return max(1, int(self.interval_var.get()))
        except (tk.TclError, ValueError):
            return DEFAULT_INTERVAL_MIN

    # -------------------------------------------------------- user actions
    def start_scan(self):
        if self.scanning:
            return
        names = parse_names(self.names_text.get("1.0", "end"))
        if not names:
            messagebox.showwarning("No computers", "Paste at least one computer name.")
            return

        folder = self.folder_var.get().strip() or DEFAULT_FOLDER
        csv_name = self.csv_var.get().strip() or DEFAULT_CSV
        if any(c in csv_name for c in '<>:"/\\|?*'):
            messagebox.showwarning("Invalid file name", "The CSV name contains invalid characters.")
            return
        if not csv_name.lower().endswith(".csv"):
            csv_name += ".csv"
            self.csv_var.set(csv_name)

        try:
            os.makedirs(folder, exist_ok=True)
        except OSError as e:
            messagebox.showerror("Output folder", f"Cannot use folder:\n{e}")
            return

        path = os.path.join(folder, csv_name)
        if os.path.exists(path) and not messagebox.askyesno(
                "File exists", f"{path}\nalready exists. Overwrite it?"):
            return

        self._cancel_recheck()
        self.rows.clear()
        self.success.clear()
        self.failed.clear()
        self._refresh_lists()
        self.csv_path = path
        self.stop_event.clear()
        self._log("=" * 50)
        self._run_scan(names, "Scan")

    def recheck_now(self):
        if self.scanning or not self.failed:
            return
        self._cancel_recheck()
        self.stop_event.clear()
        self._run_scan(list(self.failed.keys()), "Recheck")

    def stop(self):
        self.stop_event.set()
        self._cancel_recheck()
        if self.scanning:
            self.status_var.set("Stopping... waiting for in-flight connections to finish")
            self._log("Stop requested.")
        else:
            self._log("Automatic recheck cancelled.")
            self.status_var.set("Idle")
        self._update_buttons()

    def clear_all(self):
        if self.scanning:
            messagebox.showinfo("Busy", "Stop the current scan first.")
            return
        self._cancel_recheck()
        self.rows.clear()
        self.success.clear()
        self.failed.clear()
        self._refresh_lists()
        self.status_var.set("Idle")
        self._update_buttons()

    # ------------------------------------------------------------ scanning
    def _run_scan(self, names, label):
        self.scanning = True
        self._update_buttons()
        try:
            workers = max(1, min(50, int(self.threads_var.get())))
        except (tk.TclError, ValueError):
            workers = DEFAULT_THREADS
        self.status_var.set(f"{label} in progress: {len(names)} machine(s)...")
        threading.Thread(
            target=self._scan_worker, args=(names, label, workers), daemon=True).start()

    def _scan_worker(self, names, label, workers):
        self.msg_queue.put(("log", f"{label}: querying {len(names)} machine(s)..."))
        with ThreadPoolExecutor(max_workers=workers) as ex:
            futures = [ex.submit(self._query_machine, n) for n in names]
            for fut in as_completed(futures):
                name, ok, payload = fut.result()
                if ok:
                    self.msg_queue.put(("ok", name, payload))
                elif payload == "Cancelled":
                    continue
                else:
                    self.msg_queue.put(("fail", name, payload))
        self.msg_queue.put(("done", label))

    def _query_machine(self, name):
        """Runs in a worker thread. Returns (name, ok, rows_or_error)."""
        if self.stop_event.is_set():
            return name, False, "Cancelled"
        if not NAME_RE.match(name):
            return name, False, "Invalid computer name"

        cmd = [
            "powershell.exe", "-NoProfile", "-NonInteractive",
            "-ExecutionPolicy", "Bypass",
            "-File", self.script_path, "-Computer", name,
        ]
        try:
            p = subprocess.run(
                cmd, capture_output=True, timeout=PER_MACHINE_TIMEOUT_SEC,
                creationflags=CREATE_NO_WINDOW)
        except subprocess.TimeoutExpired:
            return name, False, f"Timed out after {PER_MACHINE_TIMEOUT_SEC}s"
        except FileNotFoundError:
            return name, False, "powershell.exe not found (Windows required)"
        except Exception as e:  # noqa: BLE001
            return name, False, f"{type(e).__name__}: {e}"

        out = p.stdout.decode("utf-8", "replace").strip()
        err = p.stderr.decode("utf-8", "replace").strip()
        if p.returncode != 0:
            msg = " ".join(err.split()) or f"PowerShell exited with code {p.returncode}"
            return name, False, msg[:300]
        if not out:
            return name, False, "No data returned"

        try:
            data = json.loads(out)
        except json.JSONDecodeError:
            return name, False, "Could not parse output from remote machine"
        if isinstance(data, dict):
            data = [data]

        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        rows = []
        for svc in data:
            row = {"ComputerName": name, "CollectedAt": stamp}
            for key in CSV_FIELDS[1:-1]:
                row[key] = svc.get(key, "")
            rows.append(row)
        return name, True, rows

    # -------------------------------------------- queue (runs on UI thread)
    def _process_queue(self):
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                kind = msg[0]
                if kind == "log":
                    self._log(msg[1])
                elif kind == "ok":
                    _, name, rows = msg
                    self.rows[name] = rows
                    if name not in self.success:
                        self.success.append(name)
                    self.failed.pop(name, None)
                    self._log(f"OK     {name} ({len(rows)} services)")
                    self._refresh_lists()
                elif kind == "fail":
                    _, name, reason = msg
                    self.failed[name] = reason
                    self._log(f"FAILED {name}: {reason}")
                    self._refresh_lists()
                elif kind == "done":
                    self._scan_finished(msg[1])
        except queue.Empty:
            pass
        self.after(100, self._process_queue)

    def _scan_finished(self, label):
        self.scanning = False
        self._write_csv()
        self._log(f"{label} complete: {len(self.success)} successful, {len(self.failed)} failed.")

        if self.stop_event.is_set():
            self.status_var.set("Stopped")
        elif self.failed and self.auto_var.get():
            self._schedule_recheck()
        elif not self.failed:
            self._log("All machines succeeded - nothing left to recheck.")
            self.status_var.set("Done - all machines succeeded")
        else:
            self.status_var.set("Done (auto-recheck is off)")
        self._update_buttons()

    # ------------------------------------------------------------ CSV
    def _write_csv(self):
        if not self.csv_path:
            return
        try:
            self._dump_csv(self.csv_path)
            self._log(f"CSV saved: {self.csv_path}")
        except PermissionError:
            base, ext = os.path.splitext(self.csv_path)
            alt = f"{base}_{datetime.now():%H%M%S}{ext}"
            try:
                self._dump_csv(alt)
                self._log(f"WARNING: {self.csv_path} is locked (open in Excel?). Saved to {alt}")
            except OSError as e:
                self._log(f"ERROR writing CSV: {e}")
        except OSError as e:
            self._log(f"ERROR writing CSV: {e}")

    def _dump_csv(self, path):
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=CSV_FIELDS)
            writer.writeheader()
            for name in self.success:
                writer.writerows(self.rows.get(name, []))

    # --------------------------------------------------- hourly recheck
    def _schedule_recheck(self):
        self._cancel_recheck()
        minutes = self._interval_minutes()
        self.next_recheck = datetime.now() + timedelta(minutes=minutes)
        self.recheck_job = self.after(minutes * 60 * 1000, self._auto_recheck)
        self._log(
            f"{len(self.failed)} machine(s) failed. Rechecking at "
            f"{self.next_recheck:%H:%M:%S} (in {minutes} min).")

    def _cancel_recheck(self):
        if self.recheck_job is not None:
            self.after_cancel(self.recheck_job)
        self.recheck_job = None
        self.next_recheck = None

    def _auto_recheck(self):
        self.recheck_job = None
        self.next_recheck = None
        if self.scanning:  # shouldn't happen, but try again in a minute
            self.next_recheck = datetime.now() + timedelta(minutes=1)
            self.recheck_job = self.after(60_000, self._auto_recheck)
            return
        if self.failed:
            self.stop_event.clear()
            self._run_scan(list(self.failed.keys()), "Recheck")

    def _tick(self):
        if self.next_recheck and not self.scanning:
            remaining = int((self.next_recheck - datetime.now()).total_seconds())
            if remaining > 0:
                m, s = divmod(remaining, 60)
                h, m = divmod(m, 60)
                self.status_var.set(
                    f"{len(self.failed)} failed. Next automatic recheck in {h:02d}:{m:02d}:{s:02d}")
        self._update_buttons()
        self.after(1000, self._tick)

    # ------------------------------------------------------------- exit
    def _on_close(self):
        self.stop_event.set()
        self._cancel_recheck()
        shutil.rmtree(self._tmpdir, ignore_errors=True)
        self.destroy()


if __name__ == "__main__":
    ServiceCollector().mainloop()

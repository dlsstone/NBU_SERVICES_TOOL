import csv
import json
import os
import queue
import re
import subprocess
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk


APP_NAME = "NetBackup Services Tool"

DEFAULT_MASTER_SERVERS = [
    "USDC-NBU0002",
    "USDR-NBU0002",
    "BEDC-NBU0002",
    "Custom / Manual Entry",
]

DEFAULT_BPPLCLIENTS_PATH = r"C:\Program Files\Veritas\NetBackup\bin\admincmd\bpplclients.exe"

BPPLCLIENTS_FALLBACK_PATHS = [
    r"C:\Program Files\Veritas\NetBackup\bin\admincmd\bpplclients.exe",
    r"D:\Program Files\Veritas\NetBackup\bin\admincmd\bpplclients.exe",
    r"E:\Program Files\Veritas\NetBackup\bin\admincmd\bpplclients.exe",
    r"F:\Program Files\Veritas\NetBackup\bin\admincmd\bpplclients.exe",
]


class NetBackupServiceTool(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1280x760")
        self.minsize(1050, 650)
        self.clients = []
        self.connectivity_results = []
        self.worker_thread = None
        self.ui_queue = queue.Queue()
        self.dark_mode_var = tk.BooleanVar(value=False)
        self._build_ui()
        self._apply_theme()
        self.after(100, self._maximize_window)
        self.after(150, self._process_ui_queue)

    def _maximize_window(self):
        try:
            self.state("zoomed")
        except Exception:
            try:
                self.attributes("-zoomed", True)
            except Exception:
                pass

    def _build_ui(self):
        self.root_frame = ttk.Frame(self, padding=12)
        self.root_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(self.root_frame, text=APP_NAME, font=("Segoe UI", 17, "bold")).pack(anchor=tk.W)
        ttk.Label(
            self.root_frame,
            text="Phase 1: Discover NetBackup clients and test WinRM connectivity through the selected Master Server.",
        ).pack(anchor=tk.W, pady=(2, 12))

        config = ttk.LabelFrame(self.root_frame, text="Configuration", padding=10)
        config.pack(fill=tk.X, pady=(0, 10))
        config.columnconfigure(1, weight=1)
        config.columnconfigure(3, weight=1)

        ttk.Label(config, text="Master Server").grid(row=0, column=0, sticky=tk.W, padx=(0, 8), pady=4)
        self.master_var = tk.StringVar(value=DEFAULT_MASTER_SERVERS[0])
        self.master_combo = ttk.Combobox(
            config, textvariable=self.master_var, values=DEFAULT_MASTER_SERVERS, state="readonly", width=28
        )
        self.master_combo.grid(row=0, column=1, sticky=tk.W, pady=4)
        self.master_combo.bind("<<ComboboxSelected>>", self._on_master_selected)

        ttk.Label(config, text="Custom Master").grid(row=0, column=2, sticky=tk.W, padx=(20, 8), pady=4)
        self.custom_master_var = tk.StringVar()
        self.custom_master_entry = ttk.Entry(config, textvariable=self.custom_master_var, width=32, state=tk.DISABLED)
        self.custom_master_entry.grid(row=0, column=3, sticky=tk.W, pady=4)

        ttk.Label(config, text="bpplclients.exe Path").grid(row=1, column=0, sticky=tk.W, padx=(0, 8), pady=4)
        self.bpplclients_path_var = tk.StringVar(value=DEFAULT_BPPLCLIENTS_PATH)
        self.bpplclients_entry = ttk.Entry(config, textvariable=self.bpplclients_path_var)
        self.bpplclients_entry.grid(row=1, column=1, columnspan=3, sticky=tk.EW, pady=4)

        ttk.Label(
            config,
            text="If this path is not found, the app automatically tries C:, D:, E:, and F: Program Files locations.",
        ).grid(row=2, column=1, columnspan=3, sticky=tk.W, pady=(0, 4))

        ttk.Label(config, text="Reports / Logs Folder").grid(row=3, column=0, sticky=tk.W, padx=(0, 8), pady=4)
        self.output_dir_var = tk.StringVar(value=str(Path.cwd() / "reports"))
        self.output_dir_entry = ttk.Entry(config, textvariable=self.output_dir_var)
        self.output_dir_entry.grid(row=3, column=1, columnspan=2, sticky=tk.EW, pady=4)
        ttk.Button(config, text="Browse...", command=self._browse_output_dir).grid(row=3, column=3, sticky=tk.W, padx=(8, 0), pady=4)

        self.theme_toggle = ttk.Checkbutton(config, text="Dark Mode", variable=self.dark_mode_var, command=self._apply_theme)
        self.theme_toggle.grid(row=4, column=1, columnspan=3, sticky=tk.W, pady=(4, 0))

        action_frame = ttk.Frame(self.root_frame)
        action_frame.pack(fill=tk.X, pady=(0, 10))
        self.get_clients_btn = ttk.Button(action_frame, text="1. Get Clients", command=self.get_clients)
        self.get_clients_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.test_connectivity_btn = ttk.Button(action_frame, text="2. Test Connectivity", command=self.test_connectivity, state=tk.DISABLED)
        self.test_connectivity_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.export_connectivity_btn = ttk.Button(action_frame, text="Export Connectivity Report", command=self.export_connectivity_report, state=tk.DISABLED)
        self.export_connectivity_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.clear_btn = ttk.Button(action_frame, text="Clear Results", command=self.clear_results)
        self.clear_btn.pack(side=tk.LEFT)

        self.status_var = tk.StringVar(value="Ready.")
        self.progress = ttk.Progressbar(action_frame, mode="indeterminate")
        self.progress.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(12, 0))

        table_frame = ttk.LabelFrame(self.root_frame, text="Connectivity Results - Problem Clients Sort to Top", padding=8)
        table_frame.pack(fill=tk.BOTH, expand=True)
        columns = ("timestamp", "master_server", "client_name", "dns_status", "ping_status", "port_5985_status", "wsman_status", "overall_status", "failure_reason")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=18)
        headings = {"timestamp": "Timestamp", "master_server": "Master Server", "client_name": "Client", "dns_status": "DNS", "ping_status": "Ping", "port_5985_status": "Port 5985", "wsman_status": "Test-WSMan", "overall_status": "Overall", "failure_reason": "Failure Reason"}
        widths = {"timestamp": 145, "master_server": 120, "client_name": 175, "dns_status": 90, "ping_status": 90, "port_5985_status": 95, "wsman_status": 105, "overall_status": 95, "failure_reason": 390}
        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=widths[col], anchor=tk.W)
        y_scroll = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        x_scroll = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.tree.grid(row=0, column=0, sticky=tk.NSEW)
        y_scroll.grid(row=0, column=1, sticky=tk.NS)
        x_scroll.grid(row=1, column=0, sticky=tk.EW)
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        bottom = ttk.Frame(self.root_frame)
        bottom.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(bottom, textvariable=self.status_var).pack(side=tk.LEFT)

    def _apply_theme(self):
        dark = self.dark_mode_var.get()
        bg = "#1e1e1e" if dark else "#f0f0f0"
        fg = "#ffffff" if dark else "#000000"
        field_bg = "#2d2d2d" if dark else "#ffffff"
        selected_bg = "#3a5f8a" if dark else "#0078d7"
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(".", background=bg, foreground=fg, fieldbackground=field_bg)
        style.configure("TFrame", background=bg)
        style.configure("TLabel", background=bg, foreground=fg)
        style.configure("TLabelframe", background=bg, foreground=fg)
        style.configure("TLabelframe.Label", background=bg, foreground=fg)
        style.configure("TCheckbutton", background=bg, foreground=fg)
        style.configure("TEntry", fieldbackground=field_bg, foreground=fg)
        style.configure("TCombobox", fieldbackground=field_bg, foreground=fg)
        style.configure("Treeview", background=field_bg, foreground=fg, fieldbackground=field_bg, rowheight=24)
        style.configure("Treeview.Heading", background=bg, foreground=fg)
        style.map("Treeview", background=[("selected", selected_bg)], foreground=[("selected", "#ffffff")])
        self.configure(background=bg)
        self.tree.tag_configure("problem", background="#4a1f1f" if dark else "#ffe5e5")
        self.tree.tag_configure("online", background="#1f3d2a" if dark else "#e8f5e9")

    def _on_master_selected(self, _event=None):
        if self.master_var.get() == "Custom / Manual Entry":
            self.custom_master_entry.configure(state=tk.NORMAL)
            self.custom_master_entry.focus_set()
        else:
            self.custom_master_entry.configure(state=tk.DISABLED)

    def _browse_output_dir(self):
        folder = filedialog.askdirectory(title="Select reports/logs folder")
        if folder:
            self.output_dir_var.set(folder)

    def _selected_master(self):
        selected = self.master_var.get().strip()
        if selected == "Custom / Manual Entry":
            selected = self.custom_master_var.get().strip()
        if not selected:
            raise ValueError("Please select or enter a NetBackup Master Server.")
        return selected

    def _set_busy(self, busy, status=None):
        for button in [self.get_clients_btn, self.test_connectivity_btn, self.export_connectivity_btn, self.clear_btn]:
            button.configure(state=tk.DISABLED if busy else tk.NORMAL)
        if not busy:
            if not self.clients:
                self.test_connectivity_btn.configure(state=tk.DISABLED)
            if not self.connectivity_results:
                self.export_connectivity_btn.configure(state=tk.DISABLED)
        if busy:
            self.progress.start(10)
        else:
            self.progress.stop()
        if status:
            self.status_var.set(status)

    def _run_worker(self, target):
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showwarning(APP_NAME, "A task is already running.")
            return
        self._set_busy(True, "Working...")
        self.worker_thread = threading.Thread(target=target, daemon=True)
        self.worker_thread.start()

    def _process_ui_queue(self):
        try:
            while True:
                item = self.ui_queue.get_nowait()
                action = item.get("action")
                if action == "status":
                    self.status_var.set(item["message"])
                elif action == "clients_loaded":
                    self.clients = item["clients"]
                    self.connectivity_results = []
                    self._clear_tree()
                    msg = f"Loaded {len(self.clients)} client(s)."
                    if item.get("bpplclients_path_used"):
                        msg += f" bpplclients path used: {item['bpplclients_path_used']}"
                    self._set_busy(False, msg)
                    self.test_connectivity_btn.configure(state=tk.NORMAL if self.clients else tk.DISABLED)
                    self.export_connectivity_btn.configure(state=tk.DISABLED)
                elif action == "connectivity_row":
                    row = item["row"]
                    self.connectivity_results.append(row)
                    self._insert_result_row(row)
                elif action == "connectivity_complete":
                    self.connectivity_results = sort_connectivity_results(self.connectivity_results)
                    self._rebuild_tree()
                    self._set_busy(False, item["message"])
                    if self.connectivity_results:
                        self.export_connectivity_btn.configure(state=tk.NORMAL)
                elif action == "error":
                    self._set_busy(False, "Error.")
                    messagebox.showerror(APP_NAME, item["message"])
        except queue.Empty:
            pass
        self.after(150, self._process_ui_queue)

    def _clear_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

    def _rebuild_tree(self):
        self._clear_tree()
        for row in self.connectivity_results:
            self._insert_result_row(row)

    def _insert_result_row(self, row):
        values = (row.get("Timestamp", ""), row.get("MasterServer", ""), row.get("ClientName", ""), row.get("DNSStatus", ""), row.get("PingStatus", ""), row.get("Port5985Status", ""), row.get("WSManStatus", ""), row.get("OverallStatus", ""), row.get("FailureReason", ""))
        tag = "online" if row.get("OverallStatus") == "Online" else "problem"
        self.tree.insert("", tk.END, values=values, tags=(tag,))

    def clear_results(self):
        self.clients = []
        self.connectivity_results = []
        self._clear_tree()
        self.test_connectivity_btn.configure(state=tk.DISABLED)
        self.export_connectivity_btn.configure(state=tk.DISABLED)
        self.status_var.set("Results cleared.")

    def get_clients(self):
        def worker():
            try:
                master = self._selected_master()
                bpplclients_path = self.bpplclients_path_var.get().strip()
                if not bpplclients_path:
                    raise ValueError("Please provide the bpplclients.exe path.")
                self.ui_queue.put({"action": "status", "message": f"Getting clients from {master}..."})
                stdout, path_used = run_bpplclients(master, bpplclients_path)
                clients = parse_bpplclients_output(stdout)
                client_rows = [{"MasterServer": master, "ClientName": client} for client in clients]
                self.ui_queue.put({"action": "clients_loaded", "clients": client_rows, "bpplclients_path_used": path_used})
            except Exception as exc:
                self.ui_queue.put({"action": "error", "message": str(exc)})
        self._run_worker(worker)

    def test_connectivity(self):
        def worker():
            try:
                if not self.clients:
                    raise ValueError("No clients loaded. Run Get Clients first.")
                self.connectivity_results = []
                self.after(0, self._clear_tree)
                total = len(self.clients)
                local_results = []
                for index, client_row in enumerate(self.clients, start=1):
                    master = client_row["MasterServer"]
                    client = client_row["ClientName"]
                    self.ui_queue.put({"action": "status", "message": f"Testing {client} from {master} ({index}/{total})..."})
                    result = test_client_connectivity(master, client)
                    local_results.append(result)
                    self.ui_queue.put({"action": "connectivity_row", "row": result})
                offline_count = sum(1 for row in local_results if row.get("OverallStatus") != "Online")
                online_count = total - offline_count
                self.ui_queue.put({"action": "connectivity_complete", "message": f"Connectivity test complete. Online: {online_count}. Problem/offline clients: {offline_count}. Problems sorted to top."})
            except Exception as exc:
                self.ui_queue.put({"action": "error", "message": str(exc)})
        self._run_worker(worker)

    def export_connectivity_report(self):
        try:
            if not self.connectivity_results:
                messagebox.showinfo(APP_NAME, "No connectivity results to export.")
                return
            output_dir_text = self.output_dir_var.get().strip()
            if not output_dir_text:
                messagebox.showwarning(APP_NAME, "Please select a reports/logs folder first.")
                return
            output_dir = Path(output_dir_text)
            output_dir.mkdir(parents=True, exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            report_path = output_dir / f"Connectivity_Report_{timestamp}.csv"
            sorted_rows = sort_connectivity_results(self.connectivity_results)
            fieldnames = ["Timestamp", "MasterServer", "ClientName", "DNSStatus", "PingStatus", "Port5985Status", "WSManStatus", "OverallStatus", "FailureReason"]
            with report_path.open("w", newline="", encoding="utf-8-sig") as handle:
                writer = csv.DictWriter(handle, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(sorted_rows)
            self.status_var.set(f"Exported report: {report_path}")
            messagebox.showinfo(APP_NAME, f"Connectivity report exported successfully:\n\n{report_path}")
            try:
                os.startfile(output_dir)
            except Exception:
                pass
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Failed to export connectivity report:\n\n{exc}")
            self.status_var.set("Export failed.")


def sort_connectivity_results(rows):
    return sorted(rows, key=lambda row: (1 if row.get("OverallStatus") == "Online" else 0, row.get("ClientName", "").lower()))


def run_powershell(script_text, timeout=300):
    command = ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script_text]
    result = subprocess.run(command, text=True, capture_output=True, timeout=timeout, creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0)
    if result.returncode != 0:
        error = result.stderr.strip() or result.stdout.strip() or "Unknown PowerShell error."
        raise RuntimeError(error)
    return result.stdout


def ps_single_quote(value):
    return "'" + str(value).replace("'", "''") + "'"


def run_bpplclients(master_server, bpplclients_path):
    fallback_paths = [bpplclients_path]
    for candidate in BPPLCLIENTS_FALLBACK_PATHS:
        if candidate not in fallback_paths:
            fallback_paths.append(candidate)
    fallback_json = json.dumps(fallback_paths)
    script = f"""
$ErrorActionPreference = 'Stop'
$master = {ps_single_quote(master_server)}
$candidatePaths = ConvertFrom-Json -InputObject {ps_single_quote(fallback_json)}

Invoke-Command -ComputerName $master -ScriptBlock {{
    param($CandidatePaths)
    $pathUsed = $null
    foreach ($candidate in $CandidatePaths) {{
        if (Test-Path -LiteralPath $candidate) {{
            $pathUsed = $candidate
            break
        }}
    }}
    if (-not $pathUsed) {{
        throw "bpplclients.exe not found. Tried: $($CandidatePaths -join ', ')"
    }}
    $output = & $pathUsed -allunique
    [PSCustomObject]@{{
        PathUsed = $pathUsed
        Output = ($output -join [Environment]::NewLine)
    }}
}} -ArgumentList (,$candidatePaths) | ConvertTo-Json -Depth 5 -Compress
"""
    stdout = run_powershell(script, timeout=300).strip()
    if not stdout:
        raise RuntimeError("No bpplclients output was returned.")
    data = json.loads(stdout)
    if isinstance(data, list):
        data = data[0]
    return data.get("Output", ""), data.get("PathUsed", "")


def parse_bpplclients_output(output):
    clients = []
    for raw_line in output.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if line.startswith("Name") and "Hostname" in line:
            continue
        if re.match(r"^-+\s+-+\s+-+$", line):
            continue
        parts = re.split(r"\s{2,}", line)
        if len(parts) >= 3:
            client = parts[-1].strip()
            if client and client not in clients:
                clients.append(client)
    return clients


def test_client_connectivity(master_server, client_name):
    script = f"""
$ErrorActionPreference = 'Continue'
$master = {ps_single_quote(master_server)}
$client = {ps_single_quote(client_name)}
$result = Invoke-Command -ComputerName $master -ScriptBlock {{
    param($Client)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $dnsStatus = "Failed"
    $pingStatus = "Failed"
    $portStatus = "Failed"
    $wsmanStatus = "Failed"
    $failureReasons = New-Object System.Collections.Generic.List[string]
    try {{
        Resolve-DnsName -Name $Client -ErrorAction Stop | Out-Null
        $dnsStatus = "OK"
    }} catch {{
        $failureReasons.Add("DNS failed: $($_.Exception.Message)")
    }}
    try {{
        if (Test-Connection -ComputerName $Client -Count 1 -Quiet -ErrorAction Stop) {{
            $pingStatus = "OK"
        }} else {{
            $failureReasons.Add("Ping failed or timed out")
        }}
    }} catch {{
        $failureReasons.Add("Ping failed: $($_.Exception.Message)")
    }}
    try {{
        $tnc = Test-NetConnection -ComputerName $Client -Port 5985 -WarningAction SilentlyContinue
        if ($tnc.TcpTestSucceeded) {{
            $portStatus = "OK"
        }} else {{
            $failureReasons.Add("Port 5985 failed")
        }}
    }} catch {{
        $failureReasons.Add("Port 5985 test failed: $($_.Exception.Message)")
    }}
    try {{
        Test-WSMan -ComputerName $Client -ErrorAction Stop | Out-Null
        $wsmanStatus = "OK"
    }} catch {{
        $failureReasons.Add("Test-WSMan failed: $($_.Exception.Message)")
    }}
    $overall = if ($wsmanStatus -eq "OK") {{ "Online" }} else {{ "Offline" }}
    [PSCustomObject]@{{
        Timestamp = $timestamp
        MasterServer = $env:COMPUTERNAME
        ClientName = $Client
        DNSStatus = $dnsStatus
        PingStatus = $pingStatus
        Port5985Status = $portStatus
        WSManStatus = $wsmanStatus
        OverallStatus = $overall
        FailureReason = ($failureReasons -join " | ")
    }}
}} -ArgumentList $client
$result | ConvertTo-Json -Depth 5 -Compress
"""
    stdout = run_powershell(script, timeout=180).strip()
    if not stdout:
        raise RuntimeError(f"No connectivity result returned for {client_name}.")
    data = json.loads(stdout)
    if isinstance(data, list):
        data = data[0]
    data["MasterServer"] = master_server
    return data


if __name__ == "__main__":
    app = NetBackupServiceTool()
    app.mainloop()

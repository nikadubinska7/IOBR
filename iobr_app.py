from __future__ import annotations

import argparse
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from iobr_core import FILTER_COLUMNS, generate_reports, get_filter_metadata


FILTER_LABELS = {
    "season": "Season",
    "sold_to": "Sold-to",
    "d_account": "D-Account",
    "document_type": "SO document type",
    "distribution_method": "Distribution method",
}


class FilterBox(ttk.Frame):
    def __init__(self, master, title: str, on_change=None):
        super().__init__(master)
        self.values: list[tuple[str, str]] = []
        self.on_change = on_change

        header = ttk.Frame(self)
        header.pack(fill="x")
        ttk.Label(header, text=title).pack(side="left")
        ttk.Button(header, text="All", width=5, command=self.select_all).pack(side="right")
        ttk.Button(header, text="Clear", width=6, command=self.clear).pack(side="right", padx=(0, 4))

        body = ttk.Frame(self)
        body.pack(fill="both", expand=True, pady=(4, 0))
        self.listbox = tk.Listbox(body, selectmode="extended", exportselection=False, height=8)
        scrollbar = ttk.Scrollbar(body, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.listbox.bind("<<ListboxSelect>>", self._on_select)

    def set_values(self, values: list[tuple[str, str]], selected: set[str] | None = None) -> None:
        self.values = values
        self.listbox.delete(0, tk.END)
        selected = selected or set()
        for index, (value, label) in enumerate(values):
            self.listbox.insert(tk.END, label)
            if value in selected:
                self.listbox.select_set(index)

    def selected_values(self) -> list[str]:
        return [self.values[index][0] for index in self.listbox.curselection()]

    def select_all(self) -> None:
        self.listbox.select_set(0, tk.END)
        self._notify_change()

    def clear(self) -> None:
        self.listbox.selection_clear(0, tk.END)
        self._notify_change()

    def _on_select(self, _event) -> None:
        self._notify_change()

    def _notify_change(self) -> None:
        if self.on_change:
            self.on_change()


class IobrApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("iOBR Extractor")
        self.geometry("1120x760")
        self.minsize(980, 680)

        self.source_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.status = tk.StringVar(value="Choose an OBR_Output workbook to begin.")
        self.filter_boxes: dict[str, FilterBox] = {}
        self.filter_options: dict[str, list[tuple[str, str]]] = {}
        self.accounts_by_sold_to: dict[str, list[str]] = {}
        self.messages: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker: threading.Thread | None = None

        self._build_ui()
        self.after(100, self._drain_messages)

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=14)
        root.pack(fill="both", expand=True)

        path_panel = ttk.Frame(root)
        path_panel.pack(fill="x")

        ttk.Label(path_panel, text="Base file").grid(row=0, column=0, sticky="w")
        ttk.Entry(path_panel, textvariable=self.source_path).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(path_panel, text="Browse...", command=self.choose_source).grid(row=0, column=2)
        ttk.Button(path_panel, text="Load Filters", command=self.load_filters).grid(row=0, column=3, padx=(8, 0))

        ttk.Label(path_panel, text="Output folder").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(path_panel, textvariable=self.output_path).grid(row=1, column=1, sticky="ew", padx=8, pady=(8, 0))
        ttk.Button(path_panel, text="Browse...", command=self.choose_output).grid(row=1, column=2, pady=(8, 0))
        ttk.Button(path_panel, text="Generate", command=self.generate).grid(row=1, column=3, padx=(8, 0), pady=(8, 0))
        path_panel.columnconfigure(1, weight=1)

        ttk.Separator(root).pack(fill="x", pady=14)

        filter_panel = ttk.Frame(root)
        filter_panel.pack(fill="both", expand=True)
        for idx, key in enumerate(FILTER_COLUMNS):
            on_change = self.update_account_options if key == "sold_to" else None
            box = FilterBox(filter_panel, FILTER_LABELS[key], on_change=on_change)
            box.grid(row=idx // 3, column=idx % 3, sticky="nsew", padx=6, pady=6)
            self.filter_boxes[key] = box
        for col in range(3):
            filter_panel.columnconfigure(col, weight=1)
        for row in range(2):
            filter_panel.rowconfigure(row, weight=1)

        ttk.Label(
            root,
            text="No selection in a filter means all values. Current-run files overwrite matching existing files; unselected accounts remain unchanged.",
        ).pack(anchor="w", pady=(8, 6))

        log_panel = ttk.Frame(root)
        log_panel.pack(fill="both", expand=True)
        self.log = tk.Text(log_panel, height=10, wrap="word", state="disabled")
        log_scroll = ttk.Scrollbar(log_panel, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=log_scroll.set)
        self.log.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")

        ttk.Label(root, textvariable=self.status).pack(fill="x", pady=(8, 0))

    def choose_source(self) -> None:
        path = filedialog.askopenfilename(
            title="Select OBR_Output.xlsx",
            filetypes=[("Excel workbooks", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.source_path.set(path)

    def choose_output(self) -> None:
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self.output_path.set(path)

    def append_log(self, message: str) -> None:
        self.log.configure(state="normal")
        self.log.insert(tk.END, message + "\n")
        self.log.see(tk.END)
        self.log.configure(state="disabled")
        self.status.set(message)

    def load_filters(self) -> None:
        source = self.source_path.get().strip()
        if not source:
            messagebox.showerror("Missing base file", "Choose OBR_Output.xlsx first.")
            return
        try:
            metadata = get_filter_metadata(source)
        except Exception as exc:
            messagebox.showerror("Could not read workbook", str(exc))
            return

        self.filter_options = metadata["options"]
        self.accounts_by_sold_to = metadata["accounts_by_sold_to"]
        for key, box in self.filter_boxes.items():
            box.set_values(self.filter_options.get(key, []))
        self.update_account_options(preserve_selection=False)
        self.append_log("Loaded filter values from workbook.")

    def update_account_options(self, preserve_selection: bool = True) -> None:
        if not self.filter_options:
            return

        sold_to_box = self.filter_boxes["sold_to"]
        account_box = self.filter_boxes["d_account"]
        selected_sold_tos = set(sold_to_box.selected_values())
        all_sold_tos = {value for value, _ in self.filter_options.get("sold_to", [])}
        current_accounts = set(account_box.selected_values()) if preserve_selection else set()

        if not selected_sold_tos or selected_sold_tos == all_sold_tos:
            allowed_accounts = None
        else:
            allowed_accounts = set()
            for sold_to in selected_sold_tos:
                allowed_accounts.update(self.accounts_by_sold_to.get(sold_to, []))

        account_options = self.filter_options.get("d_account", [])
        if allowed_accounts is not None:
            account_options = [
                option for option in account_options if option[0] in allowed_accounts
            ]

        account_box.set_values(
            account_options,
            selected=current_accounts & {value for value, _ in account_options},
        )

    def selected_filters(self) -> dict[str, list[str]]:
        return {key: box.selected_values() for key, box in self.filter_boxes.items()}

    def generate(self) -> None:
        source = self.source_path.get().strip()
        output = self.output_path.get().strip()
        if not source:
            messagebox.showerror("Missing base file", "Choose OBR_Output.xlsx first.")
            return
        if not output:
            messagebox.showerror("Missing output folder", "Choose an output folder first.")
            return
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("Generation running", "The current run is still in progress.")
            return

        filters = self.selected_filters()
        self.append_log("Starting generation...")
        self.worker = threading.Thread(
            target=self._run_generate,
            args=(source, output, filters),
            daemon=True,
        )
        self.worker.start()

    def _run_generate(self, source: str, output: str, filters: dict[str, list[str]]) -> None:
        try:
            counts = generate_reports(source, output, filters, self._queue_log)
        except Exception as exc:
            self.messages.put(("error", exc))
        else:
            self.messages.put(("done", counts))

    def _queue_log(self, message: str) -> None:
        self.messages.put(("log", message))

    def _drain_messages(self) -> None:
        while True:
            try:
                kind, payload = self.messages.get_nowait()
            except queue.Empty:
                break
            if kind == "log":
                self.append_log(str(payload))
            elif kind == "done":
                self.append_log("Generation complete.")
                summary = "\n".join(f"{name}: {count} rows" for name, count in payload.items())
                messagebox.showinfo("Generation complete", summary)
            elif kind == "error":
                self.append_log(f"Error: {payload}")
                messagebox.showerror("Generation failed", str(payload))
        self.after(100, self._drain_messages)


def run_cli(args: argparse.Namespace) -> None:
    selected_filters = {
        "season": args.season or [],
        "sold_to": args.sold_to or [],
        "d_account": args.d_account or [],
        "document_type": args.document_type or [],
        "distribution_method": args.distribution_method or [],
    }
    counts = generate_reports(args.input, args.output, selected_filters, print)
    for report, count in counts.items():
        print(f"{report}: {count} rows")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate account-level iOBR extraction files.")
    parser.add_argument("--input", type=Path, help="Path to OBR_Output.xlsx")
    parser.add_argument("--output", type=Path, help="Output folder")
    parser.add_argument("--season", action="append", help="Season filter value")
    parser.add_argument("--sold-to", action="append", help="SOLD_TO_CUSTOMER_NBR filter value")
    parser.add_argument("--d-account", action="append", help="D_ACCOUNT filter value")
    parser.add_argument("--document-type", action="append", help="SO_DOCUMENT_TYPE_GROUP_DESC filter value")
    parser.add_argument("--distribution-method", action="append", help="DISTRIBUTION_METHOD_CD filter value")
    return parser.parse_args()


if __name__ == "__main__":
    parsed = parse_args()
    if parsed.input and parsed.output:
        run_cli(parsed)
    else:
        app = IobrApp()
        app.mainloop()

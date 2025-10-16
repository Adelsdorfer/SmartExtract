"""Desktop GUI application to browse Excel files and view 'Noti Text' content."""

from __future__ import annotations

import threading
import tkinter as tk
import tkinter.font as tkfont
from collections.abc import Iterable
from pathlib import Path
from typing import List, Optional
import warnings

import pandas as pd
import requests
from ollama import Client
from tkinter import filedialog, messagebox, ttk

# Suppress harmless openpyxl warning about missing default style.
warnings.filterwarnings(
    "ignore",
    message="Workbook contains no default style, apply openpyxl's default",
    category=UserWarning,
    module="openpyxl",
)

# Modell
MODEL = "granite4:tiny-h"
MODEL_OPTIONS = [
    "granite4:tiny-h",
    "qwen3:8b",
    "gemma3:4b",
    "deepseek-r1:8b",
    "phi4-mini-reasoning",
    "gpt-oss:20b",
]
INITIAL_SERVER_CHECK_DELAY_MS = 4000
AUTO_SERVER_CHECK_INTERVAL_MS = 8000
PROMPT_PRESETS = {
    "Technical Writer": (
        "Act as a technical writer for CT service reports. Summarize the following report in clear, "
        "human-readable English, structured as: 1) Summary of Issue, 2) Timeline with Key Actions, "
        "3) Technical Details, 4) Correspondence, 5) Final Outcome. Keep all log entries (e.g., "
        "timestamped technical messages) exactly as written (1:1, no edits)."
    ),
    "Concise Highlights": (
        "Summarize the following CT service report in three concise bullet points covering the primary issue, "
        "key actions taken, and the final outcome."
    ),
    "Root Cause Focus": (
        "Identify the root cause, diagnostic steps, replaced parts, and final resolution from the following "
        "CT service report. Present the answer as numbered sections with short explanations."
    ),
    "Customer Update": (
        "Create a customer-facing update (maximum five sentences) summarizing the work performed, current "
        "system status, and any follow-up required from the following CT service report."
    ),
}


class MainApp(tk.Tk):
    """Tkinter application for inspecting Excel files and previewing 'Noti Text' values."""

    def __init__(self) -> None:
        super().__init__()
        self.title("SmartExtract | Roland Emrich")
        self.geometry("1900x1200")

        self._current_df: Optional[pd.DataFrame] = None
        self._current_columns: List[str] = []
        self._excel_paths: List[Path] = []
        self._current_directory: Path = Path(r"C:\SmartExtract")
        self._server_reachable: bool = False

        self.selected_file = tk.StringVar()
        self.status_message = tk.StringVar()
        self.current_directory_var = tk.StringVar(value=str(self._current_directory))
        self.api_host_var = tk.StringVar(value="http://md3fgqdc:11434")
        self.server_status_var = tk.StringVar(value="Server status: unknown")
        self.prompt_var = tk.StringVar(value=PROMPT_PRESETS["Technical Writer"])
        default_prompt_choice = next(
            (name for name, text in PROMPT_PRESETS.items() if text == self.prompt_var.get()), "Custom"
        )
        self.prompt_choice_var = tk.StringVar(value=default_prompt_choice)

        self.model_var = tk.StringVar(value=MODEL)
        self.temperature_var = tk.DoubleVar(value=0.5)
        self.current_noti_text = ""
        self._configured_api_host: Optional[str] = None
        self._current_file_path: Optional[Path] = None
        self._batch_in_progress: bool = False
        self._server_check_in_progress: bool = False
        self._automatic_server_check: bool = False
        self._auto_check_after_id: Optional[str] = None
        self._updating_prompt_text: bool = False

        self.model_combobox: Optional[ttk.Combobox] = None
        self.summarize_button: Optional[ttk.Button] = None
        self.batch_process_button: Optional[ttk.Button] = None
        self.check_server_button: Optional[ttk.Button] = None
        self.server_status_label: Optional[ttk.Label] = None
        self.summary_text_widget: Optional[tk.Text] = None

        self.file_combobox: Optional[ttk.Combobox] = None
        self.tree: Optional[ttk.Treeview] = None
        self.noti_text: Optional[tk.Text] = None
        self.prompt_text_widget: Optional[tk.Text] = None
        self.prompt_preset_combobox: Optional[ttk.Combobox] = None

        self._tree_font = tkfont.nametofont("TkDefaultFont")

        self._configure_style()
        self._build_ui()
        self._populate_excel_list()
        self.update_summary_output("")
        self.after(INITIAL_SERVER_CHECK_DELAY_MS, self._delayed_initial_server_check)

    def _configure_style(self) -> None:
        """Configure ttk styling for the application."""
        style = ttk.Style(self)
        style.configure("Toolbar.TFrame", padding=8)
        style.configure("Content.TFrame", padding=8)
        style.configure("Treeview", rowheight=24)
        style.map("Treeview", background=[("selected", "#2a62a8")], foreground=[("selected", "white")])
        style.configure("Status.TLabel", foreground="firebrick")
        style.configure("Noti.TLabelframe", padding=8)
        style.configure("Noti.TLabelframe.Label", font=("Segoe UI", 10, "bold"))

    def _build_ui(self) -> None:
        """Create and lay out the application's widgets."""
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        toolbar = ttk.Frame(self, style="Toolbar.TFrame")
        toolbar.grid(row=0, column=0, sticky="ew")
        toolbar.columnconfigure(1, weight=1)
        toolbar.columnconfigure(5, weight=1)

        ttk.Label(toolbar, text="Folder:").grid(row=0, column=0, padx=(0, 8), sticky="w")
        ttk.Label(toolbar, textvariable=self.current_directory_var).grid(row=0, column=1, sticky="ew")

        choose_button = ttk.Button(toolbar, text="Choose...", command=self.choose_directory)
        choose_button.grid(row=0, column=2, padx=(8, 0))

        ttk.Label(toolbar, text="API host:").grid(row=0, column=3, padx=(16, 6), sticky="e")
        api_entry = ttk.Entry(toolbar, textvariable=self.api_host_var, width=28)
        api_entry.grid(row=0, column=4, sticky="ew")

        self.check_server_button = ttk.Button(toolbar, text="Check Server", command=self.check_server_status)
        self.check_server_button.grid(row=0, column=5, padx=(8, 0))

        self.server_status_label = ttk.Label(toolbar, textvariable=self.server_status_var)
        self.server_status_label.grid(row=0, column=6, padx=(8, 0), sticky="w")

        ttk.Separator(toolbar, orient="horizontal").grid(row=1, column=0, columnspan=7, sticky="ew", pady=4)

        ttk.Label(toolbar, text="Excel file:").grid(row=2, column=0, padx=(0, 8))

        self.file_combobox = ttk.Combobox(toolbar, textvariable=self.selected_file, state="readonly", width=50)
        self.file_combobox.grid(row=2, column=1, sticky="ew", padx=(0, 8))

        reload_button = ttk.Button(toolbar, text="Reload", command=self._populate_excel_list)
        reload_button.grid(row=2, column=2, padx=(0, 8))

        load_button = ttk.Button(toolbar, text="Load", command=self.load_selected_file)
        load_button.grid(row=2, column=3, sticky="w")

        content = ttk.Frame(self, style="Content.TFrame")
        content.grid(row=1, column=0, sticky="nsew")
        content.columnconfigure(0, weight=1)
        content.columnconfigure(1, weight=1)
        content.rowconfigure(0, weight=3)
        content.rowconfigure(1, weight=2)

        tree_frame = ttk.Frame(content)
        tree_frame.grid(row=0, column=0, columnspan=2, sticky="nsew", pady=(0, 8))
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(tree_frame, show="headings", selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree.bind("<<TreeviewSelect>>", self.on_row_select)

        tree_y_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        tree_y_scroll.grid(row=0, column=1, sticky="ns")
        tree_x_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        tree_x_scroll.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=tree_y_scroll.set, xscrollcommand=tree_x_scroll.set)

        noti_frame = ttk.Labelframe(content, text="Noti Text", style="Noti.TLabelframe")
        noti_frame.grid(row=1, column=0, sticky="nsew", padx=(0, 8))
        noti_frame.columnconfigure(1, weight=1)
        noti_frame.columnconfigure(3, weight=1)

        ttk.Label(noti_frame, textvariable=self.status_message, style="Status.TLabel").grid(
            row=0, column=0, columnspan=6, sticky="w", pady=(0, 4)
        )

        ttk.Label(noti_frame, text="Preset:").grid(row=1, column=0, sticky="w", padx=(0, 4))
        preset_values = ["Custom"] + list(PROMPT_PRESETS.keys())
        self.prompt_preset_combobox = ttk.Combobox(
            noti_frame,
            textvariable=self.prompt_choice_var,
            state="readonly",
            width=40,
            values=preset_values,
        )
        self.prompt_preset_combobox.grid(row=1, column=1, sticky="ew", padx=(0, 8))
        if self.prompt_choice_var.get() not in preset_values:
            self.prompt_choice_var.set("Custom")
        self.prompt_preset_combobox.set(self.prompt_choice_var.get())
        self.prompt_preset_combobox.bind("<<ComboboxSelected>>", self._on_prompt_preset_selected)

        ttk.Label(noti_frame, text="Model:").grid(row=1, column=2, sticky="e", padx=(0, 4))
        self.model_combobox = ttk.Combobox(
            noti_frame,
            textvariable=self.model_var,
            state="readonly",
            width=30,
            values=MODEL_OPTIONS,
        )
        self.model_combobox.grid(row=1, column=3, sticky="ew", padx=(0, 8))
        self.model_combobox.current(MODEL_OPTIONS.index(MODEL))

        ttk.Label(noti_frame, text="Temperature:").grid(row=1, column=4, sticky="e", padx=(0, 4))
        temperature_spinbox = ttk.Spinbox(
            noti_frame,
            textvariable=self.temperature_var,
            from_=0.0,
            to=2.0,
            increment=0.1,
            width=6,
        )
        temperature_spinbox.grid(row=1, column=5, sticky="w")

        self.summarize_button = ttk.Button(noti_frame, text="Process", command=self.summarize_current_text)
        self.summarize_button.grid(row=1, column=6, padx=(8, 0))
        self.summarize_button.configure(state="disabled")

        self.batch_process_button = ttk.Button(
            noti_frame,
            text="Batch Process",
            command=self.batch_process_all_rows,
        )
        self.batch_process_button.grid(row=1, column=7, padx=(8, 0))
        self.batch_process_button.configure(state="disabled")

        ttk.Label(noti_frame, text="Prompt:").grid(row=2, column=0, sticky="nw", padx=(0, 4))
        self.prompt_text_widget = tk.Text(
            noti_frame,
            wrap="word",
            height=2,
            font=("Segoe UI", 10),
            relief="flat",
            padx=6,
            pady=4,
        )
        self.prompt_text_widget.grid(row=2, column=1, columnspan=7, sticky="ew", padx=(0, 8), pady=(4, 0))
        self._set_prompt_text(self.prompt_var.get())
        self.prompt_text_widget.bind("<<Modified>>", self._on_prompt_text_modified)
        self.prompt_text_widget.edit_modified(False)

        noti_text_frame = ttk.Frame(noti_frame)
        noti_text_frame.grid(row=3, column=0, columnspan=8, sticky="nsew", pady=(8, 0))
        noti_frame.rowconfigure(3, weight=1)
        noti_text_frame.columnconfigure(0, weight=1)
        noti_text_frame.rowconfigure(0, weight=1)

        self.noti_text = tk.Text(
            noti_text_frame,
            wrap="word",
            state="disabled",
            font=("Segoe UI", 10),
            relief="flat",
            padx=6,
            pady=6,
        )
        self.noti_text.grid(row=0, column=0, sticky="nsew")

        noti_scroll = ttk.Scrollbar(noti_text_frame, orient="vertical", command=self.noti_text.yview)
        noti_scroll.grid(row=0, column=1, sticky="ns")
        self.noti_text.configure(yscrollcommand=noti_scroll.set)

        summary_frame = ttk.Labelframe(content, text="Output", style="Noti.TLabelframe")
        summary_frame.grid(row=1, column=1, sticky="nsew")
        summary_frame.columnconfigure(0, weight=1)
        summary_frame.rowconfigure(0, weight=1)

        self.summary_text_widget = tk.Text(
            summary_frame,
            wrap="word",
            state="disabled",
            font=("Segoe UI", 10),
            relief="flat",
            padx=6,
            pady=6,
        )
        self.summary_text_widget.grid(row=0, column=0, sticky="nsew")

        summary_scroll = ttk.Scrollbar(summary_frame, orient="vertical", command=self.summary_text_widget.yview)
        summary_scroll.grid(row=0, column=1, sticky="ns")
        self.summary_text_widget.configure(yscrollcommand=summary_scroll.set)

    def _delayed_initial_server_check(self) -> None:
        """Run an automatic server check shortly after startup."""
        if self._auto_check_after_id is not None or self._server_check_in_progress:
            return
        if not self.check_server_button:
            self.after(500, self._delayed_initial_server_check)
            return
        self._trigger_automatic_server_check()

    def choose_directory(self) -> None:
        """Prompt the user to select a directory containing Excel files."""
        directory = filedialog.askdirectory(initialdir=self._current_directory)
        if not directory:
            return

        self._current_directory = Path(directory)
        self.current_directory_var.set(str(self._current_directory))
        self._populate_excel_list()

    def _populate_excel_list(self) -> None:
        """Populate the Excel files combobox with files from the current directory."""
        self.status_message.set("")
        excel_candidates = list(self.scan_directory_for_excels(self._current_directory))
        self._excel_paths = excel_candidates
        self._current_file_path = None

        display_names = [path.name for path in excel_candidates]
        if self.file_combobox:
            self.file_combobox["values"] = display_names
            self.selected_file.set("")

        if not excel_candidates:
            messagebox.showinfo(
                "Excel Files",
                "No Excel files found in the selected directory.",
            )

    @staticmethod
    def scan_directory_for_excels(directory: Path) -> Iterable[Path]:
        """Yield Excel files from the provided directory."""
        if not directory.exists():
            return []

        return (
            path
            for path in sorted(directory.iterdir())
            if path.is_file() and path.suffix.lower() in {".xlsx", ".xls"}
        )

    def load_selected_file(self) -> None:
        """Load the dataframe for the selected Excel file and populate the treeview."""
        if not self.selected_file.get():
            messagebox.showinfo("Load File", "Please select an Excel file to load.")
            return

        selected_name = self.selected_file.get()
        matching = next((path for path in self._excel_paths if path.name == selected_name), None)
        if matching is None:
            messagebox.showerror("Load File", "Selected file was not found on disk.")
            return

        try:
            dataframe = pd.read_excel(matching)
        except Exception as exc:  # pylint: disable=broad-except
            messagebox.showerror("Load File", f"Failed to load Excel file.\n\n{exc}")
            return

        self._current_df = dataframe
        self._current_file_path = matching
        self.populate_treeview(dataframe)
        self.show_noti_text("")
        self.update_summary_output("")
        self.status_message.set(f"Loaded '{matching.name}'.")
        self._refresh_summarize_button_state()

    def populate_treeview(self, dataframe: pd.DataFrame) -> None:
        """Populate the treeview with the provided dataframe."""
        if not self.tree:
            return

        for item in self.tree.get_children():
            self.tree.delete(item)

        self._current_columns = [str(column) for column in dataframe.columns]
        self.tree["columns"] = self._current_columns

        for column in self._current_columns:
            self.tree.heading(column, text=column)

        max_rows = min(len(dataframe), 1000)
        sample = dataframe.head(max_rows).fillna("").astype(str)

        for column in self._current_columns:
            header_width = self._tree_font.measure(str(column)) + 20
            data_width = max((self._tree_font.measure(value) for value in sample[column]), default=header_width)
            column_width = max(header_width, min(data_width + 30, 400))
            self.tree.column(column, width=column_width, anchor="w", stretch=True)

        self.tree.tag_configure("oddrow", background="#f2f2f2")
        self.tree.tag_configure("evenrow", background="white")

        for index, row in enumerate(dataframe.itertuples(index=False, name=None)):
            values = tuple("" if pd.isna(value) else str(value) for value in row)
            tag = "oddrow" if index % 2 else "evenrow"
            self.tree.insert("", "end", values=values, tags=(tag,))

    def on_row_select(self, _: tk.Event) -> None:
        """Display the Noti Text for the selected row if available."""
        if not self.tree or not self._current_columns:
            return

        try:
            noti_index = self._current_columns.index("Noti Text")
        except ValueError:
            self.status_message.set("Column 'Noti Text' not found in current file.")
            self.show_noti_text("")
            return

        selection = self.tree.selection()
        if not selection:
            return

        item_id = selection[0]
        item = self.tree.item(item_id)
        values = item.get("values", [])

        if not values or noti_index >= len(values):
            self.show_noti_text("")
            return

        noti_value = values[noti_index]
        if isinstance(noti_value, str) and noti_value.strip():
            self.status_message.set("")
            self.show_noti_text(noti_value)
        elif pd.isna(noti_value):
            self.show_noti_text("")
        else:
            text_value = str(noti_value).strip()
            if text_value:
                self.status_message.set("")
                self.show_noti_text(text_value)
            else:
                self.show_noti_text("")

    def show_noti_text(self, value: str) -> None:
        """Render the supplied Noti Text in the dedicated text widget."""
        self.current_noti_text = value or ""
        if not self.noti_text:
            return

        self.noti_text.configure(state="normal")
        self.noti_text.delete("1.0", tk.END)
        if value:
            self.noti_text.insert("1.0", value)
        self.noti_text.configure(state="disabled")
        self._refresh_summarize_button_state()

    def update_summary_output(self, text: str) -> None:
        """Replace the content of the summary text widget."""
        if not self.summary_text_widget:
            return

        self.summary_text_widget.configure(state="normal")
        self.summary_text_widget.delete("1.0", tk.END)
        if text:
            self.summary_text_widget.insert("1.0", text)
        self.summary_text_widget.configure(state="disabled")

    def _append_summary_text(self, content: str) -> None:
        """Append streamed content to the summary widget."""
        if not self.summary_text_widget or not content:
            return

        self.summary_text_widget.configure(state="normal")
        self.summary_text_widget.insert(tk.END, content)
        self.summary_text_widget.see(tk.END)
        self.summary_text_widget.configure(state="disabled")

    def _parse_api_host(self, value: str) -> Optional[str]:
        """Normalize the configured API host string."""
        text = value.strip()
        if not text:
            return None

        if text.startswith("http://") or text.startswith("https://"):
            base = text
        else:
            base = f"http://{text}"

        scheme, _, remainder = base.partition("://")
        host_port = remainder.split("/", 1)[0]

        if ":" not in host_port:
            raise ValueError("API host must include a port (e.g. http://host:port).")

        host, port = host_port.rsplit(":", 1)
        if not port.isdigit():
            raise ValueError("API port must be numeric.")

        normalized = f"{scheme}://{host}:{port}"
        return normalized

    def _schedule_next_server_check(self) -> None:
        """Schedule the next automatic server availability check."""
        if self._auto_check_after_id is not None:
            self.after_cancel(self._auto_check_after_id)
        self._auto_check_after_id = self.after(
            AUTO_SERVER_CHECK_INTERVAL_MS, self._trigger_automatic_server_check
        )

    def _trigger_automatic_server_check(self) -> None:
        """Run a server check initiated by the automatic scheduler."""
        self._auto_check_after_id = None
        if self._server_check_in_progress:
            self._schedule_next_server_check()
            return
        self._automatic_server_check = True
        self.check_server_status()

    def _on_server_check_complete(self) -> None:
        """Reset UI state and queue the next automatic server check."""
        self._server_check_in_progress = False
        if self.check_server_button:
            self.check_server_button.configure(state="normal")
        self._schedule_next_server_check()

    def check_server_status(self) -> None:
        """Check whether the configured Ollama server is reachable."""
        automatic = self._automatic_server_check
        self._automatic_server_check = False

        if self._server_check_in_progress:
            if automatic:
                self._schedule_next_server_check()
            return

        if not self.check_server_button:
            return

        host_value = self.api_host_var.get()
        try:
            normalized_host = self._parse_api_host(host_value)
        except ValueError as exc:
            message = f"Invalid host: {exc}"
            if automatic:
                self._update_server_status(message, False)
                self._schedule_next_server_check()
            else:
                messagebox.showerror("Ollama Server", str(exc))
            return

        if normalized_host is None:
            if automatic:
                self._update_server_status("Please provide an API host.", False)
                self._schedule_next_server_check()
            else:
                messagebox.showerror("Ollama Server", "Please provide an API host.")
            return

        self._server_check_in_progress = True
        self.server_status_var.set("Server status: Checking...")
        self.check_server_button.configure(state="disabled")
        threading.Thread(
            target=self._check_server_status_worker,
            args=(normalized_host,),
            daemon=True,
        ).start()

    def _check_server_status_worker(self, host_url: str) -> None:
        """Worker thread that performs the server health check."""
        try:
            response = requests.get(f"{host_url}/api/version", timeout=5)
            if response.status_code == 200:
                self._configured_api_host = host_url
                self._schedule_status_update("Server reachable.", True)
            else:
                self._schedule_status_update("No server response.", False)
        except requests.RequestException:
            self._schedule_status_update("Unable to contact server.", False)
        finally:
            self.after(0, self._on_server_check_complete)

    def _schedule_status_update(self, message: str, reachable: bool) -> None:
        """Schedule a server status update on the Tkinter main thread."""
        self.after(0, lambda: self._update_server_status(message, reachable))

    def _update_server_status(self, message: str, reachable: bool) -> None:
        """Update the UI based on server reachability."""
        self.server_status_var.set(f"Server status: {message}")
        self._server_reachable = reachable
        if not reachable:
            self._configured_api_host = None
        self._refresh_summarize_button_state()

    def _refresh_summarize_button_state(self) -> None:
        """Enable or disable processing buttons based on current application state."""
        can_stream = (
            self._server_reachable
            and bool(self.current_noti_text.strip())
            and not self._batch_in_progress
        )
        if self.summarize_button:
            self.summarize_button.configure(state="normal" if can_stream else "disabled")

        can_batch = False
        if (
            self._server_reachable
            and not self._batch_in_progress
            and self._current_df is not None
            and not self._current_df.empty
        ):
            column_names = [str(column) for column in self._current_df.columns]
            can_batch = "Noti Text" in column_names

        if self.batch_process_button:
            self.batch_process_button.configure(state="normal" if can_batch else "disabled")

    def _set_prompt_text(self, value: str) -> None:
        """Synchronize the prompt text widget and backing variable."""
        self.prompt_var.set(value)
        if not self.prompt_text_widget:
            return
        self._updating_prompt_text = True
        self.prompt_text_widget.delete("1.0", tk.END)
        self.prompt_text_widget.insert("1.0", value)
        self.prompt_text_widget.edit_modified(False)
        self._updating_prompt_text = False

    def _get_prompt_text(self) -> str:
        """Return the current prompt text."""
        if self.prompt_text_widget:
            text = self.prompt_text_widget.get("1.0", "end-1c").strip()
            self.prompt_var.set(text)
            return text
        return self.prompt_var.get().strip()

    def _on_prompt_preset_selected(self, _event: tk.Event) -> None:
        """Update the prompt field when a preset is chosen."""
        selection = self.prompt_choice_var.get()
        if selection == "Custom":
            return
        preset_text = PROMPT_PRESETS.get(selection, "")
        self._set_prompt_text(preset_text)

    def _on_prompt_text_modified(self, _event: tk.Event) -> None:
        """Track manual prompt edits and reflect them in the preset selector."""
        if not self.prompt_text_widget or not self.prompt_text_widget.edit_modified():
            return
        if self._updating_prompt_text:
            self.prompt_text_widget.edit_modified(False)
            return

        current_text = self.prompt_text_widget.get("1.0", "end-1c").strip()
        self.prompt_var.set(current_text)
        matching_choice = next(
            (name for name, text in PROMPT_PRESETS.items() if text == current_text),
            None,
        )
        self.prompt_choice_var.set(matching_choice or "Custom")
        self.prompt_text_widget.edit_modified(False)

    def _get_temperature(self) -> float:
        """Clamp and round the temperature value."""
        try:
            value = float(self.temperature_var.get())
        except (TypeError, ValueError):
            value = 0.5

        value = max(0.0, min(2.0, value))
        return round(value, 1)

    def summarize_current_text(self) -> None:
        """Trigger streaming summarization of the current Noti Text."""
        if not self._server_reachable:
            messagebox.showinfo("Summarize", "Please ensure the server is reachable first.")
            return

        if not self.current_noti_text.strip():
            messagebox.showinfo("Summarize", "No Noti Text is currently loaded.")
            return

        temperature = self._get_temperature()
        user_prompt = self._get_prompt_text()
        combined_prompt = self._compose_prompt(user_prompt, self.current_noti_text)

        self.status_message.set("Processing Input, please wait...")
        self.update_summary_output("")
        if self.summarize_button:
            self.summarize_button.configure(state="disabled")

        selected_model = self.model_var.get().strip() or MODEL

        threading.Thread(
            target=self._run_ollama_stream,
            args=(combined_prompt, temperature, selected_model),
            daemon=True,
        ).start()

    def _run_ollama_stream(self, combined_prompt: str, temperature: float, model_name: str) -> None:
        """Stream the Ollama response and update the UI incrementally."""
        try:
            host_value = self._configured_api_host or self._parse_api_host(self.api_host_var.get())
            if not host_value:
                raise ValueError("No API host configured.")
        except ValueError as exc:
            self._handle_summarization_error(str(exc))
            return

        try:
            client = Client(host=host_value)
            stream = client.chat(
                model=model_name or MODEL,
                messages=[{"role": "user", "content": combined_prompt}],
                stream=True,
                options={"temperature": temperature},
            )

            for chunk in stream:
                message = chunk.get("message")
                if message and "content" in message:
                    content = message["content"]
                    self.after(0, self._append_summary_text, content)

            self.after(0, self._on_summarization_success)
        except Exception as exc:  # pylint: disable=broad-except
            self._handle_summarization_error(str(exc))

    def _on_summarization_success(self) -> None:
        """Handle successful completion of summarization."""
        self.status_message.set("Processing complete.")
        if self.summarize_button:
            self.summarize_button.configure(state="normal")
        self._refresh_summarize_button_state()

    def _handle_summarization_error(self, message: str) -> None:
        """Handle summarization failures on the main thread."""
        self.after(
            0,
            lambda: self._on_summarization_error_main_thread(message),
        )

    def _on_summarization_error_main_thread(self, message: str) -> None:
        """Display summarization errors with UI feedback."""
        if self.summarize_button:
            self.summarize_button.configure(state="normal")
        self.status_message.set(f"Prozessing failed: {message}")
        self.update_summary_output("")
        messagebox.showerror("Summarize", f"Prozessing failed: {message}")
        self._refresh_summarize_button_state()

    def batch_process_all_rows(self) -> None:
        """Run batch processing for every Noti Text entry in the current dataframe."""
        if self._batch_in_progress:
            return

        if not self._server_reachable:
            messagebox.showinfo("Batch Processing", "Please ensure the server is reachable first.")
            return

        if self._current_df is None or self._current_df.empty:
            messagebox.showinfo("Batch Processing", "No Excel file is currently loaded.")
            return

        noti_column_name = "Noti Text"
        if not any(str(column) == noti_column_name for column in self._current_df.columns):
            messagebox.showerror("Batch Processing", "Column 'Noti Text' not found in the loaded file.")
            return

        if self._current_file_path is None:
            messagebox.showerror("Batch Processing", "Unable to determine the source file path.")
            return

        try:
            host_value = self._configured_api_host or self._parse_api_host(self.api_host_var.get())
            if not host_value:
                raise ValueError("No API host configured.")
        except ValueError as exc:
            messagebox.showerror("Batch Processing", str(exc))
            return

        output_path = self._generate_batch_output_path(self._current_file_path)
        model_name = self.model_var.get().strip() or MODEL
        temperature = self._get_temperature()
        user_prompt = self._get_prompt_text()
        dataframe_copy = self._current_df.copy(deep=True)

        self._batch_in_progress = True
        self.status_message.set("Batch processing started...")
        if self.batch_process_button:
            self.batch_process_button.configure(state="disabled")
        if self.summarize_button:
            self.summarize_button.configure(state="disabled")

        threading.Thread(
            target=self._run_batch_processing,
            args=(
                dataframe_copy,
                host_value,
                model_name,
                temperature,
                user_prompt,
                output_path,
            ),
            daemon=True,
        ).start()

    def _run_batch_processing(
        self,
        dataframe: pd.DataFrame,
        host_value: str,
        model_name: str,
        temperature: float,
        user_prompt: str,
        output_path: Path,
    ) -> None:
        """Execute the batch processing workflow off the main thread."""
        try:
            client = Client(host=host_value)
            noti_column_name = "Noti Text"
            noti_index = next(
                (idx for idx, column in enumerate(dataframe.columns) if str(column) == noti_column_name),
                None,
            )
            if noti_index is None:
                raise ValueError("Column 'Noti Text' not found in the loaded file.")

            total_rows = len(dataframe.index)
            gpt_outputs: List[str] = []

            for row_number, (_, row) in enumerate(dataframe.iterrows(), start=1):
                self._update_status_async(f"Processing row {row_number}/{total_rows}...")
                noti_value_raw = row.iloc[noti_index]
                noti_text = "" if pd.isna(noti_value_raw) else str(noti_value_raw)

                if not noti_text.strip():
                    gpt_outputs.append("")
                    self._update_status_async(f"Row {row_number}/{total_rows}: skipped (empty Noti Text)")
                    continue

                combined_prompt = self._compose_prompt(user_prompt, noti_text)
                response_chunks: List[str] = []

                stream = client.chat(
                    model=model_name or MODEL,
                    messages=[{"role": "user", "content": combined_prompt}],
                    stream=True,
                    options={"temperature": temperature},
                )

                for chunk in stream:
                    message = chunk.get("message")
                    if message and "content" in message:
                        response_chunks.append(message["content"])

                gpt_outputs.append("".join(response_chunks).strip())
                self._update_status_async(f"Row {row_number}/{total_rows}: processed")

            processed_df = dataframe.copy()
            columns_to_drop = [column for column in processed_df.columns if str(column) == "GPT Output"]
            if columns_to_drop:
                processed_df = processed_df.drop(columns=columns_to_drop)

            processed_df.insert(noti_index + 1, "GPT Output", gpt_outputs)
            self._update_status_async("Saving results...")
            processed_df.to_excel(output_path, index=False)
            self.after(0, lambda: self._on_batch_processing_success(output_path))
        except Exception as exc:  # pylint: disable=broad-except
            self._handle_batch_processing_error(str(exc))

    def _on_batch_processing_success(self, output_path: Path) -> None:
        """Handle completion of batch processing."""
        self._batch_in_progress = False
        self.status_message.set(f"Batch processing complete: {output_path.name}")
        self._refresh_summarize_button_state()
        messagebox.showinfo("Batch Processing", f"Batch processing complete.\n\nSaved to:\n{output_path}")

    def _handle_batch_processing_error(self, message: str) -> None:
        """Schedule error handling for batch processing on the main thread."""
        self.after(0, lambda: self._on_batch_processing_error_main_thread(message))

    def _on_batch_processing_error_main_thread(self, message: str) -> None:
        """Display batch processing errors and restore UI state."""
        self._batch_in_progress = False
        self.status_message.set(f"Batch processing failed: {message}")
        messagebox.showerror("Batch Processing", f"Batch processing failed.\n\n{message}")
        self._refresh_summarize_button_state()

    def _update_status_async(self, message: str) -> None:
        """Update the status label from worker threads."""
        self.after(0, lambda: self.status_message.set(message))

    @staticmethod
    def _compose_prompt(base_prompt: str, noti_text: str) -> str:
        """Combine the user prompt with the current Noti Text snippet."""
        prompt_parts: List[str] = []
        if base_prompt:
            prompt_parts.append(base_prompt)
        prompt_parts.append(f"Noti Text:\n{noti_text}")
        return "\n\n".join(prompt_parts)

    @staticmethod
    def _generate_batch_output_path(source_path: Path) -> Path:
        """Create a unique output path for batch processing results."""
        base_candidate = source_path.with_name(f"{source_path.stem}_gpt_output{source_path.suffix}")
        if not base_candidate.exists():
            return base_candidate

        counter = 1
        while True:
            candidate = source_path.with_name(f"{source_path.stem}_gpt_output_{counter}{source_path.suffix}")
            if not candidate.exists():
                return candidate
            counter += 1


if __name__ == "__main__":
    try:
        MainApp().mainloop()
    except Exception as error:  # pylint: disable=broad-except
        messagebox.showerror("Application Error", f"An unexpected error occurred.\n\n{error}")

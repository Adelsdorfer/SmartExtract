"""Desktop GUI application to browse Excel files and view 'Noti Text' content."""

from __future__ import annotations

import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox
from tkinter import ttk
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from collections.abc import Iterable
import threading
import warnings
from urllib.parse import urlparse

import pandas as pd

# Suppress harmless openpyxl warning about missing default style.
warnings.filterwarnings(
    "ignore",
    message="Workbook contains no default style, apply openpyxl's default",
    category=UserWarning,
    module="openpyxl",
)


class MainApp(tk.Tk):
    """Tkinter application for inspecting Excel files and previewing 'Noti Text' values."""

    WINDOW_SIZE = "1900x1200"

    def __init__(self) -> None:
        super().__init__()
        self.title("Excel Noti Viewer")
        self.geometry(self.WINDOW_SIZE)

        self._current_df: Optional[pd.DataFrame] = None
        self._current_columns: List[str] = []
        self._excel_paths: List[Path] = []
        self._current_directory: Path = Path(r"e:\Notiscan")

        self.selected_file = tk.StringVar()
        self.status_message = tk.StringVar()
        self.current_directory_var = tk.StringVar(value=str(self._current_directory))
        self.api_host_var = tk.StringVar(value="http://192.168.114.1:1234")
        self.server_status_var = tk.StringVar(value="Server status: unknown")
        self.prompt_var = tk.StringVar()
        self.model_var = tk.StringVar()
        self.temperature_var = tk.DoubleVar(value=0.5)
        self._model_option_map: Dict[str, Optional[str]] = {}
        self._available_models: List[Tuple[str, Optional[str]]] = []
        self.current_noti_text = ""
        self._ollama_available = True
        self._configured_api_host: Optional[str] = None
        self._ollama_client: Optional[Any] = None

        self.model_combobox: Optional[ttk.Combobox] = None
        self.summarize_button: Optional[ttk.Button] = None
        self.check_server_button: Optional[ttk.Button] = None
        self.server_status_label: Optional[ttk.Label] = None
        self.summary_text_widget: Optional[tk.Text] = None

        self._tree_font = tkfont.nametofont("TkDefaultFont")

        self._configure_style()
        self._build_ui()
        self._populate_excel_list()
        self.update_summary_output("")
        self.after(100, self.refresh_llm_models)

    def _configure_style(self) -> None:
        """Set up ttk styles."""
        style = ttk.Style(self)
        style.configure("Toolbar.TFrame", padding=8)
        style.configure("Content.TFrame", padding=8)
        style.configure("Treeview", rowheight=24)
        style.map("Treeview", background=[("selected", "#2a62a8")], foreground=[("selected", "white")])
        style.configure("Status.TLabel", foreground="firebrick")
        style.configure("Noti.TLabelframe", padding=8)
        style.configure("Noti.TLabelframe.Label", font=("Segoe UI", 10, "bold"))

        # Configure zebra striping tags later in populate_treeview.

    def _build_ui(self) -> None:
        """Create and lay out widgets."""
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
        tree_frame.grid(row=0, column=0, columnspan=2, sticky="nsew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(tree_frame, show="headings", selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree.bind("<<TreeviewSelect>>", self.on_row_select)

        tree_vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        tree_hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=tree_vsb.set, xscrollcommand=tree_hsb.set)
        tree_vsb.grid(row=0, column=1, sticky="ns")
        tree_hsb.grid(row=1, column=0, sticky="ew")

        noti_frame = ttk.Labelframe(content, text="Noti Text", style="Noti.TLabelframe")
        noti_frame.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
        noti_frame.columnconfigure(0, weight=1)
        noti_frame.rowconfigure(2, weight=1)

        self.status_label = ttk.Label(noti_frame, textvariable=self.status_message, style="Status.TLabel")
        self.status_label.grid(row=0, column=0, sticky="w", pady=(0, 4))

        controls_frame = ttk.Frame(noti_frame)
        controls_frame.grid(row=1, column=0, sticky="ew", pady=(0, 4))
        controls_frame.columnconfigure(1, weight=1)
        controls_frame.columnconfigure(3, weight=1)

        ttk.Label(controls_frame, text="Prompt:").grid(row=0, column=0, padx=(0, 6), sticky="w")
        prompt_entry = ttk.Entry(controls_frame, textvariable=self.prompt_var, width=40)
        prompt_entry.grid(row=0, column=1, sticky="ew", padx=(0, 10))

        ttk.Label(controls_frame, text="Model:").grid(row=0, column=2, padx=(0, 6), sticky="w")
        self.model_combobox = ttk.Combobox(controls_frame, textvariable=self.model_var, state="readonly", width=30)
        self.model_combobox.grid(row=0, column=3, sticky="ew", padx=(0, 10))

        ttk.Label(controls_frame, text="Temperature:").grid(row=0, column=4, padx=(0, 6), sticky="w")
        temperature_spin = ttk.Spinbox(
            controls_frame,
            textvariable=self.temperature_var,
            from_=0.0,
            to=2.0,
            increment=0.1,
            width=6,
            format="%.1f",
        )
        temperature_spin.grid(row=0, column=5, padx=(0, 10), sticky="w")

        self.summarize_button = ttk.Button(
            controls_frame,
            text="Summarize",
            command=self.summarize_current_text,
            state="disabled",
        )
        self.summarize_button.grid(row=0, column=6, sticky="ew")

        text_frame = ttk.Frame(noti_frame)
        text_frame.grid(row=2, column=0, sticky="nsew")
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)

        self.noti_text = tk.Text(
            text_frame,
            wrap="word",
            state="disabled",
            font=("Segoe UI", 10),
            relief="flat",
            padx=6,
            pady=6,
        )
        self.noti_text.grid(row=0, column=0, sticky="nsew")

        noti_vsb = ttk.Scrollbar(text_frame, orient="vertical", command=self.noti_text.yview)
        self.noti_text.configure(yscrollcommand=noti_vsb.set)
        noti_vsb.grid(row=0, column=1, sticky="ns")

        summary_frame = ttk.Labelframe(content, text="LLM Summary", style="Noti.TLabelframe")
        summary_frame.grid(row=1, column=1, sticky="nsew", pady=(8, 0), padx=(8, 0))
        summary_frame.columnconfigure(0, weight=1)
        summary_frame.rowconfigure(0, weight=1)

        summary_container = ttk.Frame(summary_frame)
        summary_container.grid(row=0, column=0, sticky="nsew")
        summary_container.columnconfigure(0, weight=1)
        summary_container.rowconfigure(0, weight=1)

        self.summary_text_widget = tk.Text(
            summary_container,
            wrap="word",
            state="disabled",
            font=("Segoe UI", 10),
            relief="flat",
            padx=6,
            pady=6,
        )
        self.summary_text_widget.grid(row=0, column=0, sticky="nsew")

        summary_vsb = ttk.Scrollbar(summary_container, orient="vertical", command=self.summary_text_widget.yview)
        self.summary_text_widget.configure(yscrollcommand=summary_vsb.set)
        summary_vsb.grid(row=0, column=1, sticky="ns")

    def _populate_excel_list(self) -> None:
        """Scan the selected directory for Excel files and populate the dropdown."""
        self.status_message.set("")
        self._excel_paths = sorted(self.scan_directory_for_excels(self._current_directory))
        filenames = [path.name for path in self._excel_paths]
        self.file_combobox["values"] = filenames

        if filenames:
            previous_selection = self.selected_file.get()
            if previous_selection in filenames:
                index = filenames.index(previous_selection)
                self.file_combobox.current(index)
                self.selected_file.set(previous_selection)
            else:
                self.file_combobox.current(0)
                self.selected_file.set(filenames[0])
        else:
            self.file_combobox.set("")
            self.selected_file.set("")
            self.clear_treeview()
            self.show_noti_text("")
            self.update_summary_output("")
            messagebox.showinfo(
                "No Excel Files",
                f"No Excel files (.xlsx, .xls) were found in\n{self._current_directory}",
            )

    def refresh_llm_models(self) -> None:
        """Fetch available LLM models in a background thread."""
        if not self.model_combobox:
            return

        self.model_var.set("Loading...")
        self.model_combobox.configure(state="disabled")
        if self.summarize_button:
            self.summarize_button.configure(state="disabled")

        thread = threading.Thread(target=self._load_llm_models_worker, daemon=True)
        thread.start()

    def _parse_api_host(self) -> Tuple[Optional[str], Optional[str]]:
        """Return a normalized Ollama host URL or an error message."""
        raw = self.api_host_var.get().strip()
        if not raw:
            return None, "API host is empty."

        if "://" not in raw:
            raw = f"http://{raw}"

        try:
            parsed = urlparse(raw)
        except Exception:  # noqa: broad-except
            return None, "API host is not a valid URL."

        if not parsed.scheme or not parsed.hostname:
            return None, "API host must include a scheme and hostname."

        port = parsed.port or 11434

        normalized = f"{parsed.scheme}://{parsed.hostname}:{port}"
        return normalized, None

    def _create_ollama_client(self, host: str) -> Tuple[Optional[Any], Optional[str]]:
        """Return a cached Ollama client for the given host or an error message."""
        try:
            import ollama  # type: ignore
        except ImportError:
            self._ollama_available = False
            self._ollama_client = None
            return None, "ollama package is not installed."

        self._ollama_available = True

        if self._ollama_client is not None and self._configured_api_host == host:
            return self._ollama_client, None

        try:
            client = ollama.Client(host=host)
        except Exception as exc:  # noqa: broad-except
            self._ollama_client = None
            return None, f"Failed to connect to Ollama host: {exc}"

        self._ollama_client = client
        self._configured_api_host = host
        return client, None

    def check_server_status(self) -> None:
        """Trigger a server connectivity check."""
        host_url, error = self._parse_api_host()
        if error:
            self.server_status_var.set(f"Server status: {error}")
            return

        if self.check_server_button:
            self.check_server_button.configure(state="disabled")
        self.server_status_var.set("Server status: checking...")

        thread = threading.Thread(
            target=self._check_server_status_worker,
            args=(host_url,),
            daemon=True,
        )
        thread.start()

    def _check_server_status_worker(self, host_url: str) -> None:
        """Background worker to validate Ollama server availability."""
        client, error = self._create_ollama_client(host_url)
        if error:
            self.after(0, lambda: self._update_server_status(error, reachable=False))
            return

        try:
            response = client.list()
            models = response.get("models") if isinstance(response, dict) else None
        except Exception as exc:  # noqa: broad-except
            self.after(0, lambda: self._update_server_status(f"Check failed: {exc}", reachable=False))
            return

        # Force refetch on success to keep client fresh.
        self._configured_api_host = host_url
        self.after(0, lambda: self._update_server_status("Server reachable.", reachable=True))
        self.after(0, self.refresh_llm_models)

    def _update_server_status(self, message: str, reachable: bool) -> None:
        """Update server status label and button state on the UI thread."""
        prefix = "Server status: "
        self.server_status_var.set(prefix + message)
        if self.check_server_button:
            self.check_server_button.configure(state="normal")
        if reachable:
            # Keep last configured host in sync with parsed host.
            host_port, _ = self._parse_api_host()
            if host_port:
                self._configured_api_host = host_port
        else:
            self._ollama_client = None

    def _load_llm_models_worker(self) -> None:
        """Background worker to query the Ollama server for available models."""
        host_url, error = self._parse_api_host()
        if error:
            self.after(0, lambda: self._finalize_model_list([], error))
            return

        client, client_error = self._create_ollama_client(host_url)
        if client_error:
            self.after(0, lambda: self._finalize_model_list([], client_error))
            return

        try:
            response = client.list()
        except Exception as exc:  # noqa: broad-except
            error_message = f"Failed to list models: {exc}"
            self.after(0, lambda: self._finalize_model_list([], error_message))
            return

        models: List[Tuple[str, Optional[str]]] = []
        model_entries = response.get("models") if isinstance(response, dict) else None
        if isinstance(model_entries, Iterable):
            for entry in model_entries:
                if not isinstance(entry, dict):
                    continue
                name = entry.get("name") or entry.get("model")
                if not name:
                    continue
                display = str(name)
                details = entry.get("details")
                if isinstance(details, dict):
                    size = details.get("parameter_size")
                    quant = details.get("quantization_level")
                    extras = " ".join(str(part) for part in (size, quant) if part)
                    if extras:
                        display = f"{display} ({extras})"
                models.append((display, str(name)))

        # Deduplicate while preserving order.
        unique_models: List[Tuple[str, Optional[str]]] = []
        seen_keys: set[str] = set()
        for display, key in models:
            identifier = key if key is not None else "__default__"
            if identifier in seen_keys:
                continue
            seen_keys.add(identifier)
            unique_models.append((display, key))

        error_message = ""
        if not unique_models:
            error_message = "No models reported by Ollama."

        self.after(0, lambda: self._finalize_model_list(unique_models, error_message))

    def _finalize_model_list(self, models: List[Tuple[str, Optional[str]]], error_message: str) -> None:
        """Update the UI with fetched model options."""
        if not self.model_combobox:
            return

        self._available_models = models
        self._model_option_map = {display: key for display, key in models}

        if models:
            values = [display for display, _ in models]
            self.model_combobox.configure(values=values, state="readonly")

            previous_display = self.model_var.get()
            if previous_display in self._model_option_map:
                self.model_combobox.set(previous_display)
            else:
                self.model_combobox.current(0)

            if self.summarize_button:
                self.summarize_button.configure(state="normal")

            if error_message:
                self.status_message.set(f"Model warning: {error_message}")
        else:
            self.model_combobox.configure(values=[], state="disabled")
            self.model_var.set("No models available")
            if self.summarize_button:
                self.summarize_button.configure(state="disabled")

            if error_message:
                self.status_message.set(f"LLM error: {error_message}")
            else:
                self.status_message.set("No LLM models detected.")
    @staticmethod
    def scan_directory_for_excels(directory: Path) -> Iterable[Path]:
        """Return all Excel files in the given directory."""
        return (path for path in directory.iterdir() if path.suffix.lower() in {".xlsx", ".xls"} and path.is_file())

    def load_selected_file(self) -> None:
        """Read the selected Excel file and populate the treeview."""
        selection = self.selected_file.get()
        if not selection:
            messagebox.showinfo("No file selected", "Please choose an Excel file from the dropdown.")
            return

        file_path = self._current_directory / selection

        try:
            df = pd.read_excel(file_path)
        except FileNotFoundError:
            messagebox.showerror("File not found", f"The file '{selection}' could not be found.")
            return
        except PermissionError:
            messagebox.showerror("Permission denied", f"Insufficient permissions to read '{selection}'.")
            return
        except Exception as exc:
            messagebox.showerror("Error loading file", f"An error occurred while reading '{selection}'.\n\n{exc}")
            return

        if df.empty:
            messagebox.showinfo("Empty file", f"The Excel file '{selection}' does not contain any data.")
            self.clear_treeview()
            self.show_noti_text("")
            return

        self._current_df = df
        self.populate_treeview(df)
        self.show_noti_text("")
        self.update_summary_output("")
        self.status_message.set("")
        self.after(10, self.tree.focus_set)

    def choose_directory(self) -> None:
        """Prompt the user to choose a directory and refresh the file list."""
        selected_dir = filedialog.askdirectory(mustexist=True, initialdir=self._current_directory)
        if not selected_dir:
            return

        new_path = Path(selected_dir)
        if not new_path.exists() or not new_path.is_dir():
            messagebox.showerror("Invalid Directory", f"The selected path is not a valid directory:\n{selected_dir}")
            return

        self._current_directory = new_path
        self.current_directory_var.set(str(self._current_directory))
        self._populate_excel_list()

    def update_summary_output(self, text: str) -> None:
        """Render text into the summary panel."""
        if not self.summary_text_widget:
            return
        self.summary_text_widget.configure(state="normal")
        self.summary_text_widget.delete("1.0", tk.END)
        if text:
            self.summary_text_widget.insert("1.0", text)
        self.summary_text_widget.configure(state="disabled")

    def summarize_current_text(self) -> None:
        """Send the Noti Text and optional user prompt to the selected LLM."""
        if not self._ollama_available:
            messagebox.showerror(
                "Summarize",
                "ollama package is not available. Please install or configure an Ollama server.",
            )
            return

        noti_content = self.current_noti_text.strip()
        if not noti_content:
            messagebox.showinfo("Summarize", "Please select a row with a 'Noti Text' value before summarizing.")
            return

        prompt = self.prompt_var.get().strip()
        combined_prompt_parts = []
        if prompt:
            combined_prompt_parts.append(prompt)
        combined_prompt_parts.append("Noti Text:\n" + noti_content)
        combined_prompt = "\n\n".join(combined_prompt_parts)

        selected_display = self.model_var.get()
        model_key = self._model_option_map.get(selected_display)

        temperature = self._get_temperature()
        if temperature is None:
            messagebox.showerror("Summarize", "Temperature must be a number between 0.0 and 2.0.")
            return

        if self.summarize_button:
            self.summarize_button.configure(state="disabled")
        self.status_message.set("Summarizing...")
        self.update_summary_output("Summarizing...")

        thread = threading.Thread(
            target=self._run_summarization,
            args=(model_key, combined_prompt, temperature),
            daemon=True,
        )
        thread.start()

    def _get_temperature(self) -> Optional[float]:
        """Validate and clamp the user-provided temperature."""
        try:
            value = float(self.temperature_var.get())
        except (tk.TclError, ValueError, TypeError):
            return None

        if value < 0.0 or value > 2.0:
            value = max(0.0, min(2.0, value))
            self.temperature_var.set(round(value, 1))

        return float(value)

    def _run_summarization(self, model_key: Optional[str], combined_prompt: str, temperature: float) -> None:
        """Perform the LLM call in a worker thread."""
        host_url, error = self._parse_api_host()
        if error:
            self.after(0, lambda: self._on_summarization_error(error))
            return

        client, client_error = self._create_ollama_client(host_url)
        if client_error:
            self.after(0, lambda: self._on_summarization_error(client_error))
            return

        target_model = model_key
        if not target_model and self._available_models:
            target_model = self._available_models[0][1]

        if not target_model:
            self.after(0, lambda: self._on_summarization_error("No model selected for summarization."))
            return

        options = {"temperature": float(temperature)}

        try:
            response = client.generate(model=target_model, prompt=combined_prompt, options=options)
        except Exception as exc:  # noqa: broad-except
            self.after(0, lambda: self._on_summarization_error(f"Ollama request failed: {exc}"))
            return

        summary = self._extract_response_content(response)
        if not summary:
            summary = "No response generated."

        self.after(0, lambda: self._on_summarization_success(summary))

    def _extract_response_content(self, response: object) -> str:
        """Best-effort extraction of text content from an LLM response."""
        if response is None:
            return ""

        if isinstance(response, str):
            return response.strip()

        if isinstance(response, dict):
            for key in ("content", "text", "message", "output", "response"):
                value = response.get(key)
                if value:
                    return str(value).strip()
            return ""

        # Objects with relevant attributes.
        for attr_name in ("content", "text", "message", "output", "response"):
            if hasattr(response, attr_name):
                value = getattr(response, attr_name)
                if value:
                    return str(value).strip()

        if hasattr(response, "parsed"):
            parsed = getattr(response, "parsed")
            if parsed:
                return str(parsed).strip()

        if hasattr(response, "choices"):
            choices = getattr(response, "choices")
            if isinstance(choices, Iterable):
                for choice in choices:
                    text = self._extract_response_content(choice)
                    if text:
                        return text

        if isinstance(response, Iterable) and not isinstance(response, (bytes, bytearray)):
            fragments = []
            for item in response:
                fragment = self._extract_response_content(item)
                if fragment:
                    fragments.append(fragment)
            if fragments:
                return "\n".join(fragments).strip()

        return str(response).strip()

    def _on_summarization_success(self, summary: str) -> None:
        """Handle successful summarization."""
        if self.summarize_button:
            self.summarize_button.configure(state="normal")

        self.status_message.set("Summarization complete.")
        self.update_summary_output(summary)

    def _on_summarization_error(self, message: str) -> None:
        """Handle summarization failures."""
        if self.summarize_button:
            self.summarize_button.configure(state="normal")

        self.status_message.set(f"Summarization failed: {message}")
        self.update_summary_output("")
        messagebox.showerror("Summarize", message)

    def populate_treeview(self, df: pd.DataFrame) -> None:
        """Populate the treeview with dataframe data."""
        self.clear_treeview()

        self._current_columns = list(df.columns.astype(str))
        self.tree["columns"] = self._current_columns

        for col in self._current_columns:
            self.tree.heading(col, text=col)

        # Autosize columns: set width based on column header and sample data.
        max_sample_rows = min(len(df), 1000)
        samples = df.head(max_sample_rows).fillna("").astype(str)

        for col in self._current_columns:
            header_width = self._tree_font.measure(col) + 20
            max_value_width = max((self._tree_font.measure(value) for value in samples[col]), default=header_width)
            col_width = max(header_width, min(max_value_width + 30, 400))
            self.tree.column(col, width=col_width, anchor="w", stretch=True)

        # Apply zebra striping tags.
        self.tree.tag_configure("oddrow", background="#f2f2f2")
        self.tree.tag_configure("evenrow", background="white")

        # Insert rows efficiently.
        for idx, row in enumerate(df.itertuples(index=False, name=None)):
            values = tuple("" if pd.isna(value) else str(value) for value in row)
            tag = "oddrow" if idx % 2 else "evenrow"
            self.tree.insert("", "end", values=values, tags=(tag,))

    def clear_treeview(self) -> None:
        """Remove all existing rows from the treeview."""
        self.current_noti_text = ""
        self.update_summary_output("")
        for item in self.tree.get_children():
            self.tree.delete(item)

    def on_row_select(self, event: tk.Event) -> None:
        """Update the Noti Text panel with the selected row value."""
        if not self._current_columns:
            return

        noti_index = None
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
        if pd.isna(noti_value) or noti_value == "":
            self.show_noti_text("")
        else:
            self.status_message.set("")
            self.show_noti_text(str(noti_value))

    def show_noti_text(self, value: str) -> None:
        """Display text in the Noti Text panel."""
        self.current_noti_text = value or ""
        self.noti_text.configure(state="normal")
        self.noti_text.delete("1.0", tk.END)
        if value:
            self.noti_text.insert("1.0", value)
        self.noti_text.configure(state="disabled")


if __name__ == "__main__":
    try:
        MainApp().mainloop()
    except Exception as error:
        messagebox.showerror("Application Error", f"An unexpected error occurred.\n\n{error}")

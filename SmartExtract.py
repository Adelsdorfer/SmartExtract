"""Desktop GUI application to browse Excel files and view 'Noti Text' content."""

from __future__ import annotations

import threading
import tkinter as tk
import tkinter.font as tkfont
from collections.abc import Iterable
from pathlib import Path
from typing import Dict, List, Optional, Tuple
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
MODEL = "qwen3:8b"
#MODEL = "gemma3:4b"
#MODEL = "deepseek-r1:8b"
#MODEL = "phi4-mini-reasoning"
#MODEL = "granite4:tiny-h"
#MODEL = "deepseek-r1:8b"

#MODEL = "gpt-oss:20b"


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
        self.prompt_var = tk.StringVar(value="Summarize the following CT service report, focusing only on the essential technical information: parts used or replaced, root cause (if mentioned), on-site visits, and key service actions. Write the summary in a clear and structured format.")
        self.model_var = tk.StringVar(value=MODEL)
        self.temperature_var = tk.DoubleVar(value=0.5)
        self._model_option_map: Dict[str, Optional[str]] = {MODEL: MODEL}
        self._available_models: List[Tuple[str, Optional[str]]] = [(MODEL, MODEL)]
        self.current_noti_text = ""
        self._configured_api_host: Optional[str] = None

        self.model_combobox: Optional[ttk.Combobox] = None
        self.summarize_button: Optional[ttk.Button] = None
        self.check_server_button: Optional[ttk.Button] = None
        self.server_status_label: Optional[ttk.Label] = None
        self.summary_text_widget: Optional[tk.Text] = None

        self.file_combobox: Optional[ttk.Combobox] = None
        self.tree: Optional[ttk.Treeview] = None
        self.noti_text: Optional[tk.Text] = None

        self._tree_font = tkfont.nametofont("TkDefaultFont")

        self._configure_style()
        self._build_ui()
        self._populate_excel_list()
        self.update_summary_output("")

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

        ttk.Label(noti_frame, text="Prompt:").grid(row=1, column=0, sticky="w", padx=(0, 4))
        prompt_entry = ttk.Entry(noti_frame, textvariable=self.prompt_var, width=40)
        prompt_entry.grid(row=1, column=1, sticky="ew", padx=(0, 8))

        ttk.Label(noti_frame, text="Model:").grid(row=1, column=2, sticky="e", padx=(0, 4))
        self.model_combobox = ttk.Combobox(
            noti_frame,
            textvariable=self.model_var,
            state="readonly",
            width=30,
            values=[MODEL],
        )
        self.model_combobox.grid(row=1, column=3, sticky="ew", padx=(0, 8))
        self.model_combobox.current(0)
        self.model_combobox.configure(state="disabled")

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

        noti_text_frame = ttk.Frame(noti_frame)
        noti_text_frame.grid(row=2, column=0, columnspan=7, sticky="nsew", pady=(8, 0))
        noti_frame.rowconfigure(2, weight=1)
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

        summary_frame = ttk.Labelframe(content, text="LLM Output", style="Noti.TLabelframe")
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

    def check_server_status(self) -> None:
        """Check whether the configured Ollama server is reachable."""
        if not self.check_server_button:
            return

        host_value = self.api_host_var.get()
        try:
            normalized_host = self._parse_api_host(host_value)
        except ValueError as exc:
            messagebox.showerror("Ollama Server", str(exc))
            return

        if normalized_host is None:
            messagebox.showerror("Ollama Server", "Please provide an API host.")
            return

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
            if self.check_server_button:
                self.after(0, lambda: self.check_server_button.configure(state="normal"))

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
        """Enable the summarize button only when the server and Noti Text are ready."""
        if not self.summarize_button:
            return

        if self._server_reachable and self.current_noti_text.strip():
            self.summarize_button.configure(state="normal")
        else:
            self.summarize_button.configure(state="disabled")

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
        user_prompt = self.prompt_var.get().strip()
        prompt_parts: List[str] = []
        if user_prompt:
            prompt_parts.append(user_prompt)
        prompt_parts.append(f"Noti Text:\n{self.current_noti_text}")
        combined_prompt = "\n\n".join(prompt_parts)

        self.status_message.set("Processing Input, please wait...")
        self.update_summary_output("")
        if self.summarize_button:
            self.summarize_button.configure(state="disabled")

        threading.Thread(
            target=self._run_ollama_stream,
            args=(combined_prompt, temperature),
            daemon=True,
        ).start()

    def _run_ollama_stream(self, combined_prompt: str, temperature: float) -> None:
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
                model=MODEL,
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


if __name__ == "__main__":
    try:
        MainApp().mainloop()
    except Exception as error:  # pylint: disable=broad-except
        messagebox.showerror("Application Error", f"An unexpected error occurred.\n\n{error}")

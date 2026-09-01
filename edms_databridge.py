"""
EDMS DataBridge
-----------------------
Author: Ashley Powell (GitHub: Ash Kapow)
Built for: EDMS

A simple Windows GUI tool: user uploads a JSON export (e.g. from Ambunet),
the app converts it into a clean Excel file they can actually use.

Currently uses a GENERIC flattening approach since we don't yet know the
real structure of Ambunet's export. Once you have a sample export, replace
the `process_data()` function with logic specific to that schema (e.g.
splitting patients/shifts/HR records into separate sheets, renaming
columns, converting date formats, etc).

--- Build into a standalone .exe ---
Run build.bat (see that file for the exact pyinstaller command/flags).
The .exe will be in the generated dist/ folder. That single file is what
you hand to the non-technical user - no installer, no Python needed.
"""

import json
import sys
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd
from tkinterdnd2 import DND_FILES, TkinterDnD


APP_TITLE = "EDMS DataBridge"


def resource_path(relative_path: str) -> Path:
    """Resolve a bundled asset path, in both dev mode and a PyInstaller onefile build."""
    base_path = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
    return base_path / relative_path


def load_logo_image():
    """Load assets/logo.png if it exists. Returns None if it's missing or unreadable."""
    logo_path = resource_path("assets/logo.png")
    if not logo_path.exists():
        return None
    try:
        return tk.PhotoImage(file=str(logo_path))
    except tk.TclError:
        return None


def load_json(filepath: str):
    """Load and parse the uploaded JSON file. Raises on invalid JSON."""
    with open(filepath, "r", encoding="utf-8-sig") as f:
        return json.load(f)


def parse_dnd_filepaths(data: str) -> list:
    """
    Split a tkinterdnd2 <<Drop>> event's data string into individual file
    paths. Paths containing spaces arrive wrapped in {curly braces};
    others are just space-separated.
    """
    paths = []
    i, n = 0, len(data)
    while i < n:
        if data[i].isspace():
            i += 1
        elif data[i] == "{":
            end = data.index("}", i)
            paths.append(data[i + 1:end])
            i = end + 1
        else:
            end = i
            while end < n and not data[end].isspace():
                end += 1
            paths.append(data[i:end])
            i = end
    return paths


def process_data(data):
    """
    THIS IS THE FUNCTION TO REPLACE ONCE YOU SEE THE REAL AMBUNET EXPORT.

    Right now it does a generic best-effort flatten of whatever JSON shape
    it's given, and returns a dict of {sheet_name: DataFrame}.

    Later, once you know the actual structure (e.g. top-level keys like
    "patients", "shifts", "employees"), you'll likely want to:
      - Split each top-level key into its own sheet
      - Rename/reorder columns to something human-readable
      - Reformat dates, IDs, etc.
      - Drop internal/system fields the user doesn't need to see

    For now this just tries to produce *something* readable no matter what
    shape of JSON comes in.
    """
    sheets = {}

    if isinstance(data, list):
        # A flat (or nested) list of records -> one sheet
        sheets["Data"] = pd.json_normalize(data)

    elif isinstance(data, dict):
        # If the top-level dict has list-valued keys, treat each as its own
        # "table"/sheet (this matches how a lot of SaaS exports are shaped,
        # e.g. {"patients": [...], "shifts": [...], "employees": [...]})
        list_keys = {k: v for k, v in data.items() if isinstance(v, list)}
        if list_keys:
            for key, records in list_keys.items():
                sheet_name = str(key)[:31]  # Excel sheet name limit
                sheets[sheet_name] = pd.json_normalize(records)
        else:
            # Single flat object, no list fields - just show it as one row
            sheets["Data"] = pd.json_normalize(data)
    else:
        raise ValueError(
            "Unrecognized JSON structure (expected a list or object at the top level)."
        )

    return sheets


def save_as_excel(sheets: dict, output_path: str):
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)


class App(TkinterDnD.Tk):
    WINDOW_WIDTH = 480

    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.resizable(False, False)

        self.drop_target_register(DND_FILES)
        self.dnd_bind("<<Drop>>", self.handle_drop)

        icon_path = resource_path("assets/logo.ico")
        if icon_path.exists():
            self.iconbitmap(str(icon_path))

        ttk.Style(self).configure("Upload.TButton", font=("Segoe UI", 12))

        self._logo_image = load_logo_image()  # kept as an attribute so Tk doesn't GC it
        if self._logo_image is not None:
            ttk.Label(self, image=self._logo_image).pack(pady=(20, 4))
            title_pady = (0, 8)
        else:
            title_pady = (24, 8)

        ttk.Label(
            self, text=APP_TITLE, font=("Segoe UI", 16, "bold")
        ).pack(pady=title_pady)

        ttk.Label(
            self,
            text="Click below, or drag a JSON file onto this window,\n"
                 "and this will create a formatted Excel file next to it.",
            justify="center",
        ).pack(pady=(0, 20))

        ttk.Button(
            self,
            text="Upload JSON File",
            padding=(20, 12),
            style="Upload.TButton",
            command=self.handle_upload,
        ).pack()

        self.status_label = ttk.Label(
            self, text="", foreground="gray20", wraplength=self.WINDOW_WIDTH - 40
        )
        self.status_label.pack(pady=(20, 0))

        ttk.Label(
            self,
            text="Built for EDMS by Ashley Powell (Ash Kapow)",
            font=("Segoe UI", 8),
            foreground="gray50",
        ).pack(side="bottom", pady=(0, 10))

        self._fit_window_to_content()

    def _set_status(self, text):
        """Update the status text and resize the window's height to fit it -
        a fixed height would silently clip content (e.g. a long saved-file
        path, or the footer) whenever a message needs more room than
        whatever was guessed at design time."""
        self.status_label.config(text=text)
        self._fit_window_to_content()

    def _fit_window_to_content(self):
        self.update_idletasks()
        self.geometry(f"{self.WINDOW_WIDTH}x{self.winfo_reqheight()}")

    def handle_upload(self):
        filepath = filedialog.askopenfilename(
            title="Select the Ambunet JSON export",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not filepath:
            return
        self.process_file(filepath)

    def handle_drop(self, event):
        paths = parse_dnd_filepaths(event.data)
        if not paths:
            return
        self.process_file(paths[0])

    def process_file(self, filepath):
        self._set_status("Processing...")

        try:
            data = load_json(filepath)
            sheets = process_data(data)

            src = Path(filepath)
            default_out = src.with_name(src.stem + "_formatted.xlsx")

            output_path = filedialog.asksaveasfilename(
                title="Save formatted file as",
                initialfile=default_out.name,
                initialdir=str(src.parent),
                defaultextension=".xlsx",
                filetypes=[("Excel file", "*.xlsx")],
            )
            if not output_path:
                self._set_status("Cancelled.")
                return

            save_as_excel(sheets, output_path)

            self._set_status(f"Done! Saved to:\n{output_path}")
            messagebox.showinfo(
                APP_TITLE,
                f"Success! Your formatted file is ready:\n\n{output_path}",
            )

        except json.JSONDecodeError:
            self._set_status("")
            messagebox.showerror(
                APP_TITLE,
                "That file doesn't look like valid JSON.\n"
                "Please double check the file you uploaded.",
            )
        except Exception as e:
            self._set_status("")
            messagebox.showerror(
                APP_TITLE,
                f"Something went wrong:\n\n{e}",
            )
            traceback.print_exc()


if __name__ == "__main__":
    app = App()
    app.mainloop()

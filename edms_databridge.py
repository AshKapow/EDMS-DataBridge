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
1. pip install pyinstaller pandas openpyxl
2. pyinstaller --onefile --windowed --name "EDMSDataBridge" edms_databridge.py
3. The .exe will be in the generated dist/ folder. That single file is what
   you hand to the non-technical user - no installer, no Python needed.
"""

import json
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd


APP_TITLE = "EDMS DataBridge"


def load_json(filepath: str):
    """Load and parse the uploaded JSON file. Raises on invalid JSON."""
    with open(filepath, "r", encoding="utf-8-sig") as f:
        return json.load(f)


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


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("480x290")
        self.resizable(False, False)

        tk.Label(
            self, text=APP_TITLE, font=("Segoe UI", 16, "bold")
        ).pack(pady=(24, 8))

        tk.Label(
            self,
            text="Click below, choose the JSON file from Ambunet,\n"
                 "and this will create a formatted Excel file next to it.",
            justify="center",
        ).pack(pady=(0, 20))

        tk.Button(
            self,
            text="Upload JSON File",
            font=("Segoe UI", 12),
            width=22,
            height=2,
            command=self.handle_upload,
        ).pack()

        self.status_label = tk.Label(self, text="", fg="gray20")
        self.status_label.pack(pady=(20, 0))

        tk.Label(
            self,
            text="Built for EDMS by Ashley Powell (Ash Kapow)",
            font=("Segoe UI", 8),
            fg="gray50",
        ).pack(side="bottom", pady=(0, 10))

    def handle_upload(self):
        filepath = filedialog.askopenfilename(
            title="Select the Ambunet JSON export",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not filepath:
            return

        self.status_label.config(text="Processing...")
        self.update_idletasks()

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
                self.status_label.config(text="Cancelled.")
                return

            save_as_excel(sheets, output_path)

            self.status_label.config(text=f"Done! Saved to:\n{output_path}")
            messagebox.showinfo(
                APP_TITLE,
                f"Success! Your formatted file is ready:\n\n{output_path}",
            )

        except json.JSONDecodeError:
            self.status_label.config(text="")
            messagebox.showerror(
                APP_TITLE,
                "That file doesn't look like valid JSON.\n"
                "Please double check the file you uploaded.",
            )
        except Exception as e:
            self.status_label.config(text="")
            messagebox.showerror(
                APP_TITLE,
                f"Something went wrong:\n\n{e}",
            )
            traceback.print_exc()


if __name__ == "__main__":
    app = App()
    app.mainloop()

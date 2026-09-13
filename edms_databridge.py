"""
EDMS DataBridge
-----------------------
Author: Ash Kapow
Built for: EDMS

A simple Windows GUI tool: user uploads a JSON export (e.g. from Ambunet)
- a zip, a folder, or a single JSON file - and the app converts it into a
clean Excel file they can actually use.

Ambunet's real export is a zip (containing a folder of one *.json file
per entity, e.g. employees.json, incidents.json, epcrs.json) in MongoDB
Extended JSON format (IDs as {"$oid": ...}, dates as {"$date": ...}, etc).
`clean_data()` unwraps that into plain values and drops known-sensitive
fields (e.g. password hashes) before `process_data()` flattens each
entity into its own sheet. `process_data()` is still a GENERIC flatten
per entity, not bespoke per-entity column mapping/renaming - see the
README's open questions for what's still deliberately deferred.

--- Build into a standalone .exe ---
Run build.bat (see that file for the exact pyinstaller command/flags).
The .exe will be in the generated dist/ folder. That single file is what
you hand to the non-technical user - no installer, no Python needed.
"""

import json
import sys
import traceback
import zipfile
from datetime import datetime, timedelta, timezone
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


def _decode_json_bytes(raw_bytes: bytes) -> str:
    """
    Decode a JSON file's raw bytes, tolerating mixed encodings across a
    118-file export. utf-8-sig is standard for JSON, but some files in
    the real Ambunet export turned out to be Windows-1252 (e.g. a "£" in
    an expense/invoice field decodes fine in cp1252, not utf-8). latin-1
    never raises - it's the last-resort fallback since a wrong-but-parsed
    character beats crashing the whole 118-file batch over one file.
    """
    for encoding in ("utf-8-sig", "cp1252"):
        try:
            return raw_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw_bytes.decode("latin-1")


def load_json(filepath: str):
    """Load and parse the uploaded JSON file. Raises on invalid JSON."""
    with open(filepath, "rb") as f:
        return json.loads(_decode_json_bytes(f.read()))


def load_entity_folder(folder_path):
    """
    Load every *.json file found in a folder (searched recursively, since
    depending how someone extracts Ambunet's zip, the inner folder may or
    may not still be there) into one dict keyed by filename - e.g.
    employees.json -> "employees" -> the parsed list of employee records.

    Returns (data, skipped): a single unparseable file doesn't abort the
    whole batch (this is a 100+ file export - one bad file shouldn't cost
    you the other 117). skipped is a list of (filename, error message)
    for anything that couldn't be read, so it's still visible afterward.
    """
    folder_path = Path(folder_path)
    data, skipped = {}, []
    for json_file in sorted(folder_path.rglob("*.json")):
        try:
            data[json_file.stem] = load_json(str(json_file))
        except json.JSONDecodeError as e:
            skipped.append((json_file.name, str(e)))
    return data, skipped


def load_entity_zip(zip_path):
    """Same as load_entity_folder(), but reads directly from a zip archive
    without extracting it first - matches Ambunet's real export format
    (a zip containing one folder of *.json files, one per entity)."""
    data, skipped = {}, []
    with zipfile.ZipFile(zip_path) as zf:
        for name in zf.namelist():
            if name.lower().endswith(".json"):
                try:
                    text = _decode_json_bytes(zf.read(name))
                    data[Path(name).stem] = json.loads(text)
                except json.JSONDecodeError as e:
                    skipped.append((Path(name).name, str(e)))
    return data, skipped


SENSITIVE_FIELD_NAMES = {"password", "passwordhash", "secret", "apikey", "token"}


def unwrap_extended_json(value):
    """
    Recursively convert MongoDB Extended JSON wrapper objects into plain
    Python values: {"$oid": "abc"} -> "abc", {"$date": "2024-01-01..."} ->
    "2024-01-01...", {"$date": {"$numberLong": "..."}} -> an ISO date
    string, {"$numberLong"/"$numberInt"/"$numberDouble"/"$numberDecimal":
    "123"} -> a plain int/float.

    Without this, a generic flatten turns these into unreadable columns
    like "dob.$date.$numberLong" instead of a normal ID/date/number.
    """
    if isinstance(value, dict):
        keys = set(value.keys())
        if keys == {"$oid"}:
            return value["$oid"]
        if keys == {"$date"}:
            inner = value["$date"]
            if isinstance(inner, dict) and set(inner.keys()) == {"$numberLong"}:
                millis = int(inner["$numberLong"])
                # datetime.fromtimestamp() raises OSError on Windows for
                # pre-1970 dates (e.g. a date of birth) - pure arithmetic
                # from the epoch works for any date, any platform.
                epoch = datetime(1970, 1, 1, tzinfo=timezone.utc)
                return (epoch + timedelta(milliseconds=millis)).isoformat()
            return inner
        if keys == {"$numberLong"} or keys == {"$numberInt"}:
            return int(next(iter(value.values())))
        if keys == {"$numberDouble"} or keys == {"$numberDecimal"}:
            return float(next(iter(value.values())))
        return {k: unwrap_extended_json(v) for k, v in value.items()}
    if isinstance(value, list):
        return [unwrap_extended_json(v) for v in value]
    return value


def redact_sensitive_fields(value):
    """Recursively drop fields whose name suggests a credential/secret
    (password hashes, tokens, etc.) - these should never land in an Excel
    file handed to non-technical staff, even in demo data."""
    if isinstance(value, dict):
        return {
            k: redact_sensitive_fields(v)
            for k, v in value.items()
            if k.lower() not in SENSITIVE_FIELD_NAMES
        }
    if isinstance(value, list):
        return [redact_sensitive_fields(v) for v in value]
    return value


def clean_data(value):
    """Unwrap MongoDB Extended JSON types and drop sensitive fields."""
    return redact_sensitive_fields(unwrap_extended_json(value))


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
    Cleans (unwraps Extended JSON, redacts sensitive fields) then does a
    generic best-effort flatten of whatever JSON shape it's given, and
    returns a dict of {sheet_name: DataFrame}. For a multi-file export
    (see load_entity_folder()/load_entity_zip()), `data` is already a
    dict of {entity_name: [records]} by the time it gets here, so each
    entity naturally becomes its own sheet via the dict branch below.

    STILL GENERIC PER ENTITY, NOT BESPOKE: column names are still whatever
    the raw field names are, deeply-nested repeating sub-records (e.g. an
    EPCR's vitals-over-time or a vehicle's service history) still flatten
    with numeric-indexed columns rather than their own linked sheet, and
    nothing is renamed/reordered for readability yet. Once there's a
    reason to polish a specific entity's output, that's the kind of
    per-entity logic to add here (see the README's open questions).
    """
    data = clean_data(data)
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
            text="Click a button below, or drag your export (zip, folder,\n"
                 "or a single JSON file) onto this window.",
            justify="center",
        ).pack(pady=(0, 20))

        button_row = ttk.Frame(self)
        button_row.pack()
        ttk.Button(
            button_row,
            text="Upload ZIP File",
            padding=(16, 12),
            style="Upload.TButton",
            command=self.handle_upload_zip,
        ).pack(side="left", padx=(0, 8))
        ttk.Button(
            button_row,
            text="Upload Folder",
            padding=(16, 12),
            style="Upload.TButton",
            command=self.handle_upload_folder,
        ).pack(side="left")

        self.status_label = ttk.Label(
            self, text="", foreground="gray20", wraplength=self.WINDOW_WIDTH - 40
        )
        self.status_label.pack(pady=(20, 0))

        ttk.Label(
            self,
            text="Built for EDMS by Ash Kapow",
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

    def handle_upload_zip(self):
        filepath = filedialog.askopenfilename(
            title="Select the Ambunet export zip",
            filetypes=[("Zip files", "*.zip"), ("All files", "*.*")],
        )
        if not filepath:
            return
        self.process_path(Path(filepath))

    def handle_upload_folder(self):
        folder = filedialog.askdirectory(title="Select the Ambunet export folder")
        if not folder:
            return
        self.process_path(Path(folder))

    def handle_drop(self, event):
        paths = parse_dnd_filepaths(event.data)
        if not paths:
            return
        self.process_path(Path(paths[0]))

    def process_path(self, path: Path):
        """Handle a dropped or picked path, whichever of the three supported
        shapes it turns out to be: a folder, a zip, or a single JSON file."""
        self._set_status("Processing...")

        try:
            skipped = []
            if path.is_dir():
                data, skipped = load_entity_folder(path)
                default_name = f"{path.name}_formatted.xlsx"
                default_dir = path.parent
            elif path.suffix.lower() == ".zip":
                data, skipped = load_entity_zip(path)
                default_name = f"{path.stem}_formatted.xlsx"
                default_dir = path.parent
            else:
                data = load_json(str(path))
                default_name = f"{path.stem}_formatted.xlsx"
                default_dir = path.parent

            sheets = process_data(data)

            output_path = filedialog.asksaveasfilename(
                title="Save formatted file as",
                initialfile=default_name,
                initialdir=str(default_dir),
                defaultextension=".xlsx",
                filetypes=[("Excel file", "*.xlsx")],
            )
            if not output_path:
                self._set_status("Cancelled.")
                return

            save_as_excel(sheets, output_path)

            self._set_status(f"Done! Saved to:\n{output_path}")
            success_message = f"Success! Your formatted file is ready:\n\n{output_path}"
            if skipped:
                names = "\n".join(f"  - {name}" for name, _ in skipped)
                success_message += (
                    f"\n\n{len(skipped)} file(s) couldn't be read and were skipped:\n{names}"
                )
            messagebox.showinfo(APP_TITLE, success_message)

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

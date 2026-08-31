# EDMS DataBridge

[![CI](https://github.com/AshKapow/EDMS-DataBridge/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/AshKapow/EDMS-DataBridge/actions/workflows/ci.yml)
[![Python 3.11](https://img.shields.io/badge/python-3.11-blue?logo=python&logoColor=white)](https://www.python.org/)
[![Platform: Windows](https://img.shields.io/badge/platform-windows-0078D6?logo=windows&logoColor=white)](https://github.com/AshKapow/EDMS-DataBridge)
[![Ruff](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/ruff/main/assets/badge/v2.json)](https://github.com/astral-sh/ruff)
[![License: Custom](https://img.shields.io/badge/license-custom-lightgrey)](LICENSE)

A small Windows desktop tool that lets a non-technical user upload a JSON
export (e.g. from Ambunet) and get back a clean, formatted Excel file.

Built for EDMS by Ashley Powell (Ash Kapow on GitHub). See [LICENSE](LICENSE)
for usage terms.

## Background

The company runs almost all core systems (patients, HR, shifts, etc.) on a
third-party SaaS called Ambunet. If the company ever leaves Ambunet, the
only guaranteed way to get data out is a raw JSON export — no promised
format or structure. The author is one of the only technical people at the
company, so this tool exists to make that export usable by a non-technical
staff member without needing manual help under time pressure, if/when a
departure ever becomes urgent.

We do not yet have a real sample of Ambunet's export JSON, so we don't know
if it's one file or several, a flat list or deeply nested, one file per
entity type, etc. That's why the app is deliberately split into two layers:
a generic, working core (upload -> flatten any reasonable JSON shape ->
Excel) that's already useful today, and one isolated function,
`process_data()` in `edms_databridge.py`, meant to be rewritten with
real schema-specific logic once an actual sample export exists.

## Status

Early scaffold. The JSON structure of the real Ambunet export isn't known
yet, so `edms_databridge.py` currently does a **generic flatten**: it
turns any list-of-records or dict-of-lists JSON into one Excel file (with
one sheet per top-level list). Once we get a real sample export, update
`process_data()` in `edms_databridge.py` with schema-specific logic
(column renaming, date formatting, splitting entities properly, etc).

Not yet tested against a real Ambunet export (none exists yet), and not
yet tested as a built `.exe` on a clean machine.

## Tech decisions

Python + Tkinter + pandas/openpyxl, packaged as a single unsigned `.exe`
via PyInstaller (`--onefile --windowed`), chosen over C#/.NET or Electron
for speed of iteration given the author's background, and because a single
unsigned `.exe` is enough for an internal tool — no installer needed.

Output format is Excel (`.xlsx`), one sheet per top-level entity type.

## Setup (dev machine)

```
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements-dev.txt
```

`requirements-dev.txt` includes everything in `requirements.txt` (the
runtime deps used by the exe) plus `pytest` and `ruff` for testing/linting.

## Run without building an exe (for testing/dev)

```
python edms_databridge.py
```

## Testing & linting

Unit tests cover `load_json()` and `process_data()` — the core, GUI-free
logic. The Tkinter GUI itself is exercised manually, not by automated
tests.

```
pytest
ruff check .
```

CI (GitHub Actions, `.github/workflows/ci.yml`) runs both on every push and
pull request against `main`, then does a smoke-test build of the exe with
PyInstaller and uploads it as a workflow artifact, so a working build is
always downloadable without needing a local Python setup.

## Build the standalone .exe

```
build.bat
```

This produces `dist\EDMSDataBridge.exe` — a single file with no
dependencies. That's the file to hand to end users. They just double-click
it, no Python install needed on their machine.

Note: the exe is unsigned, so Windows SmartScreen will show a warning on
first run ("Windows protected your PC"). Users click "More info" → "Run
anyway". A code-signing certificate would remove this, if it becomes worth
the cost for wider distribution.

## Project structure

```
edms-export-translater/
├── edms_databridge.py     # main app (GUI + processing logic)
├── tests/                 # pytest unit tests for the processing logic
├── requirements.txt       # runtime deps (bundled into the exe)
├── requirements-dev.txt   # runtime deps + pytest/ruff for local dev & CI
├── pyproject.toml         # pytest and ruff config
├── build.bat              # builds the standalone exe
├── .github/workflows/     # CI: lint, test, and build-smoke-test on push/PR
├── .gitignore
├── LICENSE
└── README.md
```

## Open questions / next steps

1. **Real schema** — get a real (even small, anonymized) sample JSON export
   from Ambunet, then rewrite `process_data()` for it: proper column
   names/order, date formatting, splitting entities correctly, dropping
   internal/system fields. This is the priority once a sample exists.
2. **Right target format** — is Excel actually the right end format, or is
   this data meant to feed another system, in which case CSV or a
   different JSON shape might matter more than a spreadsheet?
3. **Distribution/signing** — still unsigned, so Windows SmartScreen warns
   on first run. Is a code-signing certificate worth it, or is "click More
   info -> Run anyway" an acceptable one-time instruction for internal
   staff?
4. **Drag-and-drop** — not included yet (Tkinter needs the extra
   `tkinterdnd2` dependency for this). Worth adding once the core logic is
   settled?
5. **Clean-machine testing** — the built `.exe` hasn't been tested on a
   machine without dev tools/antivirus false-positive checks yet.

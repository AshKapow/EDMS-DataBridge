# EDMS DataBridge

[![CI](https://github.com/AshKapow/EDMS-DataBridge/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/AshKapow/EDMS-DataBridge/actions/workflows/ci.yml)
[![Python 3.11](https://img.shields.io/badge/python-3.11-blue?logo=python&logoColor=white)](https://www.python.org/)
[![Platform: Windows](https://img.shields.io/badge/platform-windows-0078D6?logo=windows&logoColor=white)](https://github.com/AshKapow/EDMS-DataBridge)
[![Ruff](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/ruff/main/assets/badge/v2.json)](https://github.com/astral-sh/ruff)
[![License: MIT](https://img.shields.io/badge/license-MIT-yellow.svg)](LICENSE)

A small Windows desktop tool that lets a non-technical user upload (or
drag and drop) a JSON export (e.g. from Ambunet) and get back clean,
formatted Excel and PDF files.

Built primarily for EDMS (Emergency Doctors Medical Service) by Ash Kapow,
and released as open source (MIT license —
see [LICENSE](LICENSE)) for anyone else who runs into the same problem.

## Background

The company runs almost all core systems (patients, HR, shifts, etc.) on a
third-party SaaS called Ambunet. If the company ever leaves Ambunet, the
only guaranteed way to get data out is a raw JSON export. The author is
one of the only technical people at the company, so this tool exists to
make that export usable by a non-technical staff member without needing
manual help under time pressure, if/when a departure ever becomes urgent.

We now have a real (demo) Ambunet export to work from: a zip containing
one JSON file per entity (118 of them - employees, shifts, incidents,
EPCRs, etc.), in MongoDB Extended JSON format (IDs as `{"$oid": ...}`,
dates as `{"$date": ...}`, etc.). `clean_data()` in `edms_databridge.py`
unwraps that into plain values before anything else touches it.

## Status

Accepts a zip, a folder, or a single JSON file (button click or drag-and-
drop). Each entity is classified as either tabular (flattened into its own
Excel sheet - still a **generic** flatten, not bespoke per-entity column
naming/ordering) or document-shaped (clinical/incident case records and
formal company documents like policies/SOPs get one PDF per record
instead, laid out with real sections and tables rather than a flattened
row). See `DOCUMENT_ENTITIES` in `edms_databridge.py` for the current
classification - it was judgment-called against the demo data and is
still pending a full review of the actual output against every entity.

Not yet tested as a built `.exe` on a clean machine (no dev tools/antivirus
false-positive check).

## Tech decisions

Python + Tkinter + pandas/openpyxl/reportlab, packaged as a single
unsigned `.exe` via PyInstaller (`--onefile --windowed`), chosen over
C#/.NET or Electron for speed of iteration given the author's background,
and because a single unsigned `.exe` is enough for an internal tool — no
installer needed.

Output is a mix of Excel (`.xlsx`, one sheet per tabular entity) and PDF
(one per record, for document-shaped entities) - see Status above and
`DOCUMENT_ENTITIES` in `edms_databridge.py`.

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

Unit tests cover the core, GUI-free logic: JSON loading/parsing (including
mixed file encodings), the generic flatten (`process_data()`), Extended
JSON unwrapping and sensitive-field redaction (`clean_data()`), multi-file
folder/zip ingestion (including filtering out macOS packaging junk), PDF
generation, error logging, and the version/update-check helpers. The
Tkinter GUI itself (widgets, dialogs, actual drag-and-drop) is exercised
manually, not by automated tests.

```
pytest
ruff check .
```

CI (GitHub Actions, `.github/workflows/ci.yml`) runs both on every push and
pull request against `main`, then does a smoke-test build of the exe with
PyInstaller and uploads it as a workflow artifact, so a working build is
always downloadable without needing a local Python setup. On `main`
specifically it can also cut a full GitHub Release - see Versioning &
releases below.

## Build the standalone .exe

```
build.bat
```

`build.bat` auto-activates `.venv` if it exists next to the script, so
just run it after the dev setup above. It produces
`dist\EDMSDataBridge.exe` — a single file with no dependencies. That's
the file to hand to end users. They just double-click it, no Python
install needed on their machine.

Note: the exe is unsigned, so Windows SmartScreen will show a warning on
first run ("Windows protected your PC"). Users click "More info" → "Run
anyway". A code-signing certificate would remove this, if it becomes worth
the cost for wider distribution.

## Versioning & releases

The app follows [semantic versioning](https://semver.org/) (MAJOR.MINOR.PATCH).
The single source of truth is the [VERSION](VERSION) file - everything
else derives from it:

- **In-app footer** reads it at runtime and shows e.g. "EDMS DataBridge
  v0.1.0".
- **`version_info.txt`** (the exe's Windows file-properties metadata) is
  generated from it by `generate_version_info.py` - don't hand-edit
  `version_info.txt`, it'll just get overwritten on the next build.
  `build.bat` and CI both regenerate it automatically before building.
- **GitHub Releases**: on every push to `main`, CI checks whether
  `VERSION` names a version that doesn't have a release yet. If it's new,
  CI builds the exe, tags the commit `vX.Y.Z`, and creates a GitHub
  Release with the exe attached - that's what the team should download
  from, rather than building it themselves. Most pushes don't bump the
  version, so most CI runs skip this step entirely.

**To cut a release**: bump the version in the `VERSION` file (following
semver - patch for fixes, minor for new features, major for breaking
changes) as part of your PR. Once that PR merges to `main`, the release
is created automatically within a few minutes.

The app also does a quiet, best-effort check on startup for whether a
newer release exists (via the GitHub API) and shows a small clickable
notice if so - it never blocks startup or shows anything if the check
fails (no network, GitHub unreachable, etc).

## Branding

The in-app header shows the official EDMS "ED" mark (`assets/logo.png`).
The title bar, taskbar, and the built exe's file icon use a different,
DataBridge-specific mark (`assets/logo.ico`): a single bold document
icon, representing the file this tool produces — see
[assets/README.md](assets/README.md) for provenance and generation
details.

## Contributing

`main` is protected: changes go through a pull request with CI passing,
rather than a direct push, even for the repo owner. Required approvals
is set to 0 rather than 1, since GitHub never allows a PR author to
approve their own PR — with a single collaborator, requiring 1 would
make every PR permanently unmergeable without an admin override.

## Project structure

```
├── edms_databridge.py     # main app (GUI + processing logic)
├── assets/                # optional logo.png / logo.ico (see assets/README.md)
├── tests/                 # pytest unit tests for the processing logic
├── requirements.txt       # runtime deps (bundled into the exe)
├── requirements-dev.txt   # runtime deps + pytest/ruff for local dev & CI
├── pyproject.toml         # pytest and ruff config
├── build.bat              # builds the standalone exe
├── VERSION                # single source of truth for the app's version
├── generate_version_info.py # generates version_info.txt from VERSION
├── version_info.txt       # Windows file-properties metadata for the exe
├── .github/workflows/     # CI: lint, test, and build-smoke-test on push/PR
├── .github/ISSUE_TEMPLATE/ # bug report / feature request forms
├── .gitignore
├── LICENSE
└── README.md
```

## Open questions / next steps

1. **Bespoke per-entity polish** — `process_data()` still does a generic
   flatten for every tabular entity: raw field names as column headers,
   no reordering, deeply-nested repeating sub-records (e.g. a vehicle's
   service history) flatten to numeric-indexed columns rather than a
   proper linked sheet. Worth doing per-entity once there's a reason to
   (see `DOCUMENT_ENTITIES`'s more polished treatment for the document
   entities as the template for what "worth it" looks like).
2. **Review the DOCUMENT_ENTITIES classification** — which of the 118
   entities get PDF treatment vs. an Excel sheet was judgment-called
   against the demo export, including a couple of outright guesses
   (`vdis`, `peaactions`). Needs a full pass against the actual generated
   output before it's trusted.
3. **Distribution/signing** — still unsigned, so Windows SmartScreen warns
   on first run. Is a code-signing certificate worth it, or is "click More
   info -> Run anyway" an acceptable one-time instruction for internal
   staff?
4. **Clean-machine testing** — the built `.exe` hasn't been tested on a
   machine without dev tools/antivirus false-positive checks yet.

"""
EDMS DataBridge
-----------------------
Author: Ash Kapow
Built for: EDMS

A simple Windows GUI tool: user uploads a JSON export (e.g. from Ambunet)
- a zip, a folder, or a single JSON file - and the app converts it into
clean Excel files (and PDFs for document-shaped records) they can
actually use.

Ambunet's real export is a zip (containing a folder of one *.json file
per entity, e.g. employees.json, incidents.json, epcrs.json) in MongoDB
Extended JSON format (IDs as {"$oid": ...}, dates as {"$date": ...}, etc).
`clean_data()` unwraps that into plain values and drops known-sensitive
fields (e.g. password hashes) before `process_data()` flattens each
entity into its own sheet (each saved as its own workbook).
`process_data()` is still a GENERIC flatten per entity, not bespoke
per-entity column mapping/renaming - see the README's open questions for
what's still deliberately deferred.

--- Build into a standalone .exe ---
Run build.bat (see that file for the exact pyinstaller command/flags).
The .exe will be in the generated dist/ folder. That single file is what
you hand to the non-technical user - no installer, no Python needed.
"""

import json
import os
import re
import sys
import threading
import traceback
import urllib.error
import urllib.request
import webbrowser
import zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import NamedTuple
from xml.sax.saxutils import escape

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from tkinterdnd2 import DND_FILES, TkinterDnD


APP_TITLE = "EDMS DataBridge"
GITHUB_REPO = "AshKapow/EDMS-DataBridge"
GITHUB_RELEASES_PAGE = f"https://github.com/{GITHUB_REPO}/releases/latest"


def resource_path(relative_path: str) -> Path:
    """Resolve a bundled asset path, in both dev mode and a PyInstaller onefile build."""
    base_path = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
    return base_path / relative_path


def load_version() -> str:
    """Read the app's own version from the bundled VERSION file. Falls
    back to "unknown" rather than raising - a missing version string
    shouldn't ever be the reason the app won't start."""
    try:
        return resource_path("VERSION").read_text(encoding="utf-8").strip()
    except OSError:
        return "unknown"


def parse_version(version: str) -> tuple:
    """"v1.2.3" or "1.2.3" -> (1, 2, 3), for comparing two semver strings.
    Non-numeric/malformed parts become 0 rather than raising, since this
    also has to handle whatever a GitHub release happens to be tagged."""
    parts = version.strip().lstrip("vV").split(".")
    parts = (parts + ["0", "0", "0"])[:3]

    def to_int(p):
        try:
            return int(p)
        except ValueError:
            return 0

    return tuple(to_int(p) for p in parts)


def get_latest_release_version(timeout: float = 3.0):
    """Ask GitHub for the latest release's tag name. Returns None on any
    failure at all (no network, GitHub unreachable, rate-limited, no
    releases yet, ...) - this is a best-effort courtesy check that must
    never be the reason the app fails to start or hangs."""
    url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
    try:
        request = urllib.request.Request(url, headers={"Accept": "application/vnd.github+json"})
        with urllib.request.urlopen(request, timeout=timeout) as response:
            data = json.loads(response.read().decode("utf-8"))
        tag = data.get("tag_name")
        return tag if isinstance(tag, str) and tag else None
    except (urllib.error.URLError, TimeoutError, ValueError, OSError):
        return None


def log_dir() -> Path:
    """Where error logs are written: %APPDATA%\\EDMSDataBridge on Windows,
    falling back to the user's home directory if APPDATA isn't set. This
    is a --windowed build with no console, so traceback.print_exc() alone
    goes nowhere if something crashes outside of a dev environment."""
    base = os.environ.get("APPDATA")
    return (Path(base) if base else Path.home()) / "EDMSDataBridge"


def log_error(context: str, exc: Exception) -> Path:
    """Append a timestamped entry (with the full traceback) to the error
    log, creating the log folder if needed. Returns the log file's path
    so the user can be told where to find/share it."""
    directory = log_dir()
    directory.mkdir(parents=True, exist_ok=True)
    log_path = directory / "EDMSDataBridge.log"
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"\n[{datetime.now().isoformat(timespec='seconds')}] {context}\n")
        f.write("".join(traceback.format_exception(type(exc), exc, exc.__traceback__)))
    return log_path


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


def _is_junk_zip_entry(name: str) -> bool:
    """
    Filter out filesystem/zip noise that isn't real export data: macOS
    adds a "._filename" AppleDouble shadow file for every real file when a
    folder is zipped on a Mac (mirroring the whole export 1-for-1, none of
    them valid JSON), usually inside a __MACOSX/ folder. These aren't a
    parse failure to report - they were never meant to be read at all.
    """
    parts = Path(name).parts
    return any(p == "__MACOSX" for p in parts) or Path(name).name.startswith("._")


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
        if _is_junk_zip_entry(str(json_file.relative_to(folder_path))):
            continue
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
            if not name.lower().endswith(".json") or _is_junk_zip_entry(name):
                continue
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
                if not records:
                    # ~40% of the entities in a real export have no records
                    # at all - a sheet of nothing is just noise to wade past.
                    continue
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


def save_sheets_as_workbooks(sheets: dict, output_dir) -> list:
    """Save each sheet as its own single-sheet workbook, in a folder of its
    own named for the entity (e.g. output_dir/Shifts/Shifts.xlsx) - the
    same one-folder-per-record-type layout the PDFs use, so the output
    root is one consistent list of folders. Returns the paths written."""
    output_dir = Path(output_dir)
    paths = []
    for sheet_name, df in sheets.items():
        name = sanitize_filename(entity_folder_name(sheet_name))
        folder = output_dir / name
        folder.mkdir(parents=True, exist_ok=True)
        path = folder / f"{name}.xlsx"
        save_as_excel({name[:31]: df}, str(path))  # Excel sheet name limit
        paths.append(path)
    return paths


# A4 page width (21cm) minus the 1.5cm left/right margins used in
# render_record_pdf() below - the usable content width for PDF layout.
PAGE_CONTENT_WIDTH = 18 * cm

# Entities that read as a single narrative/clinical/formal document, not a
# table of similar records - these get one PDF per record instead of an
# Excel sheet. Classified by hand against the real Ambunet demo export
# (see README open questions); not exhaustive/final - review and adjust
# as real usage turns up more/fewer entities that belong here.
DOCUMENT_ENTITIES = {
    # Clinical/incident case records - verified against the real demo data
    # to contain genuine narrative prose (descriptions, summaries, clinical
    # notes), not just metadata.
    "epcrs", "incidents", "cadincidents", "ptspatients", "ptsriskassessments",
    # People/HR narrative records - also verified: real interview notes,
    # a full whistleblowing description, etc.
    "employeeapplications", "speakupconcerns",
    # Governance record with substantial inline content (agenda items,
    # minutes, risk review notes) - verified, not just tracking metadata.
    "meetings",
    # "peaactions": confirmed clinical (PEA = Pulseless Electrical Activity,
    # a cardiac arrest rhythm - see the nested cardiacArrest field on
    # epcrs), so this is an incident-style case record, not a checklist.
    "peaactions",
    # Completed forms: each record is one filled-in checklist/audit, whose
    # answers are nested lists/dicts that flatten into dozens of unreadable
    # columns (78 for vehiclesafetychecks) or whole lists crammed into one
    # cell. As a PDF they read like the form that was filled in, which is
    # also what's wanted as CQC evidence.
    "vdis", "vehiclecleans", "vehiclesafetychecks", "audits", "medicineaudits",
    # Case-style records with free-text narrative and a follow-up trail.
    "patientfeedbacks",
    # Each event is really its event medical plan (ops plan, med plan,
    # risk assessment) - 110 columns flattened, one document as a PDF.
    "events",
    # The following have 0 records in the demo export, so couldn't be
    # directly verified - kept as documents on domain reasoning (the same
    # reasoning that held up for every entity above that WAS verifiable),
    # but worth a real check once real records exist:
    "paperpcrs", "medicalassessments", "occupationalhealths",
    "uninjuredreports", "imagingrequests", "appraisals", "complexdecisions",
}

# Moved OUT of DOCUMENT_ENTITIES after reviewing real generated PDFs:
# policies, policydescriptions, sops, pgds, coshhsheets, statementofpurposes
# looked like "formal documents" by name, but the actual records are thin
# acknowledgment/version-tracking metadata (who signed off, when, which
# version) referencing an EXTERNAL linked file (a "docLink"/S3 key) for the
# real document text - which this tool doesn't fetch. A PDF built from just
# that metadata isn't a useful "document", so these are tabular instead.
#
# Still tabular after review: crewvaults (its content is stored as raw
# HTML, which needs real HTML-to-PDF layout to be worth it), invoices and
# quotes (no stored invoice total, so a PDF would look official but be
# incomplete - and finance will want to filter/sum them anyway).

# PDF document titles for each entity in DOCUMENT_ENTITIES. These are raw
# lowercase filename stems (e.g. "employeeapplications"), not camelCase,
# so humanize_field_name() can't recover word boundaries automatically -
# hence an explicit mapping rather than a generic split. Falls back to
# humanize_field_name() for anything not listed (see entity_display_name()).
ENTITY_DISPLAY_NAMES = {
    "epcrs": "EPCR",
    "paperpcrs": "Paper PCR",
    "incidents": "Incident Report",
    "cadincidents": "CAD Incident",
    "ptspatients": "PTS Patient Record",
    "medicalassessments": "Medical Assessment",
    "occupationalhealths": "Occupational Health Record",
    "ptsriskassessments": "PTS Risk Assessment",
    "uninjuredreports": "Uninjured Person Report",
    "imagingrequests": "Imaging Request",
    "appraisals": "Employee Appraisal",
    "employeeapplications": "Employee Application",
    "speakupconcerns": "Speak Up Concern",
    "complexdecisions": "Complex Decision Record",
    "meetings": "Meeting Minutes",
    "peaactions": "PEA Action (Pulseless Electrical Activity)",
    "vdis": "Vehicle Daily Inspection",
    "vehiclecleans": "Vehicle Clean Record",
    "vehiclesafetychecks": "Vehicle Safety Check",
    "audits": "Audit",
    "medicineaudits": "Medicine Audit",
    "patientfeedbacks": "Patient Feedback",
    "events": "Event Plan",
}


def entity_display_name(entity_name: str) -> str:
    """Human-readable title for a document entity's PDFs. See
    ENTITY_DISPLAY_NAMES; falls back to humanize_field_name() so an
    unmapped entity still gets *something* readable rather than raising."""
    return ENTITY_DISPLAY_NAMES.get(entity_name, humanize_field_name(entity_name))


# Field names that read fine expanded to their acronym instead of Title
# Case, e.g. "nhsNumber" -> "NHS Number" not "Nhs Number".
_ACRONYM_FIELD_WORDS = {
    "nhs", "gp", "dob", "id", "cqc", "dbs", "pgd", "sop", "cad", "pts",
    "gcs", "bp", "spo2", "etco2", "bgl", "pefr", "avpu", "pmhx", "epcr",
    "vin", "mot",
}


def humanize_field_name(key: str) -> str:
    """Turn a camelCase/snake_case JSON field name into a human-readable
    label, e.g. "firstName" -> "First Name", "nhsNumber" -> "NHS Number".
    Meant for a non-technical reader - raw field names are not."""
    if " " in key:
        # Already human-written, e.g. a checklist item used as a key like
        # "Front Brake Pads (Wear/Condition)" - re-casing would only mangle it.
        return key
    spaced = re.sub(r"(?<!^)(?=[A-Z])", " ", key).replace("_", " ")
    words = [w for w in spaced.split(" ") if w]
    return " ".join(
        w.upper() if w.lower() in _ACRONYM_FIELD_WORDS else w.capitalize()
        for w in words
    )


_INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*]')


def sanitize_filename(name: str) -> str:
    """Make a string safe to use as a Windows filename."""
    cleaned = _INVALID_FILENAME_CHARS.sub("_", name).strip(" ._")
    return cleaned[:120] or "record"


class PdfNaming(NamedTuple):
    """How one document entity's PDFs are filed and named, e.g.
    CAD Incidents/2026/09 - September/2026-09-11 CAD 1109262001.pdf.

    date_field (a dotted path into the record) is when the thing actually
    happened - not createdAt where there's something better, since a
    record is often written up after the fact - and drives both the
    Year/Month folders and the filename's leading date, so Explorer sorts
    chronologically. None means the record isn't an event (e.g. a patient)
    and is filed flat. ref_field is the Ambunet reference, so a PDF can be
    matched back to its record; None where there isn't one.

    Deliberately never a person's name in the filename: filenames surface
    in Windows search, recent files, OneDrive and email attachment names,
    which is the wrong place for patient identities.
    """
    folder: str
    prefix: str
    date_field: str | None = "createdAt"
    ref_field: str | None = None


PDF_NAMING = {
    "cadincidents": PdfNaming("CAD Incidents", "CAD", "createdAt", "incidentNumber"),
    "incidents": PdfNaming("Incident Reports", "Incident", "incidentDate", "incidentNumber"),
    "epcrs": PdfNaming("ePCRs", "ePCR", "incident.incidentDate", "epcrNumber"),
    "meetings": PdfNaming("Meetings", "Meeting", "startTime", "title"),
    "speakupconcerns": PdfNaming("Speak Up Concerns", "Speak Up", "createdAt", "referenceNumber"),
    "ptsriskassessments": PdfNaming(
        "PTS Risk Assessments", "PTS Risk Assessment", "visitDate", "assessmentType"
    ),
    "ptspatients": PdfNaming("PTS Patients", "PTS Patient", None, "patientID"),
    "employeeapplications": PdfNaming("Employee Applications", "Employee Application"),
    "events": PdfNaming("Events", "Event", "startDate", "title"),
    "vdis": PdfNaming("Vehicle Daily Inspections", "VDI", "date"),
    "vehiclecleans": PdfNaming("Vehicle Cleans", "Vehicle Clean", "start", "cleanType"),
    "vehiclesafetychecks": PdfNaming("Vehicle Safety Checks", "Vehicle Safety Check", "testDate"),
    "audits": PdfNaming("Audits", "Audit", "date"),
    "medicineaudits": PdfNaming("Medicine Audits", "Medicine Audit", "createdAt", "originalTag"),
    "patientfeedbacks": PdfNaming("Patient Feedback", "Patient Feedback"),
    # 0 records in the demo export, so only createdAt to go on for now.
    "paperpcrs": PdfNaming("Paper PCRs", "Paper PCR"),
    "medicalassessments": PdfNaming("Medical Assessments", "Medical Assessment"),
    "occupationalhealths": PdfNaming("Occupational Health Records", "Occupational Health"),
    "uninjuredreports": PdfNaming("Uninjured Person Reports", "Uninjured Person Report"),
    "imagingrequests": PdfNaming("Imaging Requests", "Imaging Request"),
    "appraisals": PdfNaming("Appraisals", "Appraisal"),
    "complexdecisions": PdfNaming("Complex Decisions", "Complex Decision"),
    "peaactions": PdfNaming("PEA Actions", "PEA Action"),
}

# Output folder (and workbook) names for the tabular entities - raw
# lowercase filename stems like "mandatorytrainings" have no word
# boundaries for humanize_field_name() to recover. Covers every entity in
# the demo export; anything new falls back to humanize_field_name().
TABLE_FOLDER_NAMES = {
    "announcements": "Announcements",
    "assets": "Assets",
    "auditdescriptions": "Audit Templates",
    "audittrails": "Audit Trails",
    "baselocations": "Base Locations",
    "cadcalltriagecomplaints": "CAD Call Triage Complaints",
    "cadresources": "CAD Resources",
    "cals": "CALs",
    "commandlogs": "Command Logs",
    "comments": "Comments",
    "companysettings": "Company Settings",
    "contacts": "Contacts",
    "controlleddrugaudits": "Controlled Drug Audits",
    "controlrooms": "Control Rooms",
    "coshhsheets": "COSHH Sheets",
    "counters": "Counters",
    "cpdlogs": "CPD Logs",
    "cqccomplianceresponses": "CQC Compliance Responses",
    "cqcevidenceitems": "CQC Evidence Items",
    "cqcgaps": "CQC Gaps",
    "crewvaults": "Crew Vault",
    "dbsactions": "DBS Actions",
    "dbschecks": "DBS Checks",
    "defects": "Defects",
    "documents": "Documents",
    "drivinglicencechecks": "Driving Licence Checks",
    "employees": "Employees",
    "emptemplates": "EMP Templates",
    "epcraudits": "ePCR Audits",
    "epcrsnapshots": "ePCR Snapshots",
    "eventcqcexposureassessments": "Event CQC Exposure Assessments",
    "eventexpenses": "Event Expenses",
    "expenses": "Expenses",
    "governanceactions": "Governance Actions",
    "hospitals": "Hospitals",
    "idcards": "ID Cards",
    "incidentwitnessrequests": "Incident Witness Requests",
    "invoices": "Invoices",
    "logs": "Logs",
    "maintenancelogs": "Maintenance Logs",
    "majortraumanetworks": "Major Trauma Networks",
    "makereadyinventories": "Make Ready Inventories",
    "makereadyloadlists": "Make Ready Load Lists",
    "makereadyreports": "Make Ready Reports",
    "mandatorytrainingcourses": "Mandatory Training Courses",
    "mandatorytrainings": "Mandatory Training Records",
    "medicinedescriptions": "Medicine Descriptions",
    "medicinelocations": "Medicine Locations",
    "medicinelogs": "Medicine Logs",
    "medicines": "Medicines",
    "mots": "MOTs",
    "organisations": "Organisations",
    "patientfeedbackinvites": "Patient Feedback Invites",
    "payruns": "Pay Runs",
    "pcraccesslogs": "PCR Access Logs",
    "pcrsharecodes": "PCR Share Codes",
    "pgds": "PGDs",
    "policies": "Policies",
    "policydescriptions": "Policy Descriptions",
    "ptsbookings": "PTS Bookings",
    "ptscontracts": "PTS Contracts",
    "ptsdispatcherpresences": "PTS Dispatcher Presences",
    "ptsduties": "PTS Duties",
    "ptsjourneys": "PTS Journeys",
    "ptsportalchatrequests": "PTS Portal Chat Requests",
    "ptsquotes": "PTS Quotes",
    "ptstariffs": "PTS Tariffs",
    "qualifications": "Qualifications",
    "quotes": "Quotes",
    "regulatedactivities": "Regulated Activities",
    "resourcechatmessages": "Resource Chat Messages",
    "restockclaims": "Restock Claims",
    "risks": "Risk Register",
    "rosters": "Rosters",
    "safetyalerts": "Safety Alerts",
    "shiftapplications": "Shift Applications",
    "shifts": "Shifts",
    "smslogs": "SMS Logs",
    "sops": "SOPs",
    "staffavailabilities": "Staff Availability",
    "staffpayinvoices": "Staff Pay Invoices",
    "statementofpurposes": "Statements of Purpose",
    "stockcatalogitems": "Stock Catalogue Items",
    "subcontractors": "Subcontractors",
    "tasks": "Tasks",
    "timetrackers": "Time Trackers",
    "treatmentassignments": "Treatment Assignments",
    "treatmentlocations": "Treatment Locations",
    "treatmentzones": "Treatment Zones",
    "usernotifications": "User Notifications",
    "users": "Users",
    "vehicledeployments": "Vehicle Deployments",
    "vehicles": "Vehicles",
    "warddepartments": "Ward Departments",
    "websitequeries": "Website Queries",
}


def entity_folder_name(entity_name: str) -> str:
    """The output folder for an entity - PDFs and spreadsheets alike sit in
    one flat list of record-type folders at the output root."""
    if entity_name in DOCUMENT_ENTITIES or entity_name in PDF_NAMING:
        return pdf_naming(entity_name).folder
    return TABLE_FOLDER_NAMES.get(entity_name) or humanize_field_name(entity_name)


def pdf_naming(entity_name: str) -> PdfNaming:
    """PDF_NAMING's entry for an entity, or a generic createdAt-dated one
    for a document entity with no records yet to design a better one from."""
    if entity_name in PDF_NAMING:
        return PDF_NAMING[entity_name]
    display = entity_display_name(entity_name)
    return PdfNaming(display, display, "createdAt", None)


def _get_path(record: dict, dotted: str):
    value = record
    for part in dotted.split("."):
        if not isinstance(value, dict):
            return None
        value = value.get(part)
    return value


def record_date(record: dict, dotted: str):
    """Parse an ISO date field (as clean_data() leaves them) into local
    time - Ambunet stores UTC, and a 00:30 BST call belongs on the local
    day, matching the date baked into its own incident number. None if
    missing or unparseable."""
    value = _get_path(record, dotted)
    if not isinstance(value, str) or not value:
        return None
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None
    return parsed.astimezone() if parsed.tzinfo else parsed


def record_reference(entity_name: str, record: dict, dotted: str):
    """The record's Ambunet reference as clean text, or None."""
    value = _get_path(record, dotted)
    if isinstance(value, bool) or not isinstance(value, (str, int, float)) or value == "":
        return None
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    ref = str(value).strip()
    if entity_name == "cadincidents" and ref.isdigit():
        # CAD numbers are DDMMYY + a 4-digit sequence, but Ambunet exports
        # them as bare numbers, so a date before the 10th loses its
        # leading zero (0409261004 -> 409261004). Put it back.
        ref = ref.zfill(10)
    return ref or None


def pdf_path_for(entity_name: str, record: dict):
    """Where (relative to the output folder) one record's PDF goes, as
    (folder, filename stem) - see PdfNaming. Collisions are resolved by
    the caller."""
    naming = pdf_naming(entity_name)
    folder = Path(sanitize_filename(naming.folder))
    stem = naming.prefix
    if naming.ref_field:
        ref = record_reference(entity_name, record, naming.ref_field)
        if ref:
            stem = f"{stem} {ref}"
    if naming.date_field:
        when = record_date(record, naming.date_field)
        if when:
            folder = folder / f"{when:%Y}" / f"{when:%m - %B}"
            stem = f"{when:%Y-%m-%d} {stem}"
        else:
            folder = folder / "Undated"
    return folder, sanitize_filename(stem)


def _is_empty(value) -> bool:
    return value is None or value in ("", [], {})


def _cell_text(value) -> str:
    """Render a table cell's value as plain text - used for nested list-of-
    dict fields (e.g. vitals.obs), where a cell can itself hold a dict."""
    if _is_empty(value):
        return ""
    if isinstance(value, dict):
        return "; ".join(
            f"{humanize_field_name(k)}: {_cell_text(v)}"
            for k, v in value.items()
            if not _is_empty(v)
        )
    if isinstance(value, list):
        return ", ".join(_cell_text(v) for v in value)
    if isinstance(value, bool):
        return "Yes" if value else "No"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))  # e.g. Ambunet's 2.208261002E+09 incident numbers
    return str(value)


def _para(text: str, style, bold: bool = False) -> Paragraph:
    """A Paragraph of literal text. reportlab parses Paragraph text as
    markup, so raw record text containing anything tag-like (e.g. the HTML
    stored in crewvaults.content) would otherwise crash the whole run."""
    text = escape(text)
    return Paragraph(f"<b>{text}</b>" if bold else text, style)


def _details_table(rows, styles) -> Table:
    """A two-column Label / Value table for a section's scalar fields."""
    data = [
        [
            _para(label, styles["Normal"], bold=True),
            _para(_cell_text(value), styles["Normal"]),
        ]
        for label, value in rows
    ]
    table = Table(data, colWidths=[5 * cm, PAGE_CONTENT_WIDTH - 5 * cm])
    table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.HexColor("#dddddd")),
    ]))
    return table


_MIN_READABLE_COLUMN_WIDTH = 2.5 * cm


def _collect_keys(records: list) -> list:
    keys = []
    for record in records:
        for key in record.keys():
            # A nested row's own database ID means nothing to a reader.
            if key != "_id" and key not in keys:
                keys.append(key)
    return keys


def _as_item_rows(value: dict):
    """If a dict is really a keyed list of same-shaped flat records - e.g. a
    vehicle safety check's {"Brake Fluid Level": {"safe": ..., "advised":
    ..., "comments": ...}, ...} - return it as rows for _records_table()
    with the key as an "item" column. Otherwise None, and it stays a
    subsection per key (one heading per checklist item is unreadable)."""
    if len(value) < 2 or not all(isinstance(v, dict) and v for v in value.values()):
        return None
    shapes = {tuple(v.keys()) for v in value.values()}
    flat = all(not isinstance(x, (dict, list)) for v in value.values() for x in v.values())
    if len(shapes) != 1 or not flat:
        return None
    return [{"item": k, **v} for k, v in value.items()]


def _records_table(records: list, styles) -> Table:
    """A table for a nested list of records (e.g. vitals.obs, an EPCR's
    time-series of observations, or a vehicle's service history).

    Normally one row per record, one column per field. But a vitals-style
    reading can have 20+ fields for just 2-3 readings - fields as columns
    there means every column is too narrow to hold even its own header
    without wrapping one letter per line. When there are clearly more
    fields than records, it reads far better transposed instead: one row
    per field, one column per record.
    """
    keys = _collect_keys(records)
    if not keys:
        return Table([[""]])

    fits_normally = len(keys) * _MIN_READABLE_COLUMN_WIDTH <= PAGE_CONTENT_WIDTH
    if fits_normally or len(keys) <= len(records):
        return _regular_records_table(records, keys, styles)
    return _transposed_records_table(records, keys, styles)


def _regular_records_table(records: list, keys: list, styles) -> Table:
    header = [_para(humanize_field_name(k), styles["Normal"], bold=True) for k in keys]
    rows = [header]
    for record in records:
        rows.append([_para(_cell_text(record.get(k)), styles["Normal"]) for k in keys])

    # Explicit equal-width columns rather than reportlab's auto-sizing: a
    # field with many distinct keys across records (seen in the real
    # data: 21 columns in one case) can ask for more width than the page
    # has, which crashes the auto-sizer outright rather than shrinking.
    col_width = PAGE_CONTENT_WIDTH / len(keys)
    table = Table(rows, colWidths=[col_width] * len(keys), repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e32e27")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#dddddd")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
    ]))
    return table


def _transposed_column_label(record: dict, index: int) -> str:
    """Label a transposed table's per-record column - prefer a time/date
    field if the record has one (readings are usually time-stamped), else
    just number them."""
    for key, value in record.items():
        if key.lower() in ("time", "date", "datetime") and not _is_empty(value):
            return _cell_text(value)
    return f"#{index + 1}"


def _transposed_records_table(records: list, keys: list, styles) -> Table:
    header = [_para("Field", styles["Normal"], bold=True)] + [
        _para(_transposed_column_label(r, i), styles["Normal"], bold=True)
        for i, r in enumerate(records)
    ]
    rows = [header]
    for key in keys:
        row = [_para(humanize_field_name(key), styles["Normal"], bold=True)]
        row += [_para(_cell_text(r.get(key)), styles["Normal"]) for r in records]
        rows.append(row)

    label_col = 4 * cm
    value_col = max((PAGE_CONTENT_WIDTH - label_col) / len(records), 1.5 * cm)
    table = Table(rows, colWidths=[label_col] + [value_col] * len(records), repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e32e27")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, 1), (0, -1), colors.HexColor("#f2f2f2")),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#dddddd")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
    ]))
    return table


def _build_flowables(data: dict, styles, heading_style: str) -> list:
    """Recursively turn a dict into reportlab flowables: nested dicts become
    subsections, lists of dicts become tables, everything else becomes a
    Label: Value row grouped into one details table per section."""
    next_heading = {"Heading2": "Heading3", "Heading3": "Heading4"}.get(heading_style, "Heading4")
    flowables = []
    detail_rows = []

    def flush_details():
        if detail_rows:
            flowables.append(_details_table(list(detail_rows), styles))
            detail_rows.clear()

    for key, value in data.items():
        if _is_empty(value):
            continue
        label = humanize_field_name(key)
        item_rows = _as_item_rows(value) if isinstance(value, dict) else None
        if item_rows:
            flush_details()
            flowables.append(_para(label, styles[heading_style]))
            flowables.append(_records_table(item_rows, styles))
            flowables.append(Spacer(1, 0.3 * cm))
        elif isinstance(value, dict):
            flush_details()
            flowables.append(_para(label, styles[heading_style]))
            flowables.extend(_build_flowables(value, styles, next_heading))
        elif isinstance(value, list) and isinstance(value[0], dict):
            flush_details()
            flowables.append(_para(label, styles[heading_style]))
            flowables.append(_records_table(value, styles))
            flowables.append(Spacer(1, 0.3 * cm))
        else:
            detail_rows.append((label, value))
    flush_details()
    return flowables


def _draw_letterhead(canvas, doc):
    """Draws the EDMS logo in the top-right corner - reportlab calls this
    once per page via onFirstPage/onLaterPages, so it's on every page of a
    multi-page document, not just the first."""
    logo_path = resource_path("assets/logo.png")
    if not logo_path.exists():
        return
    canvas.saveState()
    logo_size = 1.4 * cm
    x = doc.pagesize[0] - doc.rightMargin - logo_size
    y = doc.pagesize[1] - doc.topMargin - logo_size + 0.4 * cm
    canvas.drawImage(
        str(logo_path), x, y, width=logo_size, height=logo_size,
        preserveAspectRatio=True, mask="auto",
    )
    canvas.restoreState()


def render_record_pdf(entity_name: str, record: dict, output_path):
    """Render one record as a PDF laid out like a real document - an EDMS-
    branded letterhead, a title, a small reference line (ID/created/
    updated), then a section per top-level field: nested objects become
    subsections, nested lists of records become tables, everything else
    becomes Label: Value rows."""
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(
        str(output_path), pagesize=A4,
        topMargin=1.5 * cm, bottomMargin=1.5 * cm, leftMargin=1.5 * cm, rightMargin=1.5 * cm,
    )

    story = [_para(entity_display_name(entity_name), styles["Title"])]

    reference_bits = [
        f"{humanize_field_name(key)}: {record[key]}"
        for key in ("_id", "createdAt", "updatedAt")
        if not _is_empty(record.get(key))
    ]
    if reference_bits:
        story.append(_para(" | ".join(reference_bits), styles["Normal"]))
    story.append(Spacer(1, 0.5 * cm))

    body = {k: v for k, v in record.items() if k not in ("_id", "__v", "createdAt", "updatedAt")}
    story.extend(_build_flowables(body, styles, "Heading2"))

    doc.build(story, onFirstPage=_draw_letterhead, onLaterPages=_draw_letterhead)


def generate_pdfs(data: dict, output_dir) -> dict:
    """For every entity in DOCUMENT_ENTITIES present in `data`, render one
    PDF per record under output_dir, filed and named per PDF_NAMING.
    Returns {entity: count} for entities that actually produced any PDFs."""
    output_dir = Path(output_dir)
    counts = {}
    for entity, records in data.items():
        if entity not in DOCUMENT_ENTITIES or not isinstance(records, list) or not records:
            continue
        ref_field = pdf_naming(entity).ref_field
        made = 0
        for record in records:
            if not isinstance(record, dict):
                continue
            if ref_field and "." not in ref_field and ref_field in record:
                # Show the same cleaned-up reference inside the PDF as in its
                # filename, not e.g. a raw 2208261002.0 or a dropped leading 0.
                ref = record_reference(entity, record, ref_field)
                if ref:
                    record = {**record, ref_field: ref}
            folder, stem = pdf_path_for(entity, record)
            pdf_dir = output_dir / folder
            pdf_dir.mkdir(parents=True, exist_ok=True)
            pdf_path = pdf_dir / f"{stem}.pdf"
            n = 1
            while pdf_path.exists():
                n += 1
                pdf_path = pdf_dir / f"{stem} ({n}).pdf"
            render_record_pdf(entity, record, pdf_path)
            made += 1
        if made:
            counts[entity] = made
    return counts


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

        self._version = load_version()

        ttk.Label(
            self,
            text=f"EDMS DataBridge v{self._version} · Built for EDMS by Ash Kapow",
            font=("Segoe UI", 8),
            foreground="gray50",
        ).pack(side="bottom", pady=(0, 10))

        self.update_notice_label = ttk.Label(
            self, text="", foreground="#1a5fb4", cursor="hand2",
            wraplength=self.WINDOW_WIDTH - 40, justify="center",
        )
        self.update_notice_label.bind("<Button-1>", lambda e: webbrowser.open(GITHUB_RELEASES_PAGE))
        # Not packed yet - stays invisible/zero-height until an update is
        # actually found, so nothing shifts on startup in the normal case.

        self._fit_window_to_content()
        self._check_for_update_async()

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

    def _check_for_update_async(self):
        """Best-effort, non-blocking check for a newer release. Runs on a
        background thread so a slow/unreachable network can never delay
        startup; silently does nothing if the check fails or finds
        nothing newer (see get_latest_release_version())."""

        def worker():
            latest = get_latest_release_version()
            if latest and parse_version(latest) > parse_version(self._version):
                self.after(0, self._show_update_notice, latest)

        threading.Thread(target=worker, daemon=True).start()

    def _show_update_notice(self, latest_version):
        self.update_notice_label.config(
            text=f"A newer version ({latest_version}) is available — click to download"
        )
        self.update_notice_label.pack(side="bottom", pady=(0, 4))
        self._fit_window_to_content()

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

            # Cleaned once here (rather than only inside process_data())
            # since PDF generation needs the same unwrapped/redacted data.
            data = clean_data(data)
            empty_count = 0
            if isinstance(data, dict):
                empty_count = sum(1 for v in data.values() if isinstance(v, list) and not v)
                pdf_data = {
                    k: v for k, v in data.items() if k in DOCUMENT_ENTITIES and isinstance(v, list)
                }
                tab_data = {k: v for k, v in data.items() if k not in pdf_data}
                pdf_data = {k: v for k, v in pdf_data.items() if v}
            else:
                pdf_data, tab_data = {}, data

            sheets = process_data(tab_data)
            if not sheets and not pdf_data:
                self._set_status("")
                messagebox.showwarning(APP_TITLE, "No records were found in that export.")
                return

            if pdf_data or len(sheets) > 1:
                # More than one output file, so the output is a folder
                # (one subfolder per record type), not a single xlsx file.
                output_dir = filedialog.askdirectory(
                    title="Choose a folder to save the formatted output",
                    initialdir=str(default_dir),
                )
                if not output_dir:
                    self._set_status("Cancelled.")
                    return
                output_dir = Path(output_dir)
                save_sheets_as_workbooks(sheets, output_dir)
                pdf_counts = generate_pdfs(pdf_data, output_dir)
                result_location = str(output_dir)
            else:
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
                pdf_counts = {}
                result_location = output_path

            self._set_status(f"Done! Saved to:\n{result_location}")
            success_message = f"Success! Your formatted output is ready:\n\n{result_location}"
            if len(sheets) > 1:
                success_message += f"\n\n{len(sheets)} spreadsheet(s) created."
            if pdf_counts:
                total_pdfs = sum(pdf_counts.values())
                pdf_lines = "\n".join(
                    f"  - {entity_display_name(e)}: {c}" for e, c in pdf_counts.items()
                )
                success_message += f"\n\n{total_pdfs} PDF document(s) also created:\n{pdf_lines}"
            if empty_count:
                success_message += (
                    f"\n\n{empty_count} categor{'y' if empty_count == 1 else 'ies'} "
                    "had no records, so no file was made for them."
                )
            if skipped:
                shown = [f"  - {name}" for name, _ in skipped[:5]]
                if len(skipped) > 5:
                    shown.append(f"  ...and {len(skipped) - 5} more")
                success_message += (
                    f"\n\n{len(skipped)} file(s) couldn't be read and were skipped:\n"
                    + "\n".join(shown)
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
            log_path = log_error(f"Processing {path}", e)
            messagebox.showerror(
                APP_TITLE,
                f"Something went wrong:\n\n{e}\n\n"
                f"Details were saved to:\n{log_path}\n"
                "You can share this file if you need help.",
            )
            try:
                os.startfile(log_path.parent)
            except OSError:
                pass


if __name__ == "__main__":
    app = App()
    app.mainloop()

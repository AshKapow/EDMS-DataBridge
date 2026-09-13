import json
import tkinter as tk
import zipfile
from pathlib import Path

import pytest

from edms_databridge import (
    DOCUMENT_ENTITIES,
    ENTITY_DISPLAY_NAMES,
    clean_data,
    entity_display_name,
    generate_pdfs,
    get_latest_release_version,
    humanize_field_name,
    load_entity_folder,
    load_entity_zip,
    load_json,
    load_logo_image,
    load_version,
    log_error,
    parse_dnd_filepaths,
    parse_version,
    process_data,
    record_label,
    redact_sensitive_fields,
    resource_path,
    sanitize_filename,
    unwrap_extended_json,
)


def test_load_json_valid(tmp_path):
    path = tmp_path / "data.json"
    path.write_text('{"a": 1}', encoding="utf-8")
    assert load_json(str(path)) == {"a": 1}


def test_load_json_handles_bom(tmp_path):
    path = tmp_path / "data.json"
    path.write_bytes('{"a": 1}'.encode("utf-8-sig"))
    assert load_json(str(path)) == {"a": 1}


def test_load_json_handles_cp1252_encoding(tmp_path):
    # Regression test: a real file in the Ambunet demo export was encoded
    # as Windows-1252, not UTF-8 - a "£" in an expense field decoded fine
    # in cp1252 but raised UnicodeDecodeError as UTF-8 (0xA3 isn't a valid
    # UTF-8 start byte).
    path = tmp_path / "data.json"
    path.write_bytes('{"cost": "£50"}'.encode("cp1252"))
    assert load_json(str(path)) == {"cost": "£50"}


def test_load_json_invalid_raises(tmp_path):
    path = tmp_path / "data.json"
    path.write_text("not json", encoding="utf-8")
    with pytest.raises(json.JSONDecodeError):
        load_json(str(path))


def test_process_data_list_of_records():
    data = [{"a": 1, "b": 2}, {"a": 3, "b": 4}]
    sheets = process_data(data)
    assert list(sheets.keys()) == ["Data"]
    assert sheets["Data"].shape == (2, 2)
    assert list(sheets["Data"].columns) == ["a", "b"]


def test_process_data_dict_of_lists_one_sheet_per_key():
    data = {
        "patients": [{"id": 1, "name": "Alice"}],
        "shifts": [{"id": 10, "date": "2024-01-01"}],
    }
    sheets = process_data(data)
    assert set(sheets.keys()) == {"patients", "shifts"}
    assert len(sheets["patients"]) == 1
    assert len(sheets["shifts"]) == 1


def test_process_data_dict_with_no_list_keys_is_single_row():
    data = {"a": 1, "b": {"c": 2}}
    sheets = process_data(data)
    assert list(sheets.keys()) == ["Data"]
    assert len(sheets["Data"]) == 1
    assert "b.c" in sheets["Data"].columns


def test_process_data_ignores_non_list_keys_when_list_keys_present():
    data = {"patients": [{"id": 1}], "meta": "exported 2024-01-01"}
    sheets = process_data(data)
    assert set(sheets.keys()) == {"patients"}


def test_process_data_normalizes_nested_objects_in_lists():
    data = [{"id": 1, "address": {"city": "Springfield", "zip": "11111"}}]
    sheets = process_data(data)
    assert "address.city" in sheets["Data"].columns


def test_process_data_truncates_long_sheet_names_to_excel_limit():
    long_key = "a" * 40
    data = {long_key: [{"x": 1}]}
    sheets = process_data(data)
    sheet_name = list(sheets.keys())[0]
    assert sheet_name == long_key[:31]
    assert len(sheet_name) == 31


@pytest.mark.parametrize("bad_data", ["just a string", 42, None])
def test_process_data_rejects_non_list_non_dict_input(bad_data):
    with pytest.raises(ValueError):
        process_data(bad_data)


def test_resource_path_resolves_relative_to_script_dir_in_dev_mode():
    path = resource_path("assets/logo.png")
    assert path == Path(__file__).parent.parent / "assets" / "logo.png"


def test_load_logo_image_returns_none_when_file_missing(monkeypatch):
    monkeypatch.setattr(
        "edms_databridge.resource_path", lambda relative_path: Path("no/such/file.png")
    )
    assert load_logo_image() is None


def test_load_logo_image_loads_the_real_asset():
    # load_logo_image() is only ever called after App's Tk root exists (see
    # App.__init__), so a root is created here to match that precondition.
    root = tk.Tk()
    try:
        image = load_logo_image()
        assert image is not None
        assert isinstance(image, tk.PhotoImage)
    finally:
        root.destroy()


def test_parse_dnd_filepaths_single_simple_path():
    assert parse_dnd_filepaths("C:/Users/ashley/export.json") == [
        "C:/Users/ashley/export.json"
    ]


def test_parse_dnd_filepaths_path_with_spaces_in_braces():
    assert parse_dnd_filepaths("{C:/Users/ashley/My Exports/export.json}") == [
        "C:/Users/ashley/My Exports/export.json"
    ]


def test_parse_dnd_filepaths_multiple_paths():
    data = "{C:/My Exports/a.json} C:/b.json"
    assert parse_dnd_filepaths(data) == ["C:/My Exports/a.json", "C:/b.json"]


def test_parse_dnd_filepaths_empty_string():
    assert parse_dnd_filepaths("") == []


def test_unwrap_extended_json_oid():
    assert unwrap_extended_json({"$oid": "abc123"}) == "abc123"


def test_unwrap_extended_json_date_string():
    assert unwrap_extended_json({"$date": "2024-01-01T00:00:00Z"}) == "2024-01-01T00:00:00Z"


def test_unwrap_extended_json_date_as_epoch_millis():
    # 2024-01-01T00:00:00Z in epoch milliseconds
    result = unwrap_extended_json({"$date": {"$numberLong": "1704067200000"}})
    assert result == "2024-01-01T00:00:00+00:00"


def test_unwrap_extended_json_date_before_1970():
    # Regression test: datetime.fromtimestamp() raises OSError on Windows
    # for negative timestamps - this hit real date-of-birth fields in the
    # actual Ambunet demo export. -631152000000ms = 1950-01-01T00:00:00Z.
    result = unwrap_extended_json({"$date": {"$numberLong": "-631152000000"}})
    assert result == "1950-01-01T00:00:00+00:00"


def test_unwrap_extended_json_numbers():
    assert unwrap_extended_json({"$numberLong": "123"}) == 123
    assert unwrap_extended_json({"$numberInt": "5"}) == 5
    assert unwrap_extended_json({"$numberDouble": "1.5"}) == 1.5
    assert unwrap_extended_json({"$numberDecimal": "2.5"}) == 2.5


def test_unwrap_extended_json_recurses_into_nested_structures():
    data = {"employee": {"_id": {"$oid": "e1"}}, "logs": [{"_id": {"$oid": "l1"}}]}
    result = unwrap_extended_json(data)
    assert result == {"employee": {"_id": "e1"}, "logs": [{"_id": "l1"}]}


def test_unwrap_extended_json_leaves_plain_values_alone():
    data = {"a": 1, "b": "text", "c": [1, 2], "d": None}
    assert unwrap_extended_json(data) == data


def test_redact_sensitive_fields_drops_password():
    data = {"email": "a@b.com", "security": {"password": "hashed-value"}}
    result = redact_sensitive_fields(data)
    assert result == {"email": "a@b.com", "security": {}}


def test_redact_sensitive_fields_is_case_insensitive():
    assert redact_sensitive_fields({"Password": "x"}) == {}


def test_redact_sensitive_fields_recurses_into_lists():
    data = [{"token": "secret"}, {"name": "keep me"}]
    assert redact_sensitive_fields(data) == [{}, {"name": "keep me"}]


def test_clean_data_unwraps_and_redacts_together():
    data = {"_id": {"$oid": "e1"}, "security": {"password": "hashed"}, "name": "Alice"}
    assert clean_data(data) == {"_id": "e1", "security": {}, "name": "Alice"}


def test_process_data_cleans_extended_json_and_sensitive_fields():
    # Shape matching the real Ambunet export: a dict of {entity: [records]}
    # with Mongo Extended JSON types and a sensitive field mixed in.
    data = {
        "users": [
            {
                "_id": {"$oid": "u1"},
                "email": "a@b.com",
                "security": {"password": "hashed"},
            }
        ]
    }
    sheets = process_data(data)
    df = sheets["users"]
    assert df.loc[0, "_id"] == "u1"
    assert "security.password" not in df.columns


def test_load_entity_folder_keys_by_filename_stem(tmp_path):
    (tmp_path / "employees.json").write_text('[{"name": "Alice"}]', encoding="utf-8")
    (tmp_path / "shifts.json").write_text('[{"id": 1}]', encoding="utf-8")
    data, skipped = load_entity_folder(tmp_path)
    assert data == {"employees": [{"name": "Alice"}], "shifts": [{"id": 1}]}
    assert skipped == []


def test_load_entity_folder_searches_recursively(tmp_path):
    nested = tmp_path / "ambunet_export"
    nested.mkdir()
    (nested / "employees.json").write_text('[{"name": "Alice"}]', encoding="utf-8")
    data, skipped = load_entity_folder(tmp_path)
    assert data == {"employees": [{"name": "Alice"}]}
    assert skipped == []


def test_load_entity_folder_skips_unparseable_files_without_aborting(tmp_path):
    (tmp_path / "employees.json").write_text('[{"name": "Alice"}]', encoding="utf-8")
    (tmp_path / "broken.json").write_text("not valid json", encoding="utf-8")
    data, skipped = load_entity_folder(tmp_path)
    assert data == {"employees": [{"name": "Alice"}]}
    assert len(skipped) == 1
    assert skipped[0][0] == "broken.json"


def test_load_entity_folder_ignores_macos_appledouble_files(tmp_path):
    # Regression test: a real zip the user tried was packaged on a Mac,
    # which mirrors every real file with a non-JSON "._filename" shadow
    # file - these aren't a parse failure, they should be silently
    # ignored rather than reported as 118 "couldn't be read" files.
    (tmp_path / "employees.json").write_text('[{"name": "Alice"}]', encoding="utf-8")
    (tmp_path / "._employees.json").write_bytes(b"\x00\x05\x16\x07not real json")
    data, skipped = load_entity_folder(tmp_path)
    assert data == {"employees": [{"name": "Alice"}]}
    assert skipped == []


def test_load_entity_zip_reads_json_members_regardless_of_nesting(tmp_path):
    zip_path = tmp_path / "export.zip"
    with zipfile.ZipFile(zip_path, "w") as zf:
        zf.writestr("ambunet_export/employees.json", '[{"name": "Alice"}]')
        zf.writestr("ambunet_export/shifts.json", '[{"id": 1}]')
    data, skipped = load_entity_zip(zip_path)
    assert data == {"employees": [{"name": "Alice"}], "shifts": [{"id": 1}]}
    assert skipped == []


def test_load_entity_zip_skips_unparseable_files_without_aborting(tmp_path):
    zip_path = tmp_path / "export.zip"
    with zipfile.ZipFile(zip_path, "w") as zf:
        zf.writestr("ambunet_export/employees.json", '[{"name": "Alice"}]')
        zf.writestr("ambunet_export/broken.json", "not valid json")
    data, skipped = load_entity_zip(zip_path)
    assert data == {"employees": [{"name": "Alice"}]}
    assert len(skipped) == 1
    assert skipped[0][0] == "broken.json"


def test_load_entity_zip_ignores_macos_appledouble_and_metadata_folder(tmp_path):
    zip_path = tmp_path / "export.zip"
    with zipfile.ZipFile(zip_path, "w") as zf:
        zf.writestr("ambunet_export/employees.json", '[{"name": "Alice"}]')
        zf.writestr("ambunet_export/._employees.json", b"\x00\x05\x16\x07not real json")
        zf.writestr("__MACOSX/ambunet_export/._employees.json", b"\x00\x05\x16\x07not real json")
    data, skipped = load_entity_zip(zip_path)
    assert data == {"employees": [{"name": "Alice"}]}
    assert skipped == []


def test_humanize_field_name_splits_camel_case():
    assert humanize_field_name("firstName") == "First Name"


def test_humanize_field_name_preserves_known_acronyms():
    assert humanize_field_name("nhsNumber") == "NHS Number"
    assert humanize_field_name("gcsEyes") == "GCS Eyes"


def test_humanize_field_name_handles_snake_case():
    assert humanize_field_name("first_name") == "First Name"


def test_sanitize_filename_replaces_invalid_windows_characters():
    assert sanitize_filename('a/b\\c:d*e?f"g<h>i|j') == "a_b_c_d_e_f_g_h_i_j"


def test_sanitize_filename_falls_back_when_empty():
    assert sanitize_filename("   ") == "record"


def test_record_label_prefers_a_number_field():
    record = {"_id": "abc", "epcrNumber": "0101", "name": "ignored"}
    assert record_label(record, 0) == "0101"


def test_record_label_handles_a_number_field_that_is_an_int():
    # Regression: cadincidents.incidentNumber is a plain int, not a
    # string - the original suffix check required isinstance(value, str)
    # and silently fell through to the _id fallback instead.
    record = {"_id": "abc", "incidentNumber": 1109262001}
    assert record_label(record, 0) == "1109262001"


def test_record_label_matches_compound_title_field():
    record = {"_id": "abc", "policyTitle": "Fire Safety Policy"}
    assert record_label(record, 0) == "Fire Safety Policy"


def test_record_label_falls_back_to_demographics_name():
    record = {"_id": "abc", "demographics": {"firstName": "Jane", "lastName": "Doe"}}
    assert record_label(record, 0) == "Jane Doe"


def test_record_label_falls_back_to_id_then_index():
    assert record_label({"_id": "abc123"}, 0) == "abc123"
    assert record_label({}, 4) == "record_5"


def test_generate_pdfs_writes_one_pdf_per_record(tmp_path):
    data = {
        "epcrs": [
            {"_id": "1", "epcrNumber": "E001", "demographics": {"firstName": "A"}},
            {"_id": "2", "epcrNumber": "E002", "demographics": {"firstName": "B"}},
        ],
        "employees": [{"_id": "3", "name": "Not a document entity"}],
    }
    counts = generate_pdfs(data, tmp_path)
    assert counts == {"epcrs": 2}
    pdf_dir = tmp_path / "pdfs" / "epcrs"
    pdfs = sorted(pdf_dir.glob("*.pdf"))
    assert [p.name for p in pdfs] == ["E001.pdf", "E002.pdf"]
    assert pdfs[0].read_bytes().startswith(b"%PDF")
    assert not (tmp_path / "pdfs" / "employees").exists()


def test_generate_pdfs_dedupes_filename_collisions(tmp_path):
    data = {
        "epcrs": [
            {"_id": "1", "title": "Duplicate"},
            {"_id": "2", "title": "Duplicate"},
        ]
    }
    generate_pdfs(data, tmp_path)
    pdf_dir = tmp_path / "pdfs" / "epcrs"
    names = sorted(p.name for p in pdf_dir.glob("*.pdf"))
    assert names == ["Duplicate.pdf", "Duplicate_2.pdf"]


def test_log_error_writes_traceback_to_appdata(tmp_path, monkeypatch):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    try:
        raise ValueError("something broke")
    except ValueError as exc:
        log_path = log_error("Processing test.zip", exc)

    assert log_path == tmp_path / "EDMSDataBridge" / "EDMSDataBridge.log"
    content = log_path.read_text(encoding="utf-8")
    assert "Processing test.zip" in content
    assert "ValueError: something broke" in content


def test_log_error_appends_across_multiple_errors(tmp_path, monkeypatch):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    for i in range(2):
        try:
            raise ValueError(f"error {i}")
        except ValueError as exc:
            log_error("ctx", exc)

    content = (tmp_path / "EDMSDataBridge" / "EDMSDataBridge.log").read_text(encoding="utf-8")
    assert "error 0" in content
    assert "error 1" in content


def test_load_version_reads_the_real_version_file():
    # In dev mode resource_path() resolves relative to the repo root,
    # where the real VERSION file lives.
    version = load_version()
    assert version != "unknown"
    assert len(version.split(".")) == 3


def test_load_version_falls_back_when_missing(monkeypatch):
    monkeypatch.setattr(
        "edms_databridge.resource_path", lambda relative_path: Path("no/such/VERSION")
    )
    assert load_version() == "unknown"


@pytest.mark.parametrize(
    "version,expected",
    [
        ("1.2.3", (1, 2, 3)),
        ("v1.2.3", (1, 2, 3)),
        ("V1.2.3", (1, 2, 3)),
        ("1.2", (1, 2, 0)),
        ("2", (2, 0, 0)),
        ("abc", (0, 0, 0)),
    ],
)
def test_parse_version(version, expected):
    assert parse_version(version) == expected


def test_parse_version_orders_correctly_for_comparison():
    assert parse_version("0.2.0") > parse_version("0.1.9")
    assert parse_version("1.0.0") > parse_version("0.99.99")
    assert parse_version("1.2.3") == parse_version("v1.2.3")


def test_get_latest_release_version_returns_tag_on_success(monkeypatch):
    class FakeResponse:
        def __enter__(self):
            return self

        def __exit__(self, *args):
            return False

        def read(self):
            return b'{"tag_name": "v0.2.0"}'

    monkeypatch.setattr(
        "edms_databridge.urllib.request.urlopen", lambda *a, **k: FakeResponse()
    )
    assert get_latest_release_version() == "v0.2.0"


def test_get_latest_release_version_returns_none_on_network_failure(monkeypatch):
    def raise_error(*args, **kwargs):
        raise TimeoutError("no network")

    monkeypatch.setattr("edms_databridge.urllib.request.urlopen", raise_error)
    assert get_latest_release_version() is None


def test_entity_display_name_uses_the_mapping():
    assert entity_display_name("employeeapplications") == "Employee Application"
    assert entity_display_name("cadincidents") == "CAD Incident"


def test_entity_display_name_falls_back_to_humanize():
    assert entity_display_name("someunmappedentity") == "Someunmappedentity"


def test_vdis_is_not_a_document_entity():
    # Regression: vdis was originally an unconfirmed guess in
    # DOCUMENT_ENTITIES, but it's actually Vehicle Daily Inspection - a
    # routine checklist, not a narrative case record - confirmed via
    # research and moved to the tabular (Excel) side.
    assert "vdis" not in DOCUMENT_ENTITIES


def test_every_document_entity_has_a_display_name():
    unmapped = DOCUMENT_ENTITIES - set(ENTITY_DISPLAY_NAMES)
    assert unmapped == set()

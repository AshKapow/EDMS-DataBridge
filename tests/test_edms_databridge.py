import json
import tkinter as tk
from pathlib import Path

import pytest

from edms_databridge import (
    load_json,
    load_logo_image,
    parse_dnd_filepaths,
    process_data,
    resource_path,
)


def test_load_json_valid(tmp_path):
    path = tmp_path / "data.json"
    path.write_text('{"a": 1}', encoding="utf-8")
    assert load_json(str(path)) == {"a": 1}


def test_load_json_handles_bom(tmp_path):
    path = tmp_path / "data.json"
    path.write_bytes('{"a": 1}'.encode("utf-8-sig"))
    assert load_json(str(path)) == {"a": 1}


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

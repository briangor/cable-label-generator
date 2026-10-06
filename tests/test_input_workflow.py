from pathlib import Path

from openpyxl import load_workbook

from label_generator import load_profile, validate_sheet_structure

ROOT = Path(__file__).resolve().parents[1]
PROFILE = ROOT / "profiles" / "example.yaml"
WORKBOOK = ROOT / "examples" / "example.xlsx"


def test_selected_example_sheet_structure_is_valid():
    profile = load_profile(PROFILE)
    workbook = load_workbook(WORKBOOK, data_only=True)
    validate_sheet_structure(workbook[profile["workbook"]["sheet"]], profile)


def test_invalid_sheet_structure_is_rejected(tmp_path):
    profile = load_profile(PROFILE)
    workbook = load_workbook(WORKBOOK)
    ws = workbook.create_sheet("Invalid")
    ws["A1"] = "not a cable section"
    path = tmp_path / "invalid.xlsx"
    workbook.save(path)

    workbook = load_workbook(path, data_only=True)
    try:
        validate_sheet_structure(workbook["Invalid"], profile)
    except ValueError as exc:
        assert "no matching cable sections" in str(exc)
    else:
        raise AssertionError("Invalid worksheet should have been rejected")

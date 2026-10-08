from pathlib import Path

from openpyxl import load_workbook

from label_generator import load_profile, validate_sheet_structure

ROOT = Path(__file__).resolve().parents[1]
PROFILE = ROOT / "profiles" / "example.yaml"
WORKBOOK = ROOT / "workbook" / "example.xlsx"


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


def test_declining_final_confirmation_does_not_generate_docx(tmp_path, monkeypatch):
    from types import SimpleNamespace

    import generate_labels

    profile = load_profile(PROFILE)
    args = SimpleNamespace(
        workbook_dir=WORKBOOK.parent,
        output_docx=tmp_path / "labels" / "labels_20261008-1200.docx",
        template=ROOT / "templates" / "template.docx",
    )

    confirmations = iter([True, True, False])
    monkeypatch.setattr(generate_labels, "_confirm", lambda prompt: next(confirmations))
    monkeypatch.setattr(generate_labels, "_choose", lambda items, title: items[0])

    generate_calls = []
    monkeypatch.setattr(
        generate_labels,
        "generate_document",
        lambda *args, **kwargs: generate_calls.append((args, kwargs)),
    )

    try:
        generate_labels._interactive(args, profile)
    except SystemExit as exc:
        assert str(exc) == "Generation cancelled."
    else:
        raise AssertionError("Interactive generation should have been cancelled")

    assert generate_calls == []
    assert not args.output_docx.exists()

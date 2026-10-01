from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT

import generate_labels
from conftest import make_profile


def make_template(path):
    document = Document()
    table = document.add_table(rows=4, cols=6)
    document.add_page_break()
    document.save(path)


def test_two_labels_per_cable_invariant():
    records = [
        (1, "peer-a", "device-a", "25GE1/0/1"),
        (2, "peer-b", "device-a", "25GE1/0/2"),
        (3, "peer-c", "device-a", "25GE1/0/3"),
    ]

    labels = generate_labels.duplicate_label_groups(records, group_size=2)

    assert len(labels) == 6
    generate_labels.validate_two_labels_per_cable(records, labels)


def test_two_labels_per_cable_rejects_incomplete_stream():
    records = [(1, "peer-a", "device-a", "25GE1/0/1")]

    try:
        generate_labels.validate_two_labels_per_cable(records, records)
    except ValueError as exc:
        assert "require 2 physical labels" in str(exc)
    else:
        raise AssertionError("Expected two-label validation to fail")


def test_full_generation_structure(tmp_path, synthetic_workbook):
    profile = make_profile()
    template = tmp_path / "template.docx"
    output = tmp_path / "labels.docx"
    make_template(template)

    # Make the synthetic template compatible with the configured geometry.
    document = Document(template)
    table = document.tables[0]
    for row in table.rows:
        for cell in row.cells:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    table.cell(0, 1).text = "template"
    document.save(template)

    # Three left + two right = five cables, therefore ten physical labels.
    # Four labels fit per table, producing three tables and two breaks.
    profile["validation"] = {
        "expected_cables": 5,
        "expected_labels": 10,
        "expected_tables": 3,
        "expected_pages": 6,
    }

    # Use a numeric ordering here because the synthetic records are not
    # intended to reproduce the Huawei reference sequence.
    profile["ordering"] = {"type": "numeric"}

    cable_count, label_count, _ = generate_labels.generate_document(
        synthetic_workbook,
        output,
        template,
        profile,
    )

    assert cable_count == 5
    assert label_count == 10

    document = Document(output)
    assert len(document.tables) == 3
    assert sum(
        1
        for paragraph in document.paragraphs
        if paragraph._p.xpath(".//w:br[@w:type='page']")
    ) == 2

    generate_labels.validate_generated_layout(
        Document(template), document, profile, label_count
    )

import generate_labels
from conftest import make_profile


def test_extracts_all_configured_sides(synthetic_workbook):
    profile = make_profile()
    workbook = generate_labels.load_workbook(synthetic_workbook, data_only=True)
    ws = workbook["Cables"]
    sections = generate_labels.extract_sections(ws, profile)

    counts = {
        side["name"]: sum(
            len(side_config["records"])
            for section in sections
            for side_config in section["sides"]
            if side_config["name"] == side["name"]
        )
        for side in profile["sides"]
    }

    records = [
        record
        for section in sections
        for side in section["sides"]
        for record in side["records"]
    ]

    assert len(sections) == 2
    assert counts == {"left": 3, "right": 2}
    assert len(records) == 5

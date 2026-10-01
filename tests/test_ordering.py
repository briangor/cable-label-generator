import generate_labels
from conftest import make_profile


def make_sections():
    return [
        {
            "sides": [
                {
                    "name": "left",
                    "records": [
                        (2, "PEER-B", "DEV-01", "25GE1/0/2"),
                        (1, "PEER-A", "DEV-01", "25GE1/0/1"),
                        (9, "PEER-X", "DEV-03", "25GE1/0/9"),
                    ],
                },
                {
                    "name": "right",
                    "records": [(1, "PEER-C", "DEV-02", "25GE1/0/1")],
                },
            ]
        }
    ]


def test_reference_records_keep_global_order():
    profile = make_profile()
    ordered, report = generate_labels.order_records_by_profile(make_sections(), profile)

    keys = [(record[2], record[0]) for record in ordered]
    assert keys == [("DEV-01", 1), ("DEV-01", 2), ("DEV-02", 1), ("DEV-03", 9)]
    assert report["retained"] == {("DEV-01", 1), ("DEV-01", 2), ("DEV-02", 1)}
    assert report["added"] == {("DEV-03", 9)}


def test_duplicate_device_port_is_rejected():
    profile = make_profile()
    sections = [{"sides": [{"records": [
        (1, "A", "DEV-01", "25GE1/0/1"),
        (1, "B", "DEV-01", "25GE1/0/1"),
    ]}]}]

    try:
        generate_labels.order_records_by_profile(sections, profile)
    except ValueError as exc:
        assert "Duplicate device/port records" in str(exc)
    else:
        raise AssertionError("Expected duplicate record validation to fail")

from pathlib import Path

import generate_labels


ROOT = Path(__file__).resolve().parents[1]
PROFILE = ROOT / "profiles" / "huawei_superapp.yaml"


def test_huawei_profile_layout_and_expected_set():
    profile = generate_labels.load_profile(PROFILE)
    layout = generate_labels.layout_config(profile)

    assert layout["labels_per_table"] == 30
    assert layout["pages_per_table"] == 2
    assert layout["label_rows"] == (0, 2, 4, 6, 8, 10)
    assert layout["top_label_starts"] == (1, 4, 7, 10, 13)
    assert layout["bottom_label_starts"] == (0, 3, 6, 9, 12)

    assert profile["validation"]["expected_cables"] == 730
    assert profile["validation"]["expected_labels"] == 1460
    assert profile["validation"]["expected_tables"] == 49
    assert profile["validation"]["expected_pages"] == 98

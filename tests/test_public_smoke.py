from pathlib import Path

import generate_labels


ROOT = Path(__file__).resolve().parents[1]


def test_public_profile_loads():
    profile = generate_labels.load_profile(ROOT / "profiles" / "example.yaml")
    layout = generate_labels.layout_config(profile)
    assert layout["labels_per_table"] == 30
    assert profile["validation"]["expected_cables"] == 6

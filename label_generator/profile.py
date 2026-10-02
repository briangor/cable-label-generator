from pathlib import Path
import yaml

def load_profile(path: Path) -> dict:
    """Load and minimally validate a YAML profile."""
    with path.open("r", encoding="utf-8") as handle:
        profile = yaml.safe_load(handle)

    if not isinstance(profile, dict):
        raise ValueError("Profile must contain a YAML mapping/object.")

    required = ("workbook", "sections", "sides", "ordering")
    missing = [key for key in required if key not in profile]

    if missing:
        raise ValueError(
            f"Profile is missing required sections: {', '.join(missing)}"
        )

    if not profile["sides"]:
        raise ValueError("Profile must define at least one cable side.")

    return profile



# Architecture

The generator separates reusable behavior from workbook-specific configuration.

- `label_generator/profile.py`: profile loading.
- `label_generator/extraction.py`: workbook section and cable extraction.
- `label_generator/ordering.py`: ordering and reference-dataset comparison.
- `label_generator/validation.py`: cable/label invariants and physical layout validation.
- `label_generator/rendering.py`: DOCX rendering and output generation.
- `generate_labels.py`: explicit and interactive CLI.

The Python modules do not contain project-specific device names, sheet names,
Excel mappings, or reference port sequences. Those belong in YAML profiles.

The DOCX template remains authoritative for physical table geometry, while the
profile defines the logical label grid.

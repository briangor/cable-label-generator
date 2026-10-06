# Cable Label Generator

A reusable, profile-driven Python tool for generating printable cable labels
from an Excel workbook and a Word template.

## Install

```bash
python -m venv .venv
pip install -r requirements.txt
```

For development:

```bash
pip install -r requirements-dev.txt
```

## Explicit mode

```bash
python generate_labels.py workbooks/workbook.xlsx labels.docx \
  --template templates/example.docx \
  --profile profiles/example.yaml
```

## Interactive mode

```bash
python generate_labels.py --profile profiles/example.yaml
```

Interactive mode lists the available XLSX files, confirms the workbook,
selects and confirms the worksheet, reviews detected device inventory and
generation counts, and asks for final confirmation before writing output.

If an output filename is not supplied, `labels_YYYYMMDD-HHMM.docx` is used.

## Architecture

See `docs/architecture.md`. Workbook-specific behavior belongs in YAML
profiles rather than in the reusable Python modules.

Private workbooks, templates, and infrastructure inventories should remain
outside a public repository.

## License

MIT.

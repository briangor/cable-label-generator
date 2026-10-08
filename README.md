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
python generate_labels.py workbook/example.xlsx labels.docx \
  --template templates/template.docx \
  --profile profiles/example.yaml
```

## Interactive mode

```bash
python generate_labels.py
```

Interactive mode uses `profiles/example.yaml` by default and looks for XLSX workbooks in `workbook/` by default. Use `--profile` when working with a different private or project-specific profile. It lists
the available files, confirms the workbook, selects and confirms the worksheet,
reviews detected device inventory and generation counts, and asks for final
confirmation before writing output.

Place private workbooks in `workbook/`. XLSX files in that directory are
gitignored by default, while the public `workbook/example.xlsx` remains tracked
as the sample workbook.

If an output filename is not supplied, the generated document is saved under
`labels/` using the `labels_YYYYMMDD-HHMM.docx` naming scheme. You can append
a custom suffix with `--suffix`, for example:

```bash
python generate_labels.py workbook/example.xlsx --suffix NBO
```

which produces `labels/labels_YYYYMMDD-HHMM_NBO.docx`. In interactive mode,
press Enter at the optional suffix prompt to keep the default name. Unsafe
filename characters are normalized to underscores. An explicitly supplied
output path is respected unchanged.

Generated DOCX files include document metadata such as title, author, subject,
keywords/tags, comments, category, and last modified by. These values can be
configured in the selected YAML profile. During generation, a terminal progress
bar reports completed tables.

## Architecture

See `docs/architecture.md`. Workbook-specific behavior belongs in YAML
profiles rather than in the reusable Python modules.

Private workbooks, templates, and infrastructure inventories should remain
outside a public repository. The `workbook/` directory is intended for local
user workbooks and is gitignored except for the public sample workbook.

## License

MIT.

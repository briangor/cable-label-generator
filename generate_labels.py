#!/usr/bin/env python3
"""Generate cable labels from an XLSX workbook and DOCX template.

The workbook-specific rules live in a YAML profile. The generator itself is
project-agnostic and only understands the profile schema.

Example:
    python generate_labels.py workbook.xlsx labels.docx \
        --template Labels.docx \
        --profile profiles/huawei_superapp.yaml

Dependencies:
    pip install openpyxl python-docx pyyaml
"""

from __future__ import annotations

import argparse
import math
import re
from copy import deepcopy
from pathlib import Path

import yaml
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string


# ---------------------------------------------------------------------------
# Generic label-template layout
# ---------------------------------------------------------------------------

TOP_LABEL_STARTS = (1, 4, 7, 10, 13)
BOTTOM_LABEL_STARTS = (0, 3, 6, 9, 12)
LABEL_ROWS = (0, 2, 4, 6, 8, 10)
LABELS_PER_TABLE = 30


# A cable record is:
# (port_number, peer_label, local_device, local_port)
CableRecord = tuple[int, str, str, str]


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


def parse_port_number(value) -> int | None:
    """Extract the numeric port suffix from values such as 25GE1/0/17."""
    if value is None:
        return None

    match = re.search(r"/(\d+)\s*$", str(value).strip())
    return int(match.group(1)) if match else None


def column_number(column: str) -> int:
    """Convert an Excel column letter to a 1-based column number."""
    try:
        return column_index_from_string(column)
    except ValueError as exc:
        raise ValueError(f"Invalid Excel column: {column!r}") from exc


def cell_value(ws, row: int, column: str):
    """Read a worksheet cell using an Excel column letter."""
    return ws.cell(row=row, column=column_number(column)).value


def find_section_starts(ws, section_config: dict) -> list[int]:
    """Find section header rows using profile-defined rules."""
    device_column = section_config["device_column"]
    header_row_offset = int(section_config.get("header_row_offset", 1))
    header_value = section_config.get("header_value")
    header_match = section_config.get("header_match")
    pattern = re.compile(header_match) if header_match else None

    starts = []

    for row in range(1, ws.max_row + 1 - header_row_offset):
        device = cell_value(ws, row, device_column)
        header = cell_value(ws, row + header_row_offset, device_column)

        if not isinstance(device, str):
            continue

        if pattern and not pattern.search(device):
            continue

        if header_value is not None and header != header_value:
            continue

        starts.append(row)

    return starts


def extract_sections(ws, profile: dict) -> list[dict]:
    """Extract all configured cable sides from all detected sections."""
    section_config = profile["sections"]
    side_configs = profile["sides"]
    starts = find_section_starts(ws, section_config)
    data_start_offset = int(section_config.get("data_start_offset", 2))

    sections = []

    for index, start in enumerate(starts):
        end = starts[index + 1] - 1 if index + 1 < len(starts) else ws.max_row

        sides = []

        for side in side_configs:
            device_column = side["device_column"]
            port_column = side["port_column"]
            label_column = side["label_column"]

            device_value = cell_value(ws, start, device_column)
            device = str(device_value).strip() if device_value is not None else ""

            records: list[CableRecord] = []

            for row in range(start + data_start_offset, end + 1):
                local_port = cell_value(ws, row, port_column)
                peer_label = cell_value(ws, row, label_column)

                if local_port in (None, "") or peer_label in (None, ""):
                    continue

                port_number = parse_port_number(local_port)
                if port_number is None:
                    continue

                records.append(
                    (
                        port_number,
                        str(peer_label).strip(),
                        device,
                        str(local_port).strip(),
                    )
                )

            if device or records:
                sides.append(
                    {
                        "name": side.get("name", "side"),
                        "device": device,
                        "records": records,
                    }
                )

        sections.append({"start": start, "sides": sides})

    return sections


def compare_reference_records(
    records: list[CableRecord],
    profile: dict,
) -> dict:
    """Compare current workbook records with the profile reference dataset."""
    ordering = profile["ordering"]

    if ordering.get("type", "numeric") != "reference":
        return {
            "reference": set(),
            "current": {(record[2], record[0]) for record in records},
            "retained": set(),
            "removed": set(),
            "added": {(record[2], record[0]) for record in records},
        }

    reference = ordering.get("reference", [])
    reference_keys = {(str(device), int(port)) for device, port in reference}
    current_keys = {(record[2], record[0]) for record in records}

    return {
        "reference": reference_keys,
        "current": current_keys,
        "retained": reference_keys & current_keys,
        "removed": reference_keys - current_keys,
        "added": current_keys - reference_keys,
    }


def order_records_by_profile(
    sections: list[dict],
    profile: dict,
) -> tuple[list[CableRecord], dict]:
    """Order records according to the profile and report dataset changes."""
    ordering = profile["ordering"]
    ordering_type = ordering.get("type", "numeric")

    records: list[CableRecord] = []
    for section in sections:
        for side in section["sides"]:
            records.extend(side["records"])

    change_report = compare_reference_records(records, profile)

    current_keys = [
        (record[2], record[0])
        for record in records
    ]

    if len(current_keys) != len(set(current_keys)):
        duplicates = sorted(
            {key for key in current_keys if current_keys.count(key) > 1},
            key=lambda key: (key[0], key[1]),
        )
        raise ValueError(
            "Duplicate device/port records found in workbook: "
            + ", ".join(f"{device}:{port}" for device, port in duplicates)
        )

    if ordering_type == "numeric":
        return sorted(records, key=lambda record: record[0]), change_report

    if ordering_type != "reference":
        raise ValueError(
            f"Unsupported ordering type: {ordering_type!r}"
        )

    reference = ordering.get("reference", [])
    reference_key_type = ordering.get("reference_key", "device_port")

    if reference_key_type != "device_port":
        raise ValueError(
            f"Unsupported reference_key: {reference_key_type!r}"
        )

    def record_key(record: CableRecord) -> tuple[str, int]:
        return record[2], record[0]

    # Preserve all records from the workbook. The reference list determines
    # the ordering of records that are present in both datasets.
    by_key: dict[tuple[str, int], list[CableRecord]] = {}
    discovered_order: list[tuple[str, int]] = []

    for record in records:
        key = record_key(record)
        by_key.setdefault(key, []).append(record)
        if key not in discovered_order:
            discovered_order.append(key)

    result: list[CableRecord] = []
    consumed: set[tuple[str, int]] = set()

    # First emit records in the exact global order captured from the
    # reference document. This is intentionally global rather than grouped
    # by device because the physical reference layout can cross device
    # boundaries within a five-label group.
    for entry in reference:
        if not isinstance(entry, (list, tuple)) or len(entry) != 2:
            raise ValueError(
                "Each reference ordering entry must be [device, port]."
            )

        device, port = entry
        key = (str(device), int(port))

        matching = by_key.get(key)
        if matching:
            result.extend(matching)
            consumed.add(key)

    # Records absent from the reference are retained rather than discarded.
    # By default they are appended in workbook discovery order. This gives
    # additions deterministic behavior without inventing a project-specific
    # ordering rule.
    unknown_mode = ordering.get("unknown_records", "after_reference")

    unknown_keys = [
        key for key in discovered_order if key not in consumed
    ]

    if unknown_mode == "after_reference":
        unknown_sort = ordering.get("unknown_sort", "discovered")

        if unknown_sort == "numeric":
            unknown_keys.sort(key=lambda key: (key[0], key[1]))
        elif unknown_sort != "discovered":
            raise ValueError(
                f"Unsupported unknown_sort: {unknown_sort!r}"
            )

        for key in unknown_keys:
            result.extend(by_key[key])
    elif unknown_keys:
        raise ValueError(
            "Workbook contains records not present in the ordering profile."
        )

    return result, change_report


def duplicate_label_groups(
    records: list[CableRecord],
    group_size: int = 5,
) -> list[CableRecord]:
    """Duplicate each physical label group for the two cable ends."""
    stream: list[CableRecord] = []

    for start in range(0, len(records), group_size):
        chunk = records[start:start + group_size]
        stream.extend(chunk)
        stream.extend(chunk)

    return stream


# ---------------------------------------------------------------------------
# DOCX rendering
# ---------------------------------------------------------------------------


def clear_cell(cell) -> None:
    """Clear cell text while retaining its table/cell formatting."""
    cell.text = ""
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    if cell.paragraphs:
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT


def set_label_cell(cell, text: str, *, bold: bool) -> None:
    """Write a label into a preformatted template cell."""
    clear_cell(cell)
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    run = paragraph.add_run(text)
    run.font.size = Pt(9)
    run.bold = bold


def fill_table(table, labels: list[CableRecord]) -> None:
    """Fill one 30-label template table."""
    for row in LABEL_ROWS:
        starts = (
            TOP_LABEL_STARTS
            if row in (0, 4, 8)
            else BOTTOM_LABEL_STARTS
        )

        for start in starts:
            clear_cell(table.cell(row, start))
            clear_cell(table.cell(row, start + 1))

    for index, item in enumerate(labels[:LABELS_PER_TABLE]):
        row = LABEL_ROWS[index // 5]
        starts = (
            TOP_LABEL_STARTS
            if row in (0, 4, 8)
            else BOTTOM_LABEL_STARTS
        )
        start = starts[index % 5]

        _, peer_label, local_device, local_port = item

        set_label_cell(table.cell(row, start), peer_label, bold=True)
        set_label_cell(
            table.cell(row, start + 1),
            f"{local_device}\n{local_port}",
            bold=False,
        )


def generate_document(
    input_xlsx: Path,
    output_docx: Path,
    template_docx: Path,
    profile: dict,
) -> tuple[int, int]:
    """Generate the output DOCX and return (cable_count, label_count)."""
    workbook = load_workbook(
        input_xlsx,
        data_only=True,
        read_only=False,
    )

    sheet_name = profile["workbook"]["sheet"]

    if sheet_name not in workbook.sheetnames:
        raise ValueError(
            f'Sheet "{sheet_name}" was not found. '
            f"Available sheets: {', '.join(workbook.sheetnames)}"
        )

    ws = workbook[sheet_name]
    sections = extract_sections(ws, profile)

    if not sections:
        raise ValueError(
            f'No cable sections were found in sheet "{sheet_name}".'
        )

    ordered_records, change_report = order_records_by_profile(
        sections,
        profile,
    )
    label_stream = duplicate_label_groups(ordered_records)

    cable_count = len(ordered_records)
    label_count = len(label_stream)

    template_doc = Document(template_docx)

    if not template_doc.tables:
        raise ValueError(
            "The template DOCX does not contain a table."
        )

    template_table_xml = deepcopy(template_doc.tables[0]._tbl)

    page_break_xml = None
    for paragraph in template_doc.paragraphs:
        if paragraph._p.xpath(
            ".//w:br[@w:type='page']"
        ):
            page_break_xml = deepcopy(paragraph._p)
            break

    if page_break_xml is None:
        raise ValueError(
            "The template DOCX does not contain the expected "
            "page-break paragraph."
        )

    body = template_doc._body._element
    sect_pr = body.sectPr

    for child in list(body):
        if child is sect_pr:
            continue
        body.remove(child)

    for table_index in range(
        math.ceil(len(label_stream) / LABELS_PER_TABLE)
    ):
        table_xml = deepcopy(template_table_xml)
        body.insert(len(body) - 1, table_xml)

        table = template_doc.tables[-1]
        start = table_index * LABELS_PER_TABLE
        end = start + LABELS_PER_TABLE

        fill_table(table, label_stream[start:end])

        if end < len(label_stream):
            body.insert(
                len(body) - 1,
                deepcopy(page_break_xml),
            )

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    template_doc.save(output_docx)

    return cable_count, label_count, change_report


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate cable labels from an XLSX workbook."
    )
    parser.add_argument("input_xlsx", type=Path)
    parser.add_argument("output_docx", type=Path)
    parser.add_argument(
        "--template",
        type=Path,
        default=Path("Labels.docx"),
        help="DOCX used as the formatting/layout template.",
    )
    parser.add_argument(
        "--profile",
        type=Path,
        required=True,
        help="YAML profile describing the workbook structure and ordering.",
    )

    args = parser.parse_args()
    profile = load_profile(args.profile)

    cable_count, label_count, change_report = generate_document(
        args.input_xlsx,
        args.output_docx,
        args.template,
        profile,
    )

    tables = math.ceil(label_count / LABELS_PER_TABLE)

    print(f"Profile:      {profile.get('name', args.profile.stem)}")
    print(f"Source sheet: {profile['workbook']['sheet']}")
    print(f"Cables used:  {cable_count}")
    print(f"Labels:       {label_count}")
    print(f"Tables:       {tables}")

    if change_report["reference"]:
        print(f"Reference records retained: {len(change_report['retained'])}")
        print(f"Reference records removed:  {len(change_report['removed'])}")
        print(f"New records:                {len(change_report['added'])}")

        if change_report["removed"]:
            print("Removed reference records:")
            for device, port in sorted(change_report["removed"]):
                print(f"  - {device}: {port}")

        if change_report["added"]:
            print("New workbook records:")
            for device, port in sorted(change_report["added"]):
                print(f"  + {device}: {port}")

    print(f"Output:       {args.output_docx}")


if __name__ == "__main__":
    main()

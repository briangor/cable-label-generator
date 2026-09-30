#!/usr/bin/env python3
"""
Generate a server-rack cable labels DOCX from the selected sheet.

Usage:
    python generate_labels.py input.xlsx output.docx --template Labels.docx

Dependencies:
    pip install openpyxl python-docx

The script uses the supplied Labels.docx as the formatting/template master.

Cable data is extracted from both LMU sides:
    A = Local Port
    B = Peer Port / label name

    I = Local Port
    J = Peer Port / label name
"""

from __future__ import annotations

import argparse
import math
import re
from copy import deepcopy
from pathlib import Path

from openpyxl import load_workbook
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.shared import Pt


SHEET_NAME = "3.4.1 LMU Cable"
# TODO: allow select sheet from list of sheets 

# The example document contains 15 physical columns.
# Each label occupies two adjacent cells.
TOP_LABEL_STARTS = (1, 4, 7, 10, 13)
BOTTOM_LABEL_STARTS = (0, 3, 6, 9, 12)
LABEL_ROWS = (0, 2, 4, 6, 8, 10)

# A table contains 6 label rows x 5 labels = 30 labels.
LABELS_PER_TABLE = 30


def parse_port_number(value) -> int | None:
    """Extract the numeric port suffix from values such as 25GE1/0/17."""
    if value is None:
        return None

    match = re.search(r"/(\d+)\s*$", str(value).strip())
    return int(match.group(1)) if match else None


def find_section_starts(ws) -> list[int]:
    """
    Find section header rows.

    In the source sheet, an LMU cable section starts where:
      A = local device name
      I = paired device name
    """
    starts = []

    for row in range(1, ws.max_row + 1):
        local_device = ws.cell(row, 1).value
        paired_device = ws.cell(row, 9).value

        if (
            isinstance(local_device, str)
            and isinstance(paired_device, str)
            and local_device.startswith("LMU-")
            and paired_device.startswith("LMU-")
        ):
            starts.append(row)

    return starts


def extract_label_queues(ws) -> dict[int, list[tuple]]:
    """
    Extract cable labels from both LMU sides and arrange them into
    three queues based on the local-port number modulo 3.

    The source sheet contains two cable sides:

        A/B:
            A = Local Port
            B = Peer Port / label name

        I/J:
            I = Local Port
            J = Peer Port / label name

    Each populated A/B or I/J row represents one cable record.

    Each item is:
        (port_number, peer_label, local_device, local_port)
    """
    section_starts = find_section_starts(ws)

    queues = {1: [], 2: [], 0: []}

    for index, start in enumerate(section_starts):
        end = (
            section_starts[index + 1] - 1
            if index + 1 < len(section_starts)
            else ws.max_row
        )

        # The two device names are the section headers for the
        # left and right LMU cable sides.
        local_device = str(ws.cell(start, 1).value).strip()
        paired_device = str(ws.cell(start, 9).value).strip()

        # Row start + 2 skips the section title and column headings.
        for row in range(start + 2, end + 1):
            # ---------------------------------------------------------
            # Left side: columns A/B
            # ---------------------------------------------------------
            local_port = ws.cell(row, 1).value
            peer_label = ws.cell(row, 2).value

            if local_port not in (None, "") and peer_label not in (None, ""):
                port_number = parse_port_number(local_port)

                if port_number is not None:
                    item = (
                        port_number,
                        str(peer_label).strip(),
                        local_device,
                        str(local_port).strip(),
                    )
                    queues[port_number % 3].append(item)

            # ---------------------------------------------------------
            # Right side: columns I/J
            # ---------------------------------------------------------
            local_port = ws.cell(row, 9).value
            peer_label = ws.cell(row, 10).value

            if local_port not in (None, "") and peer_label not in (None, ""):
                port_number = parse_port_number(local_port)

                if port_number is not None:
                    item = (
                        port_number,
                        str(peer_label).strip(),
                        paired_device,
                        str(local_port).strip(),
                    )
                    queues[port_number % 3].append(item)

    # Each section was already sorted by physical port before being
    # appended to its queue. Do not sort the combined queues globally:
    # section order is significant to the label layout.
    return queues


def build_label_stream(queues: dict[int, list[tuple]]) -> list[tuple]:
    """
    Build the label ordering used by the current generator.

    For each output group:
      - take up to 5 from port sequence 1,4,7,...
      - take up to 5 from port sequence 2,5,8,...
      - take up to 5 from port sequence 3,6,9,...
      - repeat each five-label group once

    Repeating the group produces two copies of every physical label.
    """
    stream = []

    while any(queues.values()):
        for remainder in (1, 2, 0):
            chunk = queues[remainder][:5]
            del queues[remainder][:5]

            if not chunk:
                continue

            # The example prints the same five labels twice.
            stream.extend(chunk)
            stream.extend(chunk)

    return stream


def clear_cell(cell) -> None:
    """Clear cell text while retaining its table/cell formatting."""
    cell.text = ""
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    if cell.paragraphs:
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT


def set_label_cell(cell, text: str, *, bold: bool) -> None:
    """Write a label into a preformatted template cell."""
    clear_cell(cell)

    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    run = paragraph.add_run(text)
    run.font.size = Pt(9)
    run.bold = bold


def fill_table(table, labels: list[tuple]) -> None:
    """
    Fill one 30-label template table.

    labels contains:
        (port_number, peer_label, local_device, local_port)
    """
    # Clear all label positions first so a partially filled final table
    # cannot retain sample text from the template.
    for row in LABEL_ROWS:
        starts = TOP_LABEL_STARTS if row in (0, 4, 8) else BOTTOM_LABEL_STARTS

        for start in starts:
            clear_cell(table.cell(row, start))
            clear_cell(table.cell(row, start + 1))

    # Fill available labels.
    for index, item in enumerate(labels[:LABELS_PER_TABLE]):
        row = LABEL_ROWS[index // 5]
        starts = TOP_LABEL_STARTS if row in (0, 4, 8) else BOTTOM_LABEL_STARTS
        start = starts[index % 5]

        _, peer_label, local_device, local_port = item

        set_label_cell(
            table.cell(row, start),
            peer_label,
            bold=True,
        )
        set_label_cell(
            table.cell(row, start + 1),
            f"{local_device}\n{local_port}",
            bold=False,
        )


def generate_document(
    input_xlsx: Path,
    output_docx: Path,
    template_docx: Path,
) -> tuple[int, int]:
    """Generate the output DOCX and return (cable_count, label_count)."""
    workbook = load_workbook(input_xlsx, data_only=True, read_only=False)

    if SHEET_NAME not in workbook.sheetnames:
        raise ValueError(
            f'Sheet "{SHEET_NAME}" was not found. '
            f"Available sheets: {', '.join(workbook.sheetnames)}"
        )

    ws = workbook[SHEET_NAME]

    queues = extract_label_queues(ws)
    cable_count = sum(len(queue) for queue in queues.values())
    label_stream = build_label_stream(queues)

    template_doc = Document(template_docx)

    if not template_doc.tables:
        raise ValueError("The template DOCX does not contain a table.")

    # The first table is the formatting master.
    template_table_xml = deepcopy(template_doc.tables[0]._tbl)

    # Locate an existing page-break paragraph from the example.
    page_break_xml = None
    for paragraph in template_doc.paragraphs:
        if paragraph._p.xpath(".//w:br[@w:type='page']"):
            page_break_xml = deepcopy(paragraph._p)
            break

    if page_break_xml is None:
        raise ValueError(
            "The template DOCX does not contain the expected page-break paragraph."
        )

    body = template_doc._body._element
    sect_pr = body.sectPr

    # Remove all existing tables and page-break paragraphs.
    # Keep sectPr as the final body element.
    for child in list(body):
        if child is sect_pr:
            continue
        body.remove(child)

    # Build fresh tables from the template master.
    for table_index in range(math.ceil(len(label_stream) / LABELS_PER_TABLE)):
        table_xml = deepcopy(template_table_xml)
        body.insert(len(body) - 1, table_xml)

        # The table is now the last table in the document.
        table = template_doc.tables[-1]

        start = table_index * LABELS_PER_TABLE
        end = start + LABELS_PER_TABLE
        fill_table(table, label_stream[start:end])

        if end < len(label_stream):
            body.insert(len(body) - 1, deepcopy(page_break_xml))

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    template_doc.save(output_docx)

    return cable_count, len(label_stream)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate server-rack cable labels from an Excel sheet."
    )
    parser.add_argument("input_xlsx", type=Path)
    parser.add_argument("output_docx", type=Path)
    parser.add_argument(
        "--template",
        type=Path,
        default=Path("Labels.docx"),
        help="Example DOCX used as the formatting template.",
    )

    args = parser.parse_args()

    cable_count, label_count = generate_document(
        args.input_xlsx,
        args.output_docx,
        args.template,
    )

    print(f"Source sheet: {SHEET_NAME}")
    print(f"Cables used:  {cable_count}")
    print(f"Labels:       {label_count}")
    print(f"Output:       {args.output_docx}")


if __name__ == "__main__":
    main()
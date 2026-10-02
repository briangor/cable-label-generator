#!/usr/bin/env python3
"""Generate printable cable labels from an XLSX workbook and DOCX template."""

from __future__ import annotations

import argparse
import math
from copy import deepcopy
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

from label_generator import extract_sections, generate_document, layout_config, load_profile


def _confirm(prompt: str) -> bool:
    return input(f"{prompt} [Y/n]: ").strip().lower() in ("", "y", "yes")


def _choose(items: list[str], title: str) -> str:
    if not items:
        raise ValueError(f"No {title.lower()} found.")
    print(f"\n{title}")
    for index, item in enumerate(items, 1):
        print(f"  {index}. {item}")
    while True:
        try:
            index = int(input("Select number: ").strip())
        except ValueError:
            print("Please enter a number.")
            continue
        if 1 <= index <= len(items):
            return items[index - 1]
        print("Please select one of the listed numbers.")


def _discover_workbooks(directory: Path) -> list[Path]:
    return sorted(
        p for p in directory.glob("*.xlsx") if not p.name.startswith("~$")
    )


def _select_sheet(workbook: Path, configured_sheet: str | None) -> str:
    book = load_workbook(workbook, read_only=True, data_only=True)
    sheets = book.sheetnames
    selected = sheets[0] if len(sheets) == 1 else _choose(sheets, "Available worksheets")
    if configured_sheet and configured_sheet != selected:
        print(f"Profile default sheet: {configured_sheet}")
    print(f"Selected worksheet: {selected}")
    if not _confirm("Use this worksheet?"):
        raise SystemExit("Generation cancelled.")
    return selected


def _review_inventory(workbook: Path, profile: dict) -> None:
    book = load_workbook(workbook, data_only=True, read_only=False)
    sheet = profile["workbook"]["sheet"]
    sections = extract_sections(book[sheet], profile)
    if not sections:
        raise ValueError(f'No cable sections were found in sheet "{sheet}".')

    devices: dict[str, int] = {}
    cables = 0
    for section in sections:
        for side in section["sides"]:
            device = side["device"]
            count = len(side["records"])
            if device:
                devices[device] = devices.get(device, 0) + count
            cables += count

    labels = cables * 2
    tables = math.ceil(labels / layout_config(profile)["labels_per_table"])
    pages = tables * layout_config(profile)["pages_per_table"]

    print("\nSource inventory")
    for device, count in devices.items():
        print(f"  {device}: {count} cable records")
    print("\nGeneration summary")
    print(f"  Cables: {cables}")
    print(f"  Physical labels: {labels}")
    print(f"  Tables: {tables}")
    print(f"  Pages: {pages}")


def _interactive(args, profile: dict):
    workbooks = _discover_workbooks(args.workbook_dir)
    workbook = args.workbook_dir / _choose(
        [p.name for p in workbooks],
        f"XLSX workbooks in {args.workbook_dir}",
    )
    if not _confirm(f"Use workbook {workbook.name}?"):
        raise SystemExit("Generation cancelled.")

    selected_sheet = _select_sheet(workbook, profile["workbook"].get("sheet"))
    profile = deepcopy(profile)
    profile["workbook"] = deepcopy(profile["workbook"])
    profile["workbook"]["sheet"] = selected_sheet

    _review_inventory(workbook, profile)

    if not _confirm("Proceed with generation?"):
        raise SystemExit("Generation cancelled.")

    output = args.output_docx or Path(
        f"labels_{datetime.now().strftime('%Y%m%d-%H%M')}.docx"
    )
    return workbook, output, args.template, profile


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate cable labels from an XLSX workbook."
    )
    parser.add_argument("input_xlsx", nargs="?", type=Path)
    parser.add_argument("output_docx", nargs="?", type=Path)
    parser.add_argument(
        "--template",
        type=Path,
        default=Path("templates/example.docx"),
        help="DOCX used as the formatting/layout template.",
    )
    parser.add_argument(
        "--profile",
        type=Path,
        required=True,
        help="YAML profile describing the workbook structure and ordering.",
    )
    parser.add_argument(
        "--workbook-dir",
        type=Path,
        default=Path("."),
        help="Directory searched for XLSX files in interactive mode.",
    )

    args = parser.parse_args()
    profile = load_profile(args.profile)

    if args.input_xlsx is None:
        input_xlsx, output_docx, template, profile = _interactive(args, profile)
    else:
        input_xlsx = args.input_xlsx
        output_docx = args.output_docx or Path(
            f"labels_{datetime.now().strftime('%Y%m%d-%H%M')}.docx"
        )
        template = args.template

    cable_count, label_count, change_report = generate_document(
        input_xlsx, output_docx, template, profile
    )

    tables = math.ceil(label_count / layout_config(profile)["labels_per_table"])
    print(f"\nProfile:      {profile.get('name', args.profile.stem)}")
    print(f"Source sheet: {profile['workbook']['sheet']}")
    print(f"Cables used:  {cable_count}")
    print(f"Labels:       {label_count}")
    print(f"Tables:       {tables}")
    print(f"Pages:        {tables * layout_config(profile)['pages_per_table']}")

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

    print(f"Output:       {output_docx}")


if __name__ == "__main__":
    main()

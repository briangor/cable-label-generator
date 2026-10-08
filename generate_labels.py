#!/usr/bin/env python3
"""Generate printable cable labels from an XLSX workbook and DOCX template."""

from __future__ import annotations

__author__ = "Brian Gor"

import argparse
import math
from copy import deepcopy
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

from label_generator import (
    duplicate_label_groups,
    extract_sections,
    generate_document,
    layout_config,
    load_profile,
    order_records_by_profile,
    validate_generated_layout,
    validate_sheet_structure,
    validate_two_labels_per_cable,
)


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


def _default_output_path() -> Path:
    """Return the standard generated-label output path."""
    timestamp = datetime.now().strftime("%Y%m%d-%H%M")
    return Path("labels") / f"labels_{timestamp}.docx"


def _resolve_output_path(output: Path | None) -> Path:
    """Use an explicit output path or the standard labels/ timestamp path."""
    return output if output is not None else _default_output_path()


def _discover_workbooks(directory: Path) -> list[Path]:
    return sorted(
        p for p in directory.glob("*.xlsx") if not p.name.startswith("~$")
    )


def _select_sheet(workbook: Path, configured_sheet: str | None, profile: dict) -> str:
    book = load_workbook(workbook, read_only=False, data_only=True)
    sheets = book.sheetnames
    selected = sheets[0] if len(sheets) == 1 else _choose(sheets, "Available worksheets")

    profile = deepcopy(profile)
    profile["workbook"] = deepcopy(profile["workbook"])
    profile["workbook"]["sheet"] = selected

    try:
        validate_sheet_structure(book[selected], profile)
    except ValueError as exc:
        print(f"\nSelected sheet failed validation: {exc}")
        raise SystemExit("Generation cancelled.") from exc

    if configured_sheet and configured_sheet != selected:
        print(f"Profile default sheet: {configured_sheet}")
    print(f"Selected worksheet: {selected}")
    print("Sheet structure: valid")
    if not _confirm("Use this worksheet?"):
        raise SystemExit("Generation cancelled.")
    return selected


def _review_inventory(workbook: Path, profile: dict) -> None:
    book = load_workbook(workbook, data_only=True, read_only=False)
    sheet = profile["workbook"]["sheet"]
    validate_sheet_structure(book[sheet], profile)
    sections = extract_sections(book[sheet], profile)
    if not sections:
        raise ValueError(f'No cable sections were found in sheet "{sheet}".')

    inventory = {}
    cables = 0
    for section in sections:
        for side in section["sides"]:
            device = side["device"]
            count = len(side["records"])
            if not device:
                continue
            entry = inventory.setdefault(
                device,
                {"rack": "-", "switch_type": "-", "switch_name": device, "cables": 0},
            )
            entry["cables"] += count
            cables += count

    inventory_config = profile.get("inventory", {})
    device_pattern = inventory_config.get("device_pattern")
    if device_pattern:
        import re
        pattern = re.compile(device_pattern)
        for device, entry in inventory.items():
            match = pattern.search(device)
            if match:
                groups = match.groupdict()
                entry["rack"] = groups.get("rack", entry["rack"])
                entry["switch_type"] = groups.get("switch_type", entry["switch_type"])
                entry["switch_name"] = groups.get("switch_name", entry["switch_name"])

    labels = cables * 2
    layout = layout_config(profile)
    tables = math.ceil(labels / layout["labels_per_table"])
    pages = tables * layout["pages_per_table"]

    print("\nSwitch inventory")
    print("  Rack     Type   Switch      Cables")
    print("  -------  -----  ----------  ------")
    for entry in inventory.values():
        print(
            f"  {entry['rack']:<7}  {entry['switch_type']:<5}  "
            f"{entry['switch_name']:<10}  {entry['cables']:>6}"
        )

    print("\nGeneration summary")
    print(f"  Total cables:          {cables}")
    print(f"  Physical labels:       {labels}")
    print(f"  Expected tables:       {tables}")
    print(f"  Expected pages:        {pages}")


def _interactive(args, profile: dict):
    workbooks = _discover_workbooks(args.workbook_dir)
    workbook = args.workbook_dir / _choose(
        [p.name for p in workbooks],
        f"XLSX workbooks in {args.workbook_dir}",
    )
    if not _confirm(f"Use workbook {workbook.name}?"):
        raise SystemExit("Generation cancelled.")

    selected_sheet = _select_sheet(workbook, profile["workbook"].get("sheet"), profile)
    profile = deepcopy(profile)
    profile["workbook"] = deepcopy(profile["workbook"])
    profile["workbook"]["sheet"] = selected_sheet

    _review_inventory(workbook, profile)

    if not _confirm("Proceed with generation?"):
        raise SystemExit("Generation cancelled.")

    output = _resolve_output_path(args.output_docx)
    return workbook, output, args.template, profile


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Generate cable labels from an XLSX workbook."
    )
    parser.add_argument("input_xlsx", nargs="?", type=Path)
    parser.add_argument("output_docx", nargs="?", type=Path)
    parser.add_argument(
        "--template",
        type=Path,
        default=Path("templates/template.docx"),
        help="DOCX used as the formatting/layout template.",
    )
    parser.add_argument(
        "--profile",
        type=Path,
        default=Path("profiles/example.yaml"),
        help=(
            "YAML profile describing the workbook structure and ordering "
            "(default: profiles/example.yaml)."
        ),
    )
    parser.add_argument(
        "--workbook-dir",
        type=Path,
        default=Path("workbook"),
        help="Directory searched for XLSX files in interactive mode (default: workbook/).",
    )
    return parser


def main() -> None:
    parser = _build_parser()
    args = parser.parse_args()
    profile = load_profile(args.profile)

    if args.input_xlsx is None:
        input_xlsx, output_docx, template, profile = _interactive(args, profile)
    else:
        input_xlsx = args.input_xlsx
        output_docx = _resolve_output_path(args.output_docx)
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

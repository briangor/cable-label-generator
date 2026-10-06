"""Reusable cable-label generation library."""

__author__ = "Brian Gor"

from .extraction import (
    CableRecord,
    cell_value,
    column_number,
    extract_sections,
    find_section_starts,
    parse_port_number,
    validate_sheet_structure,

)
from .ordering import compare_reference_records, order_records_by_profile
from .profile import load_profile
from .rendering import clear_cell, fill_table, generate_document, set_label_cell
from .validation import (
    duplicate_label_groups,
    layout_config,
    validate_generated_layout,
    validate_template_layout,
    validate_two_labels_per_cable,
)

__all__ = [
    "CableRecord",
    "cell_value",
    "column_number",
    "extract_sections",
    "find_section_starts",
    "parse_port_number",
    "validate_sheet_structure",
    "compare_reference_records",
    "order_records_by_profile",
    "load_profile",
    "clear_cell",
    "fill_table",
    "generate_document",
    "set_label_cell",
    "duplicate_label_groups",
    "layout_config",
    "validate_generated_layout",
    "validate_template_layout",
    "validate_two_labels_per_cable",
]

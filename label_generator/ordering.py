from .extraction import CableRecord

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


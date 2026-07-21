"""CSV audit and diagnostic report writers."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, List

from .config import BuilderConfig, SLOT_TO_ROLL_INDEX
from .manifest import ManifestIndex
from .utils import csv_write, norm_name
from .wishlist import select_weapon_candidates


UNRESOLVED_FIELDS = [
    "source_row", "type", "name", "reason", "original_reason", "weapon", "weapon_hash",
    "slot", "socket_index", "mapping_method", "mapping_name_hits", "generated",
    "version_status", "requested_option_count", "supported_option_count", "perks",
    "same_name_candidate_hashes", "same_name_candidate_types", "socket_candidate_hashes",
    "socket_candidate_names",
]

AUDIT_FIELDS = [
    "source_row", "weapon", "weapon_hash", "weapon_sqlite_id", "manifest_name",
    "version_status", "included_slots", "dropped_perks", "requested_option_count",
    "supported_option_count",
]
for _slot in SLOT_TO_ROLL_INDEX:
    AUDIT_FIELDS.extend([
        f"{_slot}_requested_options", f"{_slot}_supported_options", _slot,
        f"{_slot}_hash", f"{_slot}_sqlite_id", f"{_slot}_type", f"{_slot}_socket_index",
    ])
AUDIT_FIELDS.append("dim_line")


def write_extracted_report(path: Path, expanded: List[Dict[str, Any]]) -> None:
    rows = []
    for record in expanded:
        slots = record.get("perk_slots", {}) or {}
        rows.append({
            "source_row": record.get("source_row", ""),
            "weapon": record.get("weapon", ""),
            "slot_1_barrel": slots.get("slot_1_barrel", ""),
            "slot_2_magazine": slots.get("slot_2_magazine", ""),
            "slot_3_trait": slots.get("slot_3_trait", ""),
            "slot_4_trait": slots.get("slot_4_trait", ""),
            "expanded_perks_in_dim_order": " / ".join(record.get("perks", [])),
            "notes": record.get("notes", ""),
            "raw_perk_columns_json": json.dumps(
                record.get("raw_perk_columns", {}), ensure_ascii=False
            ),
            "parsed_perk_columns_json": json.dumps(
                record.get("parsed_perk_columns", {}), ensure_ascii=False
            ),
        })
    csv_write(path, rows, [
        "source_row", "weapon", "slot_1_barrel", "slot_2_magazine", "slot_3_trait",
        "slot_4_trait", "expanded_perks_in_dim_order", "notes", "raw_perk_columns_json",
        "parsed_perk_columns_json",
    ])


def write_weapon_candidate_report(
    path: Path,
    config: BuilderConfig,
    index: ManifestIndex,
    expanded: List[Dict[str, Any]],
) -> None:
    names = []
    seen = set()
    for record in expanded:
        weapon = record.get("weapon", "")
        if weapon and norm_name(weapon) not in seen:
            names.append(weapon)
            seen.add(norm_name(weapon))

    rows = []
    for weapon in names:
        candidates = index.find_weapons(weapon)
        selected = select_weapon_candidates(config, weapon, candidates)
        candidate_hashes = [candidate.hash for candidate in candidates]
        selected_hashes = [candidate.hash for candidate in selected]
        for candidate in candidates:
            rows.append({
                "weapon": weapon,
                "candidate_count": len(candidates),
                "selected_count": len(selected),
                "selected": "yes" if candidate.hash in selected_hashes else "no",
                "hash": candidate.hash,
                "sqlite_id": candidate.sql_id,
                "manifest_name": candidate.name,
                "item_type_display": candidate.item_type_display,
                "item_type_and_tier_display": candidate.item_type_and_tier_display,
                "tier_type_name": candidate.tier_type_name,
                "item_type": candidate.item_type,
                "has_plug": candidate.has_plug,
                "all_candidate_hashes": ";".join(map(str, candidate_hashes)),
                "selected_hashes": ";".join(map(str, selected_hashes)),
            })
        if not candidates:
            rows.append({"weapon": weapon, "candidate_count": 0, "selected_count": 0, "selected": "no"})
    csv_write(path, rows, [
        "weapon", "candidate_count", "selected_count", "selected", "hash", "sqlite_id",
        "manifest_name", "item_type_display", "item_type_and_tier_display", "tier_type_name",
        "item_type", "has_plug", "all_candidate_hashes", "selected_hashes",
    ])


def write_perk_candidate_report(
    path: Path,
    config: BuilderConfig,
    index: ManifestIndex,
    expanded: List[Dict[str, Any]],
) -> None:
    rows = []
    seen = set()
    for record in expanded:
        weapon = record.get("weapon", "")
        candidates = select_weapon_candidates(config, weapon, index.find_weapons(weapon))
        for weapon_candidate in candidates:
            for slot, perk in (record.get("perk_slots", {}) or {}).items():
                key = weapon_candidate.hash, slot, norm_name(perk)
                if not perk or key in seen:
                    continue
                seen.add(key)
                selected, detail = index.resolve_perk_for_weapon_slot(
                    weapon_candidate.hash,
                    slot,
                    perk,
                    record.get("parsed_perk_columns", {}),
                )
                rows.append({
                    "source_row": record.get("source_row", ""),
                    "weapon": weapon,
                    "weapon_hash": weapon_candidate.hash,
                    "weapon_sqlite_id": weapon_candidate.sql_id,
                    "slot": slot,
                    "roll_index": detail.get("roll_index", ""),
                    "socket_index": detail.get("socket_index", ""),
                    "requested_perk": perk,
                    "selected_hash": selected.hash if selected else "",
                    "selected_sqlite_id": selected.sql_id if selected else "",
                    "selected_name": selected.name if selected else "",
                    "selected_type": selected.item_type_display if selected else "",
                    "reason": detail.get("reason", ""),
                    "mapping_method": detail.get("mapping_method", ""),
                    "mapping_name_hits": detail.get("mapping_name_hits", 0),
                    "socket_candidate_count": len(detail.get("candidate_hashes", [])),
                    "socket_candidate_hashes": ";".join(map(str, detail.get("candidate_hashes", []))),
                    "socket_candidate_names": ";".join(detail.get("candidate_names", [])),
                    "same_name_hashes_in_socket": ";".join(map(str, detail.get("same_name_hashes", []))),
                    "plug_categories": ";".join(detail.get("plug_categories", [])),
                })
    csv_write(path, rows, [
        "source_row", "weapon", "weapon_hash", "weapon_sqlite_id", "slot", "roll_index",
        "socket_index", "requested_perk", "selected_hash", "selected_sqlite_id", "selected_name",
        "selected_type", "reason", "mapping_method", "mapping_name_hits", "socket_candidate_count",
        "socket_candidate_hashes", "socket_candidate_names", "same_name_hashes_in_socket",
        "plug_categories",
    ])


def write_generation_outputs(
    config: BuilderConfig,
    lines: List[str],
    unresolved: List[Dict[str, Any]],
    audit: List[Dict[str, Any]],
) -> None:
    output_dir = config.output_dir.expanduser().resolve()
    (output_dir / config.wishlist_filename).write_text(
        "\n".join(lines) + "\n", encoding="utf-8"
    )
    if config.write_diagnostics:
        csv_write(output_dir / config.unresolved_filename, unresolved, UNRESOLVED_FIELDS)
        csv_write(output_dir / config.resolved_audit_filename, audit, AUDIT_FIELDS)

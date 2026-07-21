"""Weapon-version selection and DIM wishlist generation."""

from __future__ import annotations

import itertools
from typing import Any, Dict, List, Tuple

from .config import BuilderConfig, SLOT_TO_ROLL_INDEX
from .manifest import ManifestIndex
from .models import InventoryItem
from .table import collapse_expanded_recommendations
from .utils import norm_name, sanitize_comment


def select_weapon_candidates(
    config: BuilderConfig,
    weapon: str,
    candidates: List[InventoryItem],
) -> List[InventoryItem]:
    if not candidates:
        return []
    weapon_key = norm_name(weapon)
    for override_name, hashes in config.weapon_hash_overrides.items():
        if norm_name(override_name) == weapon_key:
            allowed = {int(item_hash) for item_hash in hashes}
            return [candidate for candidate in candidates if candidate.hash in allowed]

    mode = config.weapon_version_mode.strip().lower()
    if mode == "all":
        return candidates
    if mode == "single":
        return candidates[:1]
    raise RuntimeError(
        f"未知 weapon_version_mode: {config.weapon_version_mode!r}，只能是 'single' 或 'all'"
    )


def build_group_header(weapon: str, notes: str) -> List[str]:
    header = [f"// {sanitize_comment(weapon)} - recommended"]
    if sanitize_comment(notes):
        header.append(f"//notes: {sanitize_comment(notes)}")
    return header


def _requested_perks(slot_order: List[str], slot_options: Dict[str, List[str]]) -> str:
    return " / ".join(perk for slot in slot_order for perk in slot_options.get(slot, []))


def build_wishlist(
    config: BuilderConfig,
    index: ManifestIndex,
    expanded: List[Dict[str, Any]],
) -> Tuple[List[str], List[Dict[str, Any]], List[Dict[str, Any]]]:
    policy = config.version_perk_policy.strip().lower()
    if policy not in {"drop_unsupported", "strict"}:
        raise RuntimeError(
            f"未知 version_perk_policy: {config.version_perk_policy!r}，"
            "只能是 'drop_unsupported' 或 'strict'"
        )

    lines = [
        f"title:{config.wishlist_title}",
        f"description:{config.wishlist_description}",
        "",
    ]
    unresolved: List[Dict[str, Any]] = []
    resolved_audit: List[Dict[str, Any]] = []
    groups: Dict[Tuple[int, str, str, int, str, str], List[str]] = {}
    seen_lines: Dict[Tuple[int, str, str, int, str, str], set[str]] = {}

    for recommendation in collapse_expanded_recommendations(expanded):
        row_id = recommendation["source_row"]
        weapon = recommendation["weapon"]
        notes = recommendation.get("notes", "")
        slot_order = recommendation.get("slot_order", []) or []
        slot_options = recommendation.get("parsed_perk_columns", {}) or {}
        requested_text = _requested_perks(slot_order, slot_options)

        all_candidates = index.find_weapons(weapon)
        candidates = select_weapon_candidates(config, weapon, all_candidates)
        if not all_candidates:
            unresolved.append({
                "source_row": row_id,
                "type": "weapon",
                "name": weapon,
                "reason": "weapon_not_found",
                "weapon": weapon,
                "generated": "no",
                "version_status": "skipped",
                "perks": requested_text,
            })
            continue
        if not candidates:
            unresolved.append({
                "source_row": row_id,
                "type": "weapon",
                "name": weapon,
                "reason": "weapon_override_or_selection_empty",
                "weapon": weapon,
                "generated": "no",
                "version_status": "skipped",
                "perks": requested_text,
            })
            continue

        for weapon_candidate in candidates:
            supported_by_slot: Dict[str, List[InventoryItem]] = {}
            details_by_slot_name: Dict[Tuple[str, str], Dict[str, Any]] = {}
            dropped: List[Tuple[str, str, Dict[str, Any]]] = []
            requested_count = 0

            for slot in slot_order:
                selected_items = []
                selected_hashes = set()
                for perk_name in slot_options.get(slot, []) or []:
                    requested_count += 1
                    selected, detail = index.resolve_perk_for_weapon_slot(
                        weapon_candidate.hash, slot, perk_name, slot_options
                    )
                    details_by_slot_name[slot, perk_name] = detail
                    if selected is None:
                        dropped.append((slot, perk_name, detail))
                    elif selected.hash not in selected_hashes:
                        selected_items.append(selected)
                        selected_hashes.add(selected.hash)
                supported_by_slot[slot] = selected_items

            supported_count = sum(len(items) for items in supported_by_slot.values())
            version_status = "full" if not dropped else "partial"

            if policy == "strict" and dropped:
                for slot, perk_name, detail in dropped:
                    unresolved.append(_perk_issue(
                        row_id=row_id,
                        weapon=weapon,
                        weapon_candidate=weapon_candidate,
                        slot=slot,
                        perk_name=perk_name,
                        detail=detail,
                        generated="no",
                        version_status="skipped_strict",
                        requested_count=requested_count,
                        supported_count=supported_count,
                        requested_text=requested_text,
                        reason=detail.get("reason", "perk_not_in_mapped_weapon_socket"),
                    ))
                continue

            for slot, perk_name, detail in dropped:
                issue = _perk_issue(
                    row_id=row_id,
                    weapon=weapon,
                    weapon_candidate=weapon_candidate,
                    slot=slot,
                    perk_name=perk_name,
                    detail=detail,
                    generated="yes",
                    version_status=version_status,
                    requested_count=requested_count,
                    supported_count=supported_count,
                    requested_text=requested_text,
                    reason="perk_dropped_for_incompatible_version",
                )
                issue["original_reason"] = detail.get(
                    "reason", "perk_not_in_mapped_weapon_socket"
                )
                unresolved.append(issue)

            active_slots = [slot for slot in slot_order if supported_by_slot.get(slot)]
            if not active_slots:
                unresolved.append({
                    "source_row": row_id,
                    "type": "version",
                    "name": weapon,
                    "reason": "version_has_no_compatible_recommended_perks",
                    "weapon": weapon,
                    "weapon_hash": weapon_candidate.hash,
                    "generated": "no",
                    "version_status": "skipped",
                    "requested_option_count": requested_count,
                    "supported_option_count": 0,
                    "perks": requested_text,
                })
                continue

            group_key = (
                int(row_id),
                sanitize_comment(weapon),
                sanitize_comment(notes),
                int(weapon_candidate.hash),
                sanitize_comment(weapon_candidate.name),
                version_status,
            )
            groups.setdefault(group_key, [])
            seen_lines.setdefault(group_key, set())
            combo_lists = [supported_by_slot[slot] for slot in active_slots]

            for combination in itertools.product(*combo_lists):
                selected_by_slot = dict(zip(active_slots, combination))
                perk_hashes = [selected_by_slot[slot].hash for slot in active_slots]
                line = (
                    f"dimwishlist:item={weapon_candidate.hash}&perks="
                    + ",".join(str(item_hash) for item_hash in perk_hashes)
                )
                if line in seen_lines[group_key]:
                    continue
                groups[group_key].append(line)
                seen_lines[group_key].add(line)
                resolved_audit.append(_audit_row(
                    recommendation=recommendation,
                    weapon_candidate=weapon_candidate,
                    version_status=version_status,
                    active_slots=active_slots,
                    dropped=dropped,
                    requested_count=requested_count,
                    supported_count=supported_count,
                    supported_by_slot=supported_by_slot,
                    selected_by_slot=selected_by_slot,
                    details_by_slot_name=details_by_slot_name,
                    line=line,
                ))

    wrote_any = False
    for group_key, group_lines in groups.items():
        if not group_lines:
            continue
        _, weapon_name, notes, weapon_hash, manifest_name, status = group_key
        if wrote_any:
            lines.append("")
        header_weapon = weapon_name
        if manifest_name and norm_name(manifest_name) != norm_name(weapon_name):
            header_weapon = f"{weapon_name} / {manifest_name}"
        if status == "partial":
            header_weapon += " [兼容子集]"
        lines.extend(build_group_header(f"{header_weapon} [{weapon_hash}]", notes))
        lines.extend(group_lines)
        wrote_any = True

    unique_unresolved = {}
    for row in unresolved:
        key = (
            row.get("source_row"), row.get("type"), row.get("weapon"),
            row.get("weapon_hash"), row.get("slot"), row.get("name"),
            row.get("reason"), row.get("socket_index"),
        )
        unique_unresolved.setdefault(key, row)
    return lines, list(unique_unresolved.values()), resolved_audit


def _perk_issue(
    *,
    row_id: int,
    weapon: str,
    weapon_candidate: InventoryItem,
    slot: str,
    perk_name: str,
    detail: Dict[str, Any],
    generated: str,
    version_status: str,
    requested_count: int,
    supported_count: int,
    requested_text: str,
    reason: str,
) -> Dict[str, Any]:
    return {
        "source_row": row_id,
        "type": "perk",
        "name": perk_name,
        "reason": reason,
        "weapon": weapon,
        "weapon_hash": weapon_candidate.hash,
        "slot": slot,
        "socket_index": detail.get("socket_index", ""),
        "mapping_method": detail.get("mapping_method", ""),
        "mapping_name_hits": detail.get("mapping_name_hits", 0),
        "generated": generated,
        "version_status": version_status,
        "requested_option_count": requested_count,
        "supported_option_count": supported_count,
        "perks": requested_text,
        "same_name_candidate_hashes": ";".join(
            str(value) for value in detail.get("same_name_hashes", [])
        ),
        "same_name_candidate_types": "",
        "socket_candidate_hashes": ";".join(
            str(value) for value in detail.get("candidate_hashes", [])
        ),
        "socket_candidate_names": ";".join(detail.get("candidate_names", [])),
    }


def _audit_row(
    *,
    recommendation: Dict[str, Any],
    weapon_candidate: InventoryItem,
    version_status: str,
    active_slots: List[str],
    dropped: List[Tuple[str, str, Dict[str, Any]]],
    requested_count: int,
    supported_count: int,
    supported_by_slot: Dict[str, List[InventoryItem]],
    selected_by_slot: Dict[str, InventoryItem],
    details_by_slot_name: Dict[Tuple[str, str], Dict[str, Any]],
    line: str,
) -> Dict[str, Any]:
    slot_options = recommendation.get("parsed_perk_columns", {}) or {}
    audit: Dict[str, Any] = {
        "source_row": recommendation["source_row"],
        "weapon": recommendation["weapon"],
        "weapon_hash": weapon_candidate.hash,
        "weapon_sqlite_id": weapon_candidate.sql_id,
        "manifest_name": weapon_candidate.name,
        "version_status": version_status,
        "included_slots": ";".join(active_slots),
        "dropped_perks": ";".join(perk for _, perk, _ in dropped),
        "requested_option_count": requested_count,
        "supported_option_count": supported_count,
        "dim_line": line,
    }
    for slot in SLOT_TO_ROLL_INDEX:
        requested = slot_options.get(slot, []) or []
        supported = supported_by_slot.get(slot, []) or []
        selected = selected_by_slot.get(slot)
        detail = details_by_slot_name.get((slot, selected.name), {}) if selected else {}
        audit.update({
            f"{slot}_requested_options": ";".join(requested),
            f"{slot}_supported_options": ";".join(item.name for item in supported),
            slot: selected.name if selected else "",
            f"{slot}_hash": selected.hash if selected else "",
            f"{slot}_sqlite_id": selected.sql_id if selected else "",
            f"{slot}_type": selected.item_type_display if selected else "",
            f"{slot}_socket_index": detail.get("socket_index", ""),
        })
    return audit

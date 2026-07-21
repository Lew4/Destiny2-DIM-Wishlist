"""Resolve recognized visuals inside weapon trait sockets and build DIM rules."""

from __future__ import annotations

import itertools
import re
import shutil
from collections import defaultdict
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

from .icon_config import IconBuilderConfig
from .icon_models import GlobalIconResolution, IconContext, OfficialVisual
from .icon_reports import resolve_manifest_weapon_name, select_weapon_versions
from .icon_matching import is_normal_weapon_trait
from .manifest import ManifestIndex
from .models import InventoryItem
from .utils import clean_text, norm_name, sanitize_comment, to_dim_hash


def safe_filename(text: str, max_length: int = 80) -> str:
    value = re.sub(r"[\\/:*?\"<>|\r\n]+", "_", clean_text(text))
    value = re.sub(r"\s+", " ", value).strip(" ._")
    return (value or "unnamed")[:max_length]


def strict_trait_sockets(index: ManifestIndex, weapon_hash: int) -> List[Dict[str, Any]]:
    weapon = index.by_hash.get(to_dim_hash(weapon_hash))
    if weapon is None:
        return []
    entries = (weapon.json_obj.get("sockets") or {}).get("socketEntries", []) or []
    sockets = []
    for socket_index, entry in enumerate(entries):
        hashes = list(index._collect_socket_hashes(entry))
        initial_hash = entry.get("singleInitialItemHash")
        if initial_hash not in (None, 0):
            hashes.append(to_dim_hash(initial_hash))
        items = [index.by_hash[item_hash] for item_hash in dict.fromkeys(hashes) if item_hash in index.by_hash]
        traits = [item for item in items if is_normal_weapon_trait(item)]
        if traits:
            sockets.append({
                "socket_index": socket_index,
                "socket_type_hash": to_dim_hash(entry.get("socketTypeHash", 0)),
                "candidates": sorted(
                    {item.hash: item for item in traits}.values(),
                    key=lambda item: (item.name, item.hash),
                ),
            })
    return sockets


def _socket_visual_ids(socket: Dict[str, Any], item_visual_map: Dict[int, str]) -> set[str]:
    return {
        item_visual_map[item.hash]
        for item in socket.get("candidates", [])
        if item.hash in item_visual_map
    }


def assign_trait_sockets(
    group: Sequence[IconContext],
    sockets: Sequence[Dict[str, Any]],
    resolutions: Dict[str, GlobalIconResolution],
    item_visual_map: Dict[int, str],
) -> Tuple[Dict[str, Dict[str, Any]], Dict[str, Any]]:
    desired: Dict[str, set[str]] = {"trait_3": set(), "trait_4": set()}
    for context in group:
        result = resolutions.get(context.icon_sha256)
        if result and result.accepted and result.best_visual_id:
            desired[context.slot].add(result.best_visual_id)
    if not sockets:
        return {}, {"method": "no_strict_trait_socket", "hits": 0}

    socket_visuals = {
        int(socket["socket_index"]): _socket_visual_ids(socket, item_visual_map)
        for socket in sockets
    }
    choices: List[Optional[Dict[str, Any]]] = [None] + list(sockets)
    best = None
    for trait_3_socket in choices:
        for trait_4_socket in choices:
            if (
                trait_3_socket is not None
                and trait_4_socket is not None
                and trait_3_socket["socket_index"] == trait_4_socket["socket_index"]
            ):
                continue
            trait_3_hits = len(
                desired["trait_3"]
                & socket_visuals.get(int(trait_3_socket["socket_index"]), set())
            ) if trait_3_socket else 0
            trait_4_hits = len(
                desired["trait_4"]
                & socket_visuals.get(int(trait_4_socket["socket_index"]), set())
            ) if trait_4_socket else 0
            total = trait_3_hits + trait_4_hits
            assigned = int(trait_3_socket is not None) + int(trait_4_socket is not None)
            order_penalty = (
                int(trait_3_socket["socket_index"]) if trait_3_socket else 999
            ) + (int(trait_4_socket["socket_index"]) if trait_4_socket else 999)
            candidate = total, assigned, -order_penalty, trait_3_socket, trait_4_socket
            if best is None or candidate[:3] > best[:3]:
                best = candidate
    if best is None or best[0] <= 0:
        return {}, {"method": "no_supported_global_visual_in_trait_sockets", "hits": 0}

    mapping = {}
    if best[3] is not None:
        mapping["trait_3"] = best[3]
    if best[4] is not None:
        mapping["trait_4"] = best[4]
    return mapping, {
        "method": "global_visual_coverage_assignment",
        "hits": best[0],
        "trait_3_socket": best[3]["socket_index"] if best[3] else "",
        "trait_4_socket": best[4]["socket_index"] if best[4] else "",
    }


def resolve_global_visual_in_socket(
    result: GlobalIconResolution,
    socket: Dict[str, Any],
    item_visual_map: Dict[int, str],
) -> Tuple[Optional[InventoryItem], str, List[InventoryItem]]:
    if not result.accepted or not result.best_visual_id:
        return None, "global_icon_unresolved", []
    matched = ManifestIndex._rank_perk_candidates([
        item for item in socket.get("candidates", [])
        if item_visual_map.get(item.hash) == result.best_visual_id
    ])
    if not matched:
        return None, "recognized_perk_not_supported_by_version_slot", []
    if len({norm_name(item.name) for item in matched}) > 1:
        return None, "same_visual_maps_to_multiple_names_in_socket", matched
    return matched[0], "resolved_by_global_visual_and_socket", matched


def resolve_global_visual_in_actual_trait_socket(
    result: GlobalIconResolution,
    sockets: Sequence[Dict[str, Any]],
    item_visual_map: Dict[int, str],
) -> Tuple[Optional[InventoryItem], Optional[Dict[str, Any]], List[InventoryItem]]:
    """Find a recognized perk in its one real trait socket when the XLSX column is wrong."""
    compatible = []
    for candidate_socket in sockets:
        selected, _, matched = resolve_global_visual_in_socket(
            result, candidate_socket, item_visual_map
        )
        if selected is not None:
            compatible.append((selected, candidate_socket, matched))
    if len(compatible) == 1:
        return compatible[0]
    return None, None, []


def resolve_named_perk_in_weapon_socket(
    index: ManifestIndex,
    weapon_hash: int,
    perk_name: str,
) -> Tuple[Optional[InventoryItem], Optional[Dict[str, Any]], List[InventoryItem]]:
    """Resolve an exceptional named perk in its real weapon socket."""
    weapon = index.by_hash.get(to_dim_hash(weapon_hash))
    if weapon is None:
        return None, None, []
    entries = (weapon.json_obj.get("sockets") or {}).get("socketEntries", []) or []
    for socket_index, entry in enumerate(entries):
        hashes = list(index._collect_socket_hashes(entry))
        initial_hash = entry.get("singleInitialItemHash")
        if initial_hash not in (None, 0):
            hashes.append(to_dim_hash(initial_hash))
        candidates = [
            index.by_hash[item_hash]
            for item_hash in dict.fromkeys(hashes)
            if item_hash in index.by_hash
        ]
        matched = ManifestIndex._rank_perk_candidates([
            item for item in candidates if norm_name(item.name) == norm_name(perk_name)
        ])
        if matched:
            return matched[0], {
                "socket_index": socket_index,
                "socket_type_hash": to_dim_hash(entry.get("socketTypeHash", 0)),
                "candidates": candidates,
            }, matched
    return None, None, []


def group_contexts(
    contexts: Sequence[IconContext],
) -> Dict[Tuple[int, int, str, str], List[IconContext]]:
    groups: Dict[Tuple[int, int, str, str], List[IconContext]] = defaultdict(list)
    for context in contexts:
        groups[
            context.section_index, context.excel_row, context.weapon_name, context.usage
        ].append(context)
    for group in groups.values():
        group.sort(key=lambda context: (context.slot, context.slot_position, context.source_col))
    return groups


def is_recommendation_excluded(
    config: IconBuilderConfig,
    excel_row: int,
    weapon_name: str,
    usage: str,
    recognized_names: Sequence[str],
) -> bool:
    source_key = f"{excel_row}|{norm_name(weapon_name)}|{usage.lower()}"
    recognized = {norm_name(name) for name in recognized_names}
    for identity, excluded_names in config.recommendation_exclusions.items():
        raw_row, _, remainder = identity.partition("|")
        raw_name, _, raw_usage = remainder.rpartition("|")
        candidate_key = f"{raw_row}|{norm_name(raw_name)}|{raw_usage.lower()}"
        if candidate_key == source_key and recognized & {
            norm_name(name) for name in excluded_names
        }:
            return True
    return False


def build_matches_and_wishlist(
    config: IconBuilderConfig,
    index: ManifestIndex,
    contexts: Sequence[IconContext],
    output_dir: Path,
    resolutions: Dict[str, GlobalIconResolution],
    visuals: Sequence[OfficialVisual],
    item_visual_map: Dict[int, str],
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], List[str], List[Dict[str, Any]]]:
    visual_by_id = {visual.visual_id: visual for visual in visuals}
    unresolved_dir = output_dir / config.unresolved_icon_dirname
    if config.write_diagnostics:
        unresolved_dir.mkdir(parents=True, exist_ok=True)
    matches: List[Dict[str, Any]] = []
    unresolved: List[Dict[str, Any]] = []
    audit: List[Dict[str, Any]] = []
    wishlist_lines = [
        "title:Icon-based DIM Wishlist",
        "description:Global icon recognition followed by weapon-version socket hash resolution.",
        "",
    ]

    for (_, excel_row, weapon_name, usage), group in group_contexts(contexts).items():
        weapon_type = group[0].weapon_type if group else ""
        manifest_weapon_name = resolve_manifest_weapon_name(
            config, weapon_name, weapon_type
        )
        weapon_candidates = index.find_weapons(manifest_weapon_name)
        weapons = select_weapon_versions(
            config, manifest_weapon_name, weapon_candidates, excel_row
        )
        if not weapons:
            for context in group:
                unresolved.append({
                    "excel_row": excel_row,
                    "weapon_name": weapon_name,
                    "manifest_weapon_name": manifest_weapon_name,
                    "weapon_type": context.weapon_type,
                    "usage": usage,
                    "slot": context.slot,
                    "source_cell": context.source_cell,
                    "icon_sha256": context.icon_sha256,
                    "reason": "weapon_not_found",
                    "generated": "no",
                })
            continue

        for weapon in weapons:
            sockets = strict_trait_sockets(index, weapon.hash)
            trait_group = [
                context for context in group
                if config.icon_slot_overrides.get(context.icon_sha256, context.slot)
                in {"trait_3", "trait_4"}
            ]
            slot_sockets, mapping_info = assign_trait_sockets(
                trait_group, sockets, resolutions, item_visual_map
            )
            slot_by_socket_index = {
                int(socket["socket_index"]): slot
                for slot, socket in slot_sockets.items()
            }
            has_non_trait_override = any(
                config.icon_slot_overrides.get(context.icon_sha256, context.slot)
                not in {"trait_3", "trait_4"}
                for context in group
            )
            if not slot_sockets and not has_non_trait_override:
                for context in group:
                    global_result = resolutions.get(context.icon_sha256)
                    visual = visual_by_id.get(global_result.best_visual_id) if global_result else None
                    unresolved.append({
                        "excel_row": excel_row,
                        "weapon_name": weapon_name,
                        "manifest_weapon_name": manifest_weapon_name,
                        "weapon_type": context.weapon_type,
                        "weapon_hash": weapon.hash,
                        "usage": usage,
                        "slot": context.slot,
                        "source_cell": context.source_cell,
                        "icon_sha256": context.icon_sha256,
                        "accepted": "no",
                        "reason": mapping_info.get("method", "trait_socket_mapping_failed"),
                        "generated": "no",
                        "recognized_names": " / ".join(visual.names) if visual else "",
                        "global_score": global_result.best_score if global_result else "",
                    })
                continue

            selected_by_slot: Dict[str, List[InventoryItem]] = {
                "slot_2": [], "trait_3": [], "trait_4": [],
            }
            group_match_rows = []
            for context in group:
                global_result = resolutions.get(context.icon_sha256)
                visual = visual_by_id.get(global_result.best_visual_id) if global_result else None
                effective_slot = config.icon_slot_overrides.get(
                    context.icon_sha256, context.slot
                )
                override_name = config.icon_name_overrides.get(context.icon_sha256, "")
                socket = slot_sockets.get(effective_slot)
                resolved_slot = effective_slot
                slot_corrected = False
                excluded = is_recommendation_excluded(
                    config,
                    excel_row,
                    weapon_name,
                    usage,
                    visual.names if visual else [],
                )
                if excluded:
                    selected, reason, socket_matches = (
                        None, "user_excluded_recommendation", []
                    )
                elif override_name and effective_slot not in {"trait_3", "trait_4"}:
                    selected, socket, socket_matches = resolve_named_perk_in_weapon_socket(
                        index, weapon.hash, override_name
                    )
                    reason = (
                        "resolved_by_name_override_and_weapon_socket"
                        if selected else "override_perk_not_supported_by_weapon"
                    )
                elif global_result is None:
                    selected, reason, socket_matches = None, "global_resolution_missing", []
                elif not global_result.accepted:
                    selected, reason, socket_matches = None, global_result.reason, []
                elif socket is None:
                    selected, reason, socket_matches = None, "compatible_trait_socket_not_assigned", []
                else:
                    selected, reason, socket_matches = resolve_global_visual_in_socket(
                        global_result, socket, item_visual_map
                    )
                    if selected is None and reason == "recognized_perk_not_supported_by_version_slot":
                        actual_selected, actual_socket, actual_matches = (
                            resolve_global_visual_in_actual_trait_socket(
                                global_result, sockets, item_visual_map
                            )
                        )
                        if actual_selected is not None and actual_socket is not None:
                            selected = actual_selected
                            socket = actual_socket
                            socket_matches = actual_matches
                            resolved_slot = slot_by_socket_index.get(
                                int(actual_socket["socket_index"]), effective_slot
                            )
                            slot_corrected = resolved_slot != effective_slot
                            reason = "resolved_by_global_visual_and_actual_socket"
                if selected is not None and all(
                    item.hash != selected.hash for item in selected_by_slot[resolved_slot]
                ):
                    selected_by_slot[resolved_slot].append(selected)
                row = {
                    "excel_row": excel_row,
                    "weapon_name": weapon_name,
                    "manifest_weapon_name": manifest_weapon_name,
                    "weapon_type": context.weapon_type,
                    "weapon_hash": weapon.hash,
                    "usage": usage,
                    "source_slot": context.slot,
                    "slot": resolved_slot,
                    "slot_position": context.slot_position,
                    "source_cell": context.source_cell,
                    "icon_sha256": context.icon_sha256,
                    "exported_icon": context.exported_icon,
                    "socket_index": socket.get("socket_index", "") if socket else "",
                    "mapping_method": mapping_info.get("method", ""),
                    "mapping_hits": mapping_info.get("hits", 0),
                    "slot_corrected": "yes" if slot_corrected else "no",
                    "accepted": (
                        "excluded" if excluded else "yes" if selected is not None else "no"
                    ),
                    "reason": reason,
                    "recognized_names": " / ".join(visual.names) if visual else "",
                    "global_visual_id": global_result.best_visual_id if global_result else "",
                    "global_score": global_result.best_score if global_result else "",
                    "global_margin": global_result.margin if global_result else "",
                    "global_match_method": global_result.match_method if global_result else "",
                    "selected_perk_name": selected.name if selected else "",
                    "selected_perk_hash": selected.hash if selected else "",
                    "socket_matching_names": " / ".join(item.name for item in socket_matches),
                    "socket_matching_hashes": " / ".join(str(item.hash) for item in socket_matches),
                    "socket_candidate_names": " / ".join(
                        item.name for item in socket.get("candidates", [])
                    ) if socket else "",
                    "socket_candidate_hashes": " / ".join(
                        str(item.hash) for item in socket.get("candidates", [])
                    ) if socket else "",
                }
                matches.append(row)
                group_match_rows.append(row)
                if selected is None and not excluded:
                    unresolved.append({**row, "generated": "pending"})
                    source_path = output_dir / context.exported_icon
                    if config.write_diagnostics and source_path.exists():
                        destination = unresolved_dir / (
                            f"{safe_filename(weapon_name)}__{weapon.hash}__{usage}__"
                            f"{context.slot}__{context.slot_position}__"
                            f"{context.icon_sha256[:12]}{context.icon_extension}"
                        )
                        if not destination.exists():
                            shutil.copy2(source_path, destination)

            nonempty = [
                selected_by_slot[slot]
                for slot in ("slot_2", "trait_3", "trait_4")
                if selected_by_slot[slot]
            ]
            if not nonempty:
                _mark_pending(unresolved, weapon.hash, excel_row, usage, "no")
                continue

            combinations = []
            seen_hashes = set()
            for combination in itertools.product(*nonempty):
                hashes = tuple(item.hash for item in combination)
                if hashes not in seen_hashes:
                    seen_hashes.add(hashes)
                    combinations.append(combination)
            partial = any(row["accepted"] == "no" for row in group_match_rows)
            suffix = " [兼容子集]" if partial else ""
            display_name = weapon_name
            if norm_name(manifest_weapon_name) != norm_name(weapon_name):
                display_name = f"{weapon_name} → {manifest_weapon_name}"
            wishlist_lines.append(
                f"// {sanitize_comment(display_name)} [{weapon.hash}] ({usage}){suffix}"
            )
            wishlist_lines.append(f"//notes: tags:{usage}")
            for combination in combinations:
                perk_hashes = ",".join(str(item.hash) for item in combination)
                wishlist_lines.append(f"dimwishlist:item={weapon.hash}&perks={perk_hashes}")
                audit.append({
                    "excel_row": excel_row,
                    "weapon_name": weapon_name,
                    "manifest_weapon_name": manifest_weapon_name,
                    "weapon_hash": weapon.hash,
                    "usage": usage,
                    "slot_2_names": " / ".join(item.name for item in selected_by_slot["slot_2"]),
                    "slot_2_hashes": " / ".join(str(item.hash) for item in selected_by_slot["slot_2"]),
                    "trait_3_names": " / ".join(item.name for item in selected_by_slot["trait_3"]),
                    "trait_3_hashes": " / ".join(str(item.hash) for item in selected_by_slot["trait_3"]),
                    "trait_4_names": " / ".join(item.name for item in selected_by_slot["trait_4"]),
                    "trait_4_hashes": " / ".join(str(item.hash) for item in selected_by_slot["trait_4"]),
                    "wishlist_perks": perk_hashes,
                    "combination_count": len(combinations),
                    "partial": "yes" if partial else "no",
                    "mapping_method": mapping_info.get("method", ""),
                    "mapping_hits": mapping_info.get("hits", 0),
                })
            wishlist_lines.append("")
            _mark_pending(
                unresolved, weapon.hash, excel_row, usage, "yes" if combinations else "no"
            )
    return matches, unresolved, wishlist_lines, audit


def _mark_pending(
    unresolved: List[Dict[str, Any]],
    weapon_hash: int,
    excel_row: int,
    usage: str,
    generated: str,
) -> None:
    for row in unresolved:
        if (
            row.get("weapon_hash") == weapon_hash
            and row.get("excel_row") == excel_row
            and row.get("usage") == usage
            and row.get("generated") == "pending"
        ):
            row["generated"] = generated

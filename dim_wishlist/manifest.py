"""Bungie manifest loading and weapon-specific perk resolution."""

from __future__ import annotations

import itertools
import json
import sqlite3
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

from .config import SLOT_TO_ROLL_INDEX
from .models import InventoryItem
from .utils import norm_name, to_dim_hash


def load_inventory_items(db_path: Path) -> List[InventoryItem]:
    connection = sqlite3.connect(str(db_path))
    try:
        cursor = connection.cursor()
        tables = {row[0] for row in cursor.execute(
            "SELECT name FROM sqlite_master WHERE type='table'"
        )}
        if "DestinyInventoryItemDefinition" not in tables:
            raise RuntimeError(
                "没有找到 DestinyInventoryItemDefinition。请确认传入的是 Bungie manifest SQLite 或外层压缩文件。"
            )

        items: List[InventoryItem] = []
        for raw_id, raw_json in cursor.execute(
            "SELECT id, json FROM DestinyInventoryItemDefinition"
        ):
            try:
                obj = json.loads(raw_json)
            except Exception:
                continue
            name = obj.get("displayProperties", {}).get("name", "") or ""
            if not name:
                continue
            plug_obj = obj.get("plug") or {}
            dim_hash = obj.get("hash", raw_id)
            items.append(InventoryItem(
                hash=to_dim_hash(dim_hash),
                sql_id=int(raw_id),
                name=name,
                item_type=obj.get("itemType"),
                item_type_display=obj.get("itemTypeDisplayName", "") or "",
                item_type_and_tier_display=obj.get("itemTypeAndTierDisplayName", "") or "",
                tier_type_name=(obj.get("inventory") or {}).get("tierTypeName", "") or "",
                plug_category_identifier=plug_obj.get("plugCategoryIdentifier", "") or "",
                has_plug=bool(obj.get("plug")),
                json_obj=obj,
            ))
        return items
    finally:
        connection.close()


def load_plug_sets(db_path: Path) -> Dict[int, List[int]]:
    connection = sqlite3.connect(str(db_path))
    try:
        cursor = connection.cursor()
        tables = {row[0] for row in cursor.execute(
            "SELECT name FROM sqlite_master WHERE type='table'"
        )}
        if "DestinyPlugSetDefinition" not in tables:
            raise RuntimeError("manifest 中没有 DestinyPlugSetDefinition，无法按武器 socket 解析 perk。")

        result: Dict[int, List[int]] = {}
        for raw_id, raw_json in cursor.execute("SELECT id, json FROM DestinyPlugSetDefinition"):
            try:
                obj = json.loads(raw_json)
            except Exception:
                continue
            hashes = [
                to_dim_hash(entry["plugItemHash"])
                for entry in obj.get("reusablePlugItems", []) or []
                if entry.get("plugItemHash") not in (None, 0)
            ]
            result[to_dim_hash(obj.get("hash", raw_id))] = list(dict.fromkeys(hashes))
        return result
    finally:
        connection.close()


class ManifestIndex:
    """Indexes inventory definitions and resolves perks inside a weapon socket."""

    def __init__(self, items: List[InventoryItem], plug_sets: Dict[int, List[int]]):
        self.items = items
        self.plug_sets = plug_sets
        self.by_name: Dict[str, List[InventoryItem]] = {}
        self.by_hash: Dict[int, InventoryItem] = {}
        self._roll_socket_cache: Dict[int, List[Dict[str, Any]]] = {}
        self._slot_mapping_cache: Dict[
            Tuple[int, Tuple[Tuple[str, Tuple[str, ...]], ...]], Dict[str, Dict[str, Any]]
        ] = {}
        self._perk_cache: Dict[
            Tuple[int, str, str, Tuple[Tuple[str, Tuple[str, ...]], ...]],
            Tuple[Optional[InventoryItem], Dict[str, Any]],
        ] = {}
        for item in items:
            self.by_name.setdefault(norm_name(item.name), []).append(item)
            self.by_hash[item.hash] = item

    def find_weapons(self, name: str) -> List[InventoryItem]:
        key = norm_name(name)
        exact = [item for item in self.by_name.get(key, []) if item.item_type == 3]
        if exact:
            return sorted(exact, key=lambda item: item.hash)
        fuzzy = [
            item
            for candidate_name, values in self.by_name.items()
            if key and (key in candidate_name or candidate_name in key)
            for item in values
            if item.item_type == 3
        ]
        return sorted(fuzzy, key=lambda item: (item.name, item.hash))

    @staticmethod
    def _is_enhanced_perk(item: InventoryItem) -> bool:
        text = " ".join((
            item.name,
            item.item_type_display,
            item.item_type_and_tier_display,
            item.tier_type_name,
        )).lower()
        return "强化" in text or "enhanced" in text

    @staticmethod
    def _is_normal_trait(item: InventoryItem) -> bool:
        type_and_tier = item.item_type_and_tier_display.lower()
        is_trait = (
            item.item_type_display in {"特性", "Trait"}
            or " trait" in type_and_tier
            or type_and_tier.endswith("trait")
        )
        is_common = (
            item.tier_type_name in {"普通", "Common"}
            or "普通" in type_and_tier
            or "common" in type_and_tier
        )
        return is_trait and is_common and not ManifestIndex._is_enhanced_perk(item)

    @staticmethod
    def _perk_rank(item: InventoryItem) -> Tuple[int, int]:
        if ManifestIndex._is_normal_trait(item):
            return 0, item.hash
        if item.item_type == 19 and item.has_plug and not ManifestIndex._is_enhanced_perk(item):
            return 1, item.hash
        if item.item_type == 19 and item.has_plug:
            return 2, item.hash
        if item.item_type == 19:
            return 3, item.hash
        return 9, item.hash

    @classmethod
    def _rank_perk_candidates(cls, candidates: List[InventoryItem]) -> List[InventoryItem]:
        seen = set()
        result = []
        for candidate in sorted(candidates, key=cls._perk_rank):
            if candidate.hash not in seen:
                result.append(candidate)
                seen.add(candidate.hash)
        return result

    def find_perk_hashes(self, name: str) -> List[InventoryItem]:
        candidates = [
            item for item in self.by_name.get(norm_name(name), [])
            if item.item_type == 19 and item.has_plug
        ]
        return self._rank_perk_candidates(candidates)

    def _collect_socket_hashes(self, socket_entry: Dict[str, Any]) -> List[int]:
        hashes: List[int] = []
        for key in ("randomizedPlugSetHash", "reusablePlugSetHash"):
            set_hash = socket_entry.get(key)
            if set_hash not in (None, 0):
                hashes.extend(self.plug_sets.get(to_dim_hash(set_hash), []))
        hashes.extend(
            to_dim_hash(entry["plugItemHash"])
            for entry in socket_entry.get("reusablePlugItems", []) or []
            if entry.get("plugItemHash") not in (None, 0)
        )
        return list(dict.fromkeys(hashes))

    def get_weapon_roll_sockets(self, weapon_hash: int) -> List[Dict[str, Any]]:
        weapon_hash = to_dim_hash(weapon_hash)
        if weapon_hash in self._roll_socket_cache:
            return self._roll_socket_cache[weapon_hash]
        weapon = self.by_hash.get(weapon_hash)
        if not weapon:
            return []

        randomized: List[Dict[str, Any]] = []
        fallback: List[Dict[str, Any]] = []
        entries = (weapon.json_obj.get("sockets") or {}).get("socketEntries", []) or []
        for socket_index, entry in enumerate(entries):
            candidates = [
                self.by_hash[item_hash]
                for item_hash in self._collect_socket_hashes(entry)
                if item_hash in self.by_hash
                and self.by_hash[item_hash].item_type == 19
                and self.by_hash[item_hash].has_plug
            ]
            if not candidates:
                continue
            categories = sorted({
                candidate.plug_category_identifier.lower()
                for candidate in candidates
                if candidate.plug_category_identifier
            })
            info = {
                "socket_index": socket_index,
                "socket_type_hash": to_dim_hash(entry.get("socketTypeHash", 0)),
                "randomized_plug_set_hash": to_dim_hash(entry.get("randomizedPlugSetHash", 0))
                if entry.get("randomizedPlugSetHash") else 0,
                "reusable_plug_set_hash": to_dim_hash(entry.get("reusablePlugSetHash", 0))
                if entry.get("reusablePlugSetHash") else 0,
                "candidate_hashes": [candidate.hash for candidate in candidates],
                "candidates": candidates,
                "categories": categories,
            }
            if entry.get("randomizedPlugSetHash"):
                randomized.append(info)
                continue
            excluded = (
                "frame", "intrinsic", "origin", "masterwork", "tracker", "mod",
                "memento", "shader", "cosmetic", "deepsight", "enhancement", "crafting_plug",
            )
            if not any(token in " ".join(categories) for token in excluded):
                fallback.append(info)

        sockets = randomized[:4]
        if len(sockets) < 4:
            used = {socket["socket_index"] for socket in sockets}
            for socket in fallback:
                if socket["socket_index"] not in used:
                    sockets.append(socket)
                    used.add(socket["socket_index"])
                if len(sockets) >= 4:
                    break
        self._roll_socket_cache[weapon_hash] = sockets
        return sockets

    @staticmethod
    def _slot_category_bonus(slot: str, categories: Sequence[str]) -> int:
        text = " ".join(categories).lower()
        if slot == "slot_1_barrel":
            tokens = ("barrel", "sight", "scope", "bowstring", "string", "blade", "haft")
        elif slot == "slot_2_magazine":
            tokens = ("magazine", "battery", "arrow", "fletching", "guard", "projectile", "ammo")
        else:
            tokens = ("trait", "perk")
        return 8 if any(token in text for token in tokens) else 0

    @staticmethod
    def _slot_options_signature(
        slot_options: Optional[Dict[str, Sequence[str]]],
    ) -> Tuple[Tuple[str, Tuple[str, ...]], ...]:
        slot_options = slot_options or {}
        signature = []
        for slot in SLOT_TO_ROLL_INDEX:
            values = tuple(sorted({
                norm_name(value) for value in slot_options.get(slot, []) if norm_name(value)
            }))
            if values:
                signature.append((slot, values))
        return tuple(signature)

    def infer_slot_socket_mapping(
        self,
        weapon_hash: int,
        slot_options: Optional[Dict[str, Sequence[str]]],
    ) -> Dict[str, Dict[str, Any]]:
        weapon_hash = to_dim_hash(weapon_hash)
        signature = self._slot_options_signature(slot_options)
        cache_key = weapon_hash, signature
        if cache_key in self._slot_mapping_cache:
            return self._slot_mapping_cache[cache_key]

        sockets = self.get_weapon_roll_sockets(weapon_hash)
        slots = [slot for slot, values in signature if values] or list(SLOT_TO_ROLL_INDEX)
        slots = slots[:len(sockets)]
        if not sockets or not slots:
            self._slot_mapping_cache[cache_key] = {}
            return {}

        option_map = {slot: set(values) for slot, values in signature}
        socket_names = [
            {norm_name(candidate.name) for candidate in socket["candidates"] if norm_name(candidate.name)}
            for socket in sockets
        ]
        best_perm = None
        best_objective = None
        best_scores: Dict[str, Dict[str, int]] = {}
        for permutation in itertools.permutations(range(len(sockets)), len(slots)):
            total_hits = slots_with_hits = category_bonus = order_penalty = 0
            local_scores = {}
            for slot, roll_index in zip(slots, permutation):
                hits = len(option_map.get(slot, set()) & socket_names[roll_index])
                total_hits += hits
                slots_with_hits += int(hits > 0)
                bonus = self._slot_category_bonus(slot, sockets[roll_index]["categories"])
                category_bonus += bonus
                distance = abs(SLOT_TO_ROLL_INDEX.get(slot, roll_index) - roll_index)
                order_penalty += distance
                local_scores[slot] = {
                    "name_hits": hits,
                    "category_bonus": bonus,
                    "order_distance": distance,
                }
            objective = total_hits, slots_with_hits, category_bonus, -order_penalty
            if best_objective is None or objective > best_objective:
                best_objective = objective
                best_perm = permutation
                best_scores = local_scores

        result: Dict[str, Dict[str, Any]] = {}
        if best_perm is not None:
            for slot, roll_index in zip(slots, best_perm):
                score = best_scores[slot]
                result[slot] = {
                    "roll_index": roll_index,
                    "socket_index": sockets[roll_index]["socket_index"],
                    **score,
                    "mapping_method": "name_coverage_assignment"
                    if score["name_hits"] > 0 else "ordered_fallback",
                }
        self._slot_mapping_cache[cache_key] = result
        return result

    def resolve_perk_for_weapon_slot(
        self,
        weapon_hash: int,
        slot: str,
        perk_name: str,
        slot_options: Optional[Dict[str, Sequence[str]]] = None,
    ) -> Tuple[Optional[InventoryItem], Dict[str, Any]]:
        signature = self._slot_options_signature(slot_options)
        cache_key = to_dim_hash(weapon_hash), slot, norm_name(perk_name), signature
        if cache_key in self._perk_cache:
            return self._perk_cache[cache_key]

        sockets = self.get_weapon_roll_sockets(weapon_hash)
        mapped = self.infer_slot_socket_mapping(weapon_hash, slot_options).get(slot)
        detail: Dict[str, Any] = {
            "weapon_hash": to_dim_hash(weapon_hash), "slot": slot, "roll_index": "",
            "socket_index": "", "candidate_hashes": [], "candidate_names": [],
            "same_name_hashes": [], "mapping_method": "", "mapping_name_hits": 0,
            "reason": "",
        }
        if mapped is None:
            default_index = SLOT_TO_ROLL_INDEX.get(slot)
            if default_index is None:
                detail["reason"] = "unknown_input_slot"
                return self._cache_result(cache_key, None, detail)
            if default_index >= len(sockets):
                detail["reason"] = f"weapon_has_only_{len(sockets)}_roll_sockets"
                return self._cache_result(cache_key, None, detail)
            mapped = {
                "roll_index": default_index,
                "socket_index": sockets[default_index]["socket_index"],
                "mapping_method": "ordered_fallback",
                "name_hits": 0,
            }

        roll_index = int(mapped["roll_index"])
        if roll_index >= len(sockets):
            detail["reason"] = f"mapped_roll_index_out_of_range_{roll_index}"
            return self._cache_result(cache_key, None, detail)

        socket = sockets[roll_index]
        candidates: List[InventoryItem] = socket["candidates"]
        detail.update({
            "roll_index": roll_index,
            "socket_index": socket["socket_index"],
            "candidate_hashes": [candidate.hash for candidate in candidates],
            "candidate_names": [candidate.name for candidate in candidates],
            "plug_categories": socket.get("categories", []),
            "mapping_method": mapped.get("mapping_method", ""),
            "mapping_name_hits": mapped.get("name_hits", 0),
        })
        exact = [candidate for candidate in candidates if norm_name(candidate.name) == norm_name(perk_name)]
        detail["same_name_hashes"] = [candidate.hash for candidate in exact]
        if not exact:
            detail["reason"] = "perk_not_in_mapped_weapon_socket"
            return self._cache_result(cache_key, None, detail)

        ranked = self._rank_perk_candidates(exact)
        selected = ranked[0] if ranked else None
        detail["reason"] = (
            "resolved_from_mapped_weapon_socket" if selected else "same_name_candidates_filtered_out"
        )
        if selected:
            detail["selected_hash"] = selected.hash
        return self._cache_result(cache_key, selected, detail)

    def _cache_result(
        self,
        key: Tuple[int, str, str, Tuple[Tuple[str, Tuple[str, ...]], ...]],
        selected: Optional[InventoryItem],
        detail: Dict[str, Any],
    ) -> Tuple[Optional[InventoryItem], Dict[str, Any]]:
        result = selected, detail
        self._perk_cache[key] = result
        return result

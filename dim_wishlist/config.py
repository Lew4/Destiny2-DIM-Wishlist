"""Application configuration and shared constants."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional


@dataclass(frozen=True)
class BuilderConfig:
    """All inputs and generation policies for one builder run."""

    input_path: Path = Path("./input/命运2-凯旋丰碑全种类武器推荐-Sheet1.csv")
    manifest_path: Path = Path("./input/world_sql_content_22b6eb96bbcaa631746b584b52bcc2a6.content")
    output_dir: Path = Path("./outputs")
    sheet_name: Optional[str] = None
    cache_dir: Path = Path("./.manifest_cache")
    weapon_version_mode: str = "all"
    version_perk_policy: str = "drop_unsupported"
    write_diagnostics: bool = False
    weapon_hash_overrides: Dict[str, List[int]] = field(default_factory=dict)

    wishlist_filename: str = "dim_wishlist_resolved.txt"
    unresolved_filename: str = "dim_wishlist_unresolved.csv"
    extracted_filename: str = "dim_wishlist_extracted.csv"
    resolved_audit_filename: str = "dim_wishlist_resolved_audit.csv"
    weapon_candidates_filename: str = "dim_wishlist_weapon_candidates.csv"
    perk_candidates_filename: str = "dim_wishlist_perk_candidates.csv"
    wishlist_title: str = "Weapon Wishlist Final"
    wishlist_description: str = (
        "Weapons perk that was recommended in the final season. "
        "(from bilibili @uid: 322396395)"
    )


WEAPON_COL_ALIASES = {"名字", "名称", "武器", "武器名", "name", "weapon", "item"}
NOTE_COL_ALIASES = {"注释", "备注", "说明", "notes", "note", "comment"}
TIER_COL_ALIASES = {"tier", "等级", "评级"}
RANK_COL_ALIASES = {"rank", "排序", "排名"}
EXCLUDE_PERK_HEADERS = {"原始特性", "起源特性", "origin trait", "origin"}

# Input columns map to DIM's four random-roll slots.
SLOT_TO_ROLL_INDEX = {
    "slot_1_barrel": 0,
    "slot_2_magazine": 1,
    "slot_3_trait": 2,
    "slot_4_trait": 3,
}

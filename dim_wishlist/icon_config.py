"""Configuration for XLSX files whose perk recommendations are embedded icons."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List


@dataclass(frozen=True)
class IconBuilderConfig:
    input_xlsx: Path = Path("./examples/d2.xlsx")
    manifest_path: Path = Path("./world_sql_content_22b6eb96bbcaa631746b584b52bcc2a6.content")
    output_dir: Path = Path("./outputs")
    cache_dir: Path = Path("./.manifest_cache")
    run_mode: str = "full"
    weapon_version_mode: str = "all"
    weapon_hash_overrides: Dict[str, List[int]] = field(default_factory=dict)
    icon_name_overrides: Dict[str, str] = field(default_factory=dict)

    official_icon_cache_dir: Path = Path("./.official_icon_cache")
    bungie_icon_base: str = "https://www.bungie.net"
    http_timeout_seconds: int = 20
    http_user_agent: str = "DIM-Icon-Wishlist-Builder/1.0"
    global_min_similarity: float = 0.935
    global_min_score_margin: float = 0.025
    global_top_k: int = 12
    global_expensive_top_k: int = 8
    allow_approximate_match: bool = True
    normalized_icon_size: int = 96
    content_padding: int = 8
    max_translation_pixels: int = 2

    extracted_filename: str = "icon_extracted.csv"
    matches_filename: str = "icon_matches.csv"
    unresolved_filename: str = "icon_unresolved.csv"
    wishlist_filename: str = "dim_icon_wishlist_resolved.txt"
    audit_filename: str = "dim_icon_wishlist_audit.csv"
    weapon_candidates_filename: str = "icon_weapon_candidates.csv"
    global_matches_filename: str = "icon_global_matches.csv"
    global_unresolved_filename: str = "icon_global_unresolved.csv"
    global_review_filename: str = "icon_global_review.html"
    extracted_icon_dirname: str = "extracted_icons"
    unresolved_icon_dirname: str = "unresolved_icons"


# Relative columns inside each horizontal weapon section in the icon workbook.
REL_WEAPON_NAME = 0
REL_WEAPON_TYPE = 2
REL_PVE_TRAIT3_START = 3
REL_PVE_TRAIT4_START = 7
REL_SPECIAL = 11
REL_PVP_TRAIT3_START = 12
REL_PVP_TRAIT4_START = 16
ICON_COLUMNS_PER_SLOT = 4

NS_SHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_DRAW = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

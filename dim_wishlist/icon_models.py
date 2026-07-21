"""Data models used by the icon-based XLSX workflow."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, List

import numpy as np

from .models import InventoryItem


@dataclass(frozen=True)
class DrawingImage:
    row: int
    col: int
    target: str
    media_path: str
    data: bytes
    sha256: str
    extension: str


@dataclass(frozen=True)
class IconContext:
    section_index: int
    excel_row: int
    weapon_name: str
    weapon_type: str
    usage: str
    slot: str
    slot_position: int
    source_col: int
    source_cell: str
    special_note: str
    media_path: str
    icon_sha256: str
    icon_extension: str
    icon_bytes: bytes
    exported_icon: str


@dataclass(frozen=True)
class IconLegendNote:
    usage: str
    note: str
    icon_sha256: str


@dataclass
class ImageSignature:
    file_sha256: str
    canonical_sha256: str
    rgba: np.ndarray
    alpha: np.ndarray
    gray: np.ndarray
    edge: np.ndarray
    fast_values: np.ndarray
    dhash: int


@dataclass
class OfficialVisual:
    visual_id: str
    canonical_sha256: str
    signature: ImageSignature
    items: List[InventoryItem]
    icon_paths: List[str]
    file_sha256s: List[str]

    @property
    def names(self) -> List[str]:
        return sorted({item.name for item in self.items})

    @property
    def hashes(self) -> List[int]:
        return sorted({item.hash for item in self.items})


@dataclass
class GlobalIconResolution:
    icon_sha256: str
    accepted: bool
    reason: str
    best_visual_id: str
    best_score: float
    second_score: float
    margin: float
    match_method: str
    candidate_summary: List[Dict[str, Any]]
    occurrence_count: int

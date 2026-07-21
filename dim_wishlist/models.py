"""Small data models shared by manifest and wishlist code."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, Optional


@dataclass
class InventoryItem:
    hash: int
    sql_id: int
    name: str
    item_type: Optional[int]
    item_type_display: str
    item_type_and_tier_display: str
    tier_type_name: str
    plug_category_identifier: str
    has_plug: bool
    json_obj: Dict[str, Any]

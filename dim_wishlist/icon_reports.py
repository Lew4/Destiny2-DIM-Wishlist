"""Reports produced by the icon-based XLSX workflow."""

from __future__ import annotations

import html
import json
from pathlib import Path
from typing import Any, Dict, List, Sequence

from .icon_config import IconBuilderConfig
from .icon_models import GlobalIconResolution, IconContext, OfficialVisual
from .manifest import ManifestIndex
from .utils import csv_write, norm_name


MATCH_FIELDS = [
    "excel_row", "weapon_name", "weapon_type", "weapon_hash", "usage", "source_slot", "slot",
    "slot_position", "source_cell", "icon_sha256", "exported_icon", "socket_index",
    "mapping_method", "mapping_hits", "accepted", "reason", "recognized_names",
    "global_visual_id", "global_score", "global_margin", "global_match_method",
    "selected_perk_name", "selected_perk_hash", "socket_matching_names",
    "socket_matching_hashes", "socket_candidate_names", "socket_candidate_hashes",
]


def write_extracted_report(path: Path, contexts: Sequence[IconContext]) -> None:
    csv_write(path, [{
        "section": context.section_index,
        "excel_row": context.excel_row,
        "weapon_name": context.weapon_name,
        "weapon_type": context.weapon_type,
        "usage": context.usage,
        "slot": context.slot,
        "slot_position": context.slot_position,
        "source_cell": context.source_cell,
        "source_column_index_0based": context.source_col,
        "special_note": context.special_note,
        "media_path": context.media_path,
        "icon_sha256": context.icon_sha256,
        "exported_icon": context.exported_icon,
    } for context in contexts], [
        "section", "excel_row", "weapon_name", "weapon_type", "usage", "slot",
        "slot_position", "source_cell", "source_column_index_0based", "special_note",
        "media_path", "icon_sha256", "exported_icon",
    ])


def write_global_reports(
    output_dir: Path,
    contexts: Sequence[IconContext],
    resolutions: Dict[str, GlobalIconResolution],
    visual_by_id: Dict[str, OfficialVisual],
    config: IconBuilderConfig,
) -> None:
    first_context = {context.icon_sha256: context for context in contexts}
    rows = []
    unresolved = []
    for icon_sha, result in sorted(resolutions.items()):
        context = first_context[icon_sha]
        visual = visual_by_id.get(result.best_visual_id)
        row = {
            "icon_sha256": icon_sha,
            "exported_icon": context.exported_icon,
            "occurrence_count": result.occurrence_count,
            "accepted": "yes" if result.accepted else "no",
            "reason": result.reason,
            "recognized_names": " / ".join(visual.names) if visual else "",
            "recognized_hashes": " / ".join(map(str, visual.hashes)) if visual else "",
            "best_visual_id": result.best_visual_id,
            "best_score": result.best_score,
            "second_score": result.second_score,
            "margin": result.margin,
            "match_method": result.match_method,
            "candidate_summary_json": json.dumps(result.candidate_summary, ensure_ascii=False),
        }
        rows.append(row)
        if not result.accepted:
            unresolved.append(row)
    fields = [
        "icon_sha256", "exported_icon", "occurrence_count", "accepted", "reason",
        "recognized_names", "recognized_hashes", "best_visual_id", "best_score",
        "second_score", "margin", "match_method", "candidate_summary_json",
    ]
    csv_write(output_dir / config.global_matches_filename, rows, fields)
    csv_write(output_dir / config.global_unresolved_filename, unresolved, fields)

    document = [
        "<!doctype html><meta charset='utf-8'><title>DIM 图标全局识别审核</title>",
        "<style>body{font-family:sans-serif}table{border-collapse:collapse;width:100%}"
        "td,th{border:1px solid #ccc;padding:6px;vertical-align:middle}"
        "img{width:72px;height:72px;object-fit:contain;background:#222}"
        ".bad{background:#ffe1e1}.good{background:#e8ffe8}</style>",
        "<h1>唯一图标全局识别结果</h1><table><tr><th>图标</th><th>识别</th>"
        "<th>分数</th><th>次数</th><th>原因</th></tr>",
    ]
    for row in rows:
        css_class = "good" if row["accepted"] == "yes" else "bad"
        document.append(
            f"<tr class='{css_class}'><td><img src='{html.escape(row['exported_icon'])}'>"
            f"<br><small>{html.escape(row['icon_sha256'][:16])}</small></td>"
            f"<td>{html.escape(row['recognized_names'])}</td>"
            f"<td>{row['best_score']} / margin {row['margin']}</td>"
            f"<td>{row['occurrence_count']}</td><td>{html.escape(row['reason'])}</td></tr>"
        )
    document.append("</table>")
    (output_dir / config.global_review_filename).write_text(
        "\n".join(document), encoding="utf-8"
    )


def select_weapon_versions(config: IconBuilderConfig, weapon_name: str, candidates: Sequence[Any]) -> List[Any]:
    if not candidates:
        return []
    key = norm_name(weapon_name)
    for name, hashes in config.weapon_hash_overrides.items():
        if norm_name(name) == key:
            allowed = {int(item_hash) for item_hash in hashes}
            return [item for item in candidates if item.hash in allowed]
    if config.weapon_version_mode == "all":
        return list(candidates)
    if config.weapon_version_mode == "single":
        return list(candidates[:1])
    raise RuntimeError(f"未知 weapon_version_mode：{config.weapon_version_mode!r}")


def write_weapon_report(
    path: Path,
    config: IconBuilderConfig,
    index: ManifestIndex,
    contexts: Sequence[IconContext],
) -> None:
    rows = []
    for name in dict.fromkeys(context.weapon_name for context in contexts):
        candidates = index.find_weapons(name)
        selected_hashes = {
            item.hash for item in select_weapon_versions(config, name, candidates)
        }
        if not candidates:
            rows.append({"weapon_name": name, "candidate_count": 0, "selected": "no"})
        for item in candidates:
            rows.append({
                "weapon_name": name,
                "candidate_count": len(candidates),
                "selected": "yes" if item.hash in selected_hashes else "no",
                "weapon_hash": item.hash,
                "weapon_sqlite_id": item.sql_id,
                "manifest_name": item.name,
                "item_type_display": item.item_type_display,
            })
    csv_write(path, rows, [
        "weapon_name", "candidate_count", "selected", "weapon_hash", "weapon_sqlite_id",
        "manifest_name", "item_type_display",
    ])


def write_final_reports(
    output_dir: Path,
    config: IconBuilderConfig,
    matches: List[Dict[str, Any]],
    unresolved: List[Dict[str, Any]],
    wishlist_lines: List[str],
    audit: List[Dict[str, Any]],
) -> None:
    csv_write(output_dir / config.matches_filename, matches, MATCH_FIELDS)
    csv_write(output_dir / config.unresolved_filename, unresolved, MATCH_FIELDS + ["generated"])
    csv_write(output_dir / config.audit_filename, audit, [
        "excel_row", "weapon_name", "weapon_hash", "usage", "slot_2_names",
        "slot_2_hashes", "trait_3_names", "trait_3_hashes", "trait_4_names",
        "trait_4_hashes", "wishlist_perks",
        "combination_count", "partial", "mapping_method", "mapping_hits",
    ])
    (output_dir / config.wishlist_filename).write_text(
        "\n".join(wishlist_lines).rstrip() + "\n", encoding="utf-8"
    )

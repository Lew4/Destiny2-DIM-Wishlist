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
    "excel_row", "weapon_name", "manifest_weapon_name", "weapon_type", "weapon_hash",
    "usage", "source_slot", "slot",
    "slot_position", "source_cell", "icon_sha256", "exported_icon", "socket_index",
    "mapping_method", "mapping_hits", "slot_corrected", "accepted", "reason", "recognized_names",
    "global_visual_id", "global_score", "global_margin", "global_match_method",
    "selected_perk_name", "selected_perk_hash", "socket_matching_names",
    "socket_matching_hashes", "socket_candidate_names", "socket_candidate_hashes",
]


def resolve_manifest_weapon_name(
    config: IconBuilderConfig, weapon_name: str, weapon_type: str,
) -> str:
    """Apply narrowly scoped corrections for duplicate or mistranslated Chinese names."""
    source_name = norm_name(weapon_name)
    source_type = norm_name(weapon_type)
    for identity, manifest_name in config.weapon_identity_overrides.items():
        expected_name, _, expected_type = identity.partition("|")
        if (
            norm_name(expected_name) == source_name
            and (not expected_type or norm_name(expected_type) == source_type)
        ):
            return manifest_name
    return weapon_name


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


def select_weapon_versions(
    config: IconBuilderConfig,
    weapon_name: str,
    candidates: Sequence[Any],
    excel_row: Any = None,
) -> List[Any]:
    if not candidates:
        return []
    key = norm_name(weapon_name)
    source_key = f"{excel_row}|{key}"
    for identity, hashes in config.source_weapon_hash_overrides.items():
        raw_row, _, raw_name = identity.partition("|")
        if f"{raw_row}|{norm_name(raw_name)}" == source_key:
            allowed = {int(item_hash) for item_hash in hashes}
            return [item for item in candidates if item.hash in allowed]
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
    identities = dict.fromkeys(
        (context.excel_row, context.weapon_name, context.weapon_type)
        for context in contexts
    )
    for excel_row, name, weapon_type in identities:
        manifest_name = resolve_manifest_weapon_name(config, name, weapon_type)
        candidates = index.find_weapons(manifest_name)
        selected_hashes = {
            item.hash for item in select_weapon_versions(
                config, manifest_name, candidates, excel_row
            )
        }
        if not candidates:
            rows.append({
                "excel_row": excel_row, "weapon_name": name,
                "manifest_weapon_name": manifest_name,
                "weapon_type": weapon_type, "candidate_count": 0, "selected": "no",
            })
        for item in candidates:
            rows.append({
                "weapon_name": name,
                "excel_row": excel_row,
                "manifest_weapon_name": manifest_name,
                "weapon_type": weapon_type,
                "candidate_count": len(candidates),
                "selected": "yes" if item.hash in selected_hashes else "no",
                "weapon_hash": item.hash,
                "weapon_sqlite_id": item.sql_id,
                "manifest_name": item.name,
                "item_type_display": item.item_type_display,
            })
    csv_write(path, rows, [
        "excel_row", "weapon_name", "manifest_weapon_name", "weapon_type",
        "candidate_count", "selected",
        "weapon_hash", "weapon_sqlite_id", "manifest_name", "item_type_display",
    ])


def _source_key(row: Dict[str, Any]) -> tuple[Any, ...]:
    return (
        row.get("excel_row"), row.get("weapon_name"), row.get("weapon_type"),
        row.get("usage"), row.get("source_cell"), row.get("icon_sha256"),
    )


def _group_key(row: Dict[str, Any]) -> tuple[Any, ...]:
    return (
        row.get("excel_row"), row.get("weapon_name"), row.get("weapon_type"),
        row.get("usage"),
    )


def classify_version_results(
    matches: List[Dict[str, Any]], unresolved: List[Dict[str, Any]],
) -> tuple[List[Dict[str, Any]], List[Dict[str, Any]], List[Dict[str, Any]]]:
    """Separate real source issues from harmless same-name historical-version misses."""
    accepted_sources = {
        _source_key(row) for row in matches if row.get("accepted") == "yes"
    }
    excluded_sources = {
        _source_key(row) for row in matches if row.get("accepted") == "excluded"
    }
    real_unresolved_by_source: Dict[tuple[Any, ...], Dict[str, Any]] = {}
    historical_candidates = []
    for row in unresolved:
        key = _source_key(row)
        if key in excluded_sources:
            continue
        if key in accepted_sources:
            historical_candidates.append(row)
        else:
            real_unresolved_by_source.setdefault(key, row)

    all_rows = [
        row for row in matches + unresolved
        if row.get("accepted") in {"yes", "no"}
        and _source_key(row) not in excluded_sources
    ]
    expected_by_group: Dict[tuple[Any, ...], set[tuple[Any, ...]]] = {}
    accepted_by_version: Dict[tuple[tuple[Any, ...], Any], set[tuple[Any, ...]]] = {}
    versions_by_group: Dict[tuple[Any, ...], set[Any]] = {}
    example_by_group: Dict[tuple[Any, ...], Dict[str, Any]] = {}
    for row in all_rows:
        group = _group_key(row)
        source = _source_key(row)
        weapon_hash = row.get("weapon_hash")
        expected_by_group.setdefault(group, set()).add(source)
        versions_by_group.setdefault(group, set()).add(weapon_hash)
        example_by_group.setdefault(group, row)
        if row.get("accepted") == "yes":
            accepted_by_version.setdefault((group, weapon_hash), set()).add(source)

    source_issue_rows = []
    history_summary_rows = []
    full_group_keys = set()
    for group, expected in expected_by_group.items():
        version_hashes = sorted(
            (value for value in versions_by_group[group] if value not in (None, "")),
            key=lambda value: int(value),
        )
        full = []
        partial = []
        unsupported = []
        best_coverage = 0
        for weapon_hash in version_hashes:
            count = len(accepted_by_version.get((group, weapon_hash), set()))
            best_coverage = max(best_coverage, count)
            if count == len(expected):
                full.append(weapon_hash)
            elif count:
                partial.append(weapon_hash)
            else:
                unsupported.append(weapon_hash)
        example = example_by_group[group]
        row = {
            "excel_row": group[0],
            "weapon_name": group[1],
            "manifest_weapon_name": example.get("manifest_weapon_name", group[1]),
            "weapon_type": group[2],
            "usage": group[3],
            "recommendation_count": len(expected),
            "best_matched_count": best_coverage,
            "full_version_hashes": " / ".join(map(str, full)),
            "partial_version_hashes": " / ".join(map(str, partial)),
            "unsupported_version_hashes": " / ".join(map(str, unsupported)),
        }
        if full:
            full_group_keys.add(group)
            if partial or unsupported:
                history_summary_rows.append(row)
        else:
            missing_names = sorted({
                item.get("recognized_names", "")
                for item in real_unresolved_by_source.values()
                if _group_key(item) == group and item.get("recognized_names")
            })
            row["reason"] = (
                "recommendations_split_across_versions"
                if not missing_names else "recommendations_not_supported_by_any_complete_version"
            )
            row["unmatched_recommendations"] = " / ".join(missing_names)
            source_issue_rows.append(row)

    historical_rows = [
        row for row in historical_candidates if _group_key(row) in full_group_keys
    ]
    return list(real_unresolved_by_source.values()), historical_rows, source_issue_rows + history_summary_rows


def write_final_reports(
    output_dir: Path,
    config: IconBuilderConfig,
    matches: List[Dict[str, Any]],
    unresolved: List[Dict[str, Any]],
    wishlist_lines: List[str],
    audit: List[Dict[str, Any]],
) -> Dict[str, int]:
    real_unresolved, historical_rows, group_rows = classify_version_results(
        matches, unresolved
    )
    excluded_by_source: Dict[tuple[Any, ...], Dict[str, Any]] = {}
    for row in matches:
        if row.get("accepted") == "excluded":
            excluded_by_source.setdefault(_source_key(row), row)
    source_issue_rows = [row for row in group_rows if row.get("reason")]
    history_summary_rows = [row for row in group_rows if not row.get("reason")]
    summary_fields = [
        "excel_row", "weapon_name", "manifest_weapon_name", "weapon_type", "usage",
        "recommendation_count", "best_matched_count", "full_version_hashes",
        "partial_version_hashes", "unsupported_version_hashes", "reason",
        "unmatched_recommendations",
    ]
    if config.write_diagnostics:
        csv_write(output_dir / config.matches_filename, matches, MATCH_FIELDS)
        csv_write(
            output_dir / config.excluded_recommendations_filename,
            list(excluded_by_source.values()), MATCH_FIELDS,
        )
        csv_write(
            output_dir / config.unresolved_filename, real_unresolved,
            MATCH_FIELDS + ["generated"],
        )
        csv_write(
            output_dir / config.history_compatibility_filename, historical_rows,
            MATCH_FIELDS + ["generated"],
        )
        csv_write(
            output_dir / config.source_issues_filename, source_issue_rows, summary_fields,
        )
        csv_write(
            output_dir / config.history_summary_filename, history_summary_rows, summary_fields,
        )
        csv_write(output_dir / config.audit_filename, audit, [
            "excel_row", "weapon_name", "manifest_weapon_name", "weapon_hash", "usage",
            "slot_2_names", "slot_2_hashes", "trait_3_names", "trait_3_hashes",
            "trait_4_names", "trait_4_hashes", "wishlist_perks", "combination_count",
            "notes", "partial", "mapping_method", "mapping_hits",
        ])
    (output_dir / config.wishlist_filename).write_text(
        "\n".join(wishlist_lines).rstrip() + "\n", encoding="utf-8"
    )
    return {
        "source_unresolved": len(real_unresolved),
        "source_issue_groups": len(source_issue_rows),
        "historical_version_rows": len(historical_rows),
        "historical_groups": len(history_summary_rows),
        "excluded_recommendations": len(excluded_by_source),
    }

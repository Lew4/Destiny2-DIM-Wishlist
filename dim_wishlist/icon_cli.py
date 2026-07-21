"""CLI and orchestration for icon-based recommendation workbooks."""

from __future__ import annotations

import argparse
import json
import shutil
from dataclasses import replace
from pathlib import Path
from typing import Optional, Sequence

from .icon_config import IconBuilderConfig
from .icon_matching import build_official_visual_catalog, resolve_global_icons
from .icon_models import ImageSignature
from .icon_reports import (
    write_extracted_report,
    write_final_reports,
    write_global_reports,
    write_weapon_report,
)
from .icon_wishlist import build_matches_and_wishlist
from .icon_xlsx import extract_icon_contexts
from .manifest import ManifestIndex, load_inventory_items, load_plug_sets
from .utils import csv_write, ensure_sqlite_manifest


def build_parser() -> argparse.ArgumentParser:
    defaults = IconBuilderConfig()
    parser = argparse.ArgumentParser(
        description="从嵌入perk图标的XLSX生成PVE/PVP DIM Wishlist。"
    )
    parser.add_argument("--input", type=Path, default=defaults.input_xlsx, help="图标型XLSX")
    parser.add_argument("--manifest", type=Path, default=defaults.manifest_path)
    parser.add_argument("--output-dir", type=Path, default=defaults.output_dir)
    parser.add_argument("--cache-dir", type=Path, default=defaults.cache_dir)
    parser.add_argument(
        "--official-icon-cache", type=Path, default=defaults.official_icon_cache_dir,
    )
    parser.add_argument("--run-mode", choices=("full", "extract_only"), default=defaults.run_mode)
    parser.add_argument(
        "--weapon-version-mode", choices=("all", "single"),
        default=defaults.weapon_version_mode,
    )
    parser.add_argument(
        "--min-similarity", type=float, default=defaults.global_min_similarity,
    )
    parser.add_argument(
        "--min-score-margin", type=float, default=defaults.global_min_score_margin,
    )
    parser.add_argument(
        "--no-approximate-match", action="store_true",
        help="只接受文件或归一化像素完全一致的图标",
    )
    parser.add_argument(
        "--diagnostics", action="store_true",
        help="额外保留CSV/HTML/提取图标；默认outputs中只保留最终Wishlist",
    )
    return parser


def config_from_args(argv: Optional[Sequence[str]] = None) -> IconBuilderConfig:
    args = build_parser().parse_args(argv)
    return replace(
        IconBuilderConfig(),
        input_xlsx=args.input,
        manifest_path=args.manifest,
        output_dir=args.output_dir,
        cache_dir=args.cache_dir,
        official_icon_cache_dir=args.official_icon_cache,
        run_mode=args.run_mode,
        weapon_version_mode=args.weapon_version_mode,
        global_min_similarity=args.min_similarity,
        global_min_score_margin=args.min_score_margin,
        allow_approximate_match=not args.no_approximate_match,
        write_diagnostics=args.diagnostics or args.run_mode == "extract_only",
    )


def run(config: IconBuilderConfig) -> int:
    xlsx_path = config.input_xlsx.expanduser().resolve()
    output_dir = config.output_dir.expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    if not config.write_diagnostics:
        diagnostic_files = (
            config.extracted_filename, config.matches_filename,
            config.unresolved_filename, config.audit_filename,
            config.weapon_candidates_filename, config.global_matches_filename,
            config.global_unresolved_filename, config.global_review_filename,
            config.history_compatibility_filename, config.history_summary_filename,
            config.source_issues_filename, config.excluded_recommendations_filename,
            "icon_extraction_stats.json", "official_icon_catalog_errors.csv",
        )
        for filename in diagnostic_files:
            path = output_dir / filename
            if path.is_file():
                path.unlink()
        for dirname in (config.extracted_icon_dirname, config.unresolved_icon_dirname):
            path = output_dir / dirname
            if path.is_dir():
                shutil.rmtree(path)

    contexts, stats = extract_icon_contexts(xlsx_path, output_dir, config)
    if config.write_diagnostics:
        write_extracted_report(output_dir / config.extracted_filename, contexts)
        (output_dir / "icon_extraction_stats.json").write_text(
            json.dumps(stats, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    print(f"[提取] 图片位置: {stats['drawing_count']}")
    print(f"[提取] 唯一图片: {stats['unique_media_count']}")
    print(f"[提取] perk图标位置: {stats['perk_icon_position_count']}")
    print(f"[提取] perk唯一图标: {stats['unique_perk_icon_count']}")
    if config.write_diagnostics:
        print(f"[输出] {output_dir / config.extracted_filename}")

    if config.run_mode == "extract_only":
        print("run_mode=extract_only：已停止在图标位置提取阶段。")
        return 0

    db_path = ensure_sqlite_manifest(
        config.manifest_path.expanduser().resolve(),
        config.cache_dir.expanduser().resolve(),
    )
    print(f"[manifest] SQLite: {db_path}")
    index = ManifestIndex(load_inventory_items(db_path), load_plug_sets(db_path))
    if config.write_diagnostics:
        write_weapon_report(
            output_dir / config.weapon_candidates_filename, config, index, contexts
        )

    official_cache = config.official_icon_cache_dir.expanduser()
    if not official_cache.is_absolute():
        official_cache = output_dir / official_cache
    signature_cache: dict[str, ImageSignature] = {}
    print("[全局图标] 正在建立官方普通特性图标库……")
    visuals, item_visual_map, catalog_errors = build_official_visual_catalog(
        index, official_cache, signature_cache, config
    )
    if config.write_diagnostics:
        csv_write(output_dir / "official_icon_catalog_errors.csv", catalog_errors, [
            "icon_path", "item_hashes", "item_names", "reason",
        ])
    if not visuals:
        raise RuntimeError("官方普通特性图标库为空。请检查网络、manifest语言/版本和图标缓存。")
    print(f"[全局图标] 官方唯一视觉图标: {len(visuals)}")

    resolutions = resolve_global_icons(contexts, visuals, signature_cache, config)
    if config.write_diagnostics:
        write_global_reports(
            output_dir,
            contexts,
            resolutions,
            {visual.visual_id: visual for visual in visuals},
            config,
        )
    accepted_unique = sum(result.accepted for result in resolutions.values())
    print(f"[全局图标] 唯一图标识别成功: {accepted_unique}/{len(resolutions)}")

    matches, unresolved, wishlist_lines, audit = build_matches_and_wishlist(
        config, index, contexts, output_dir, resolutions, visuals, item_visual_map,
    )
    report_stats = write_final_reports(
        output_dir, config, matches, unresolved, wishlist_lines, audit
    )
    accepted_count = sum(row.get("accepted") == "yes" for row in matches)
    print(f"[武器槽位] 成功图标位置: {accepted_count}/{len(matches)}")
    print(
        f"[源推荐] 真正未解决图标: {report_stats['source_unresolved']}，"
        f"问题推荐组: {report_stats['source_issue_groups']}，"
        f"明确排除: {report_stats['excluded_recommendations']}"
    )
    print(
        f"[历史版本] 单独归档: {report_stats['historical_version_rows']} 条，"
        f"涉及 {report_stats['historical_groups']} 个已有完整版本的推荐组"
    )
    recommendation_notes = {
        (context.section_index, context.excel_row, context.usage)
        for context in contexts if context.recommendation_note
    }
    print(f"[备注] 武器 PVE/PVP 说明: {len(recommendation_notes)} 条")
    if config.write_diagnostics:
        print(f"[输出] {output_dir / config.global_review_filename}")
    print(f"[输出] {output_dir / config.wishlist_filename}")
    return 0


def main(argv: Optional[Sequence[str]] = None) -> int:
    return run(config_from_args(argv))

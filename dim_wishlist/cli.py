"""Command-line entry point and application orchestration."""

from __future__ import annotations

import argparse
from dataclasses import replace
from pathlib import Path
from typing import Optional, Sequence

from .config import BuilderConfig
from .manifest import ManifestIndex, load_inventory_items, load_plug_sets
from .reports import (
    write_extracted_report,
    write_generation_outputs,
    write_perk_candidate_report,
    write_weapon_candidate_report,
)
from .table import detect_columns, expand_rows, read_table
from .utils import ensure_sqlite_manifest
from .wishlist import build_wishlist


def build_parser() -> argparse.ArgumentParser:
    defaults = BuilderConfig()
    parser = argparse.ArgumentParser(
        description="把中文 Destiny 2 武器推荐表转换为 DIM Wishlist。"
    )
    parser.add_argument("--input", type=Path, default=defaults.input_path, help="CSV/XLSX 推荐表")
    parser.add_argument("--manifest", type=Path, default=defaults.manifest_path, help="Bungie manifest")
    parser.add_argument("--output-dir", type=Path, default=defaults.output_dir, help="输出目录")
    parser.add_argument("--sheet", default=defaults.sheet_name, help="XLSX 工作表名")
    parser.add_argument("--cache-dir", type=Path, default=defaults.cache_dir, help="manifest 解压缓存")
    parser.add_argument(
        "--weapon-version-mode", choices=("all", "single"),
        default=defaults.weapon_version_mode,
    )
    parser.add_argument(
        "--version-perk-policy", choices=("drop_unsupported", "strict"),
        default=defaults.version_perk_policy,
    )
    parser.add_argument("--title", default=defaults.wishlist_title, help="Wishlist 标题")
    parser.add_argument(
        "--description", default=defaults.wishlist_description, help="Wishlist 描述"
    )
    parser.add_argument(
        "--diagnostics", action="store_true",
        help="额外保留CSV审计报告；默认outputs中只保留最终Wishlist",
    )
    return parser


def config_from_args(argv: Optional[Sequence[str]] = None) -> BuilderConfig:
    args = build_parser().parse_args(argv)
    return replace(
        BuilderConfig(),
        input_path=args.input,
        manifest_path=args.manifest,
        output_dir=args.output_dir,
        sheet_name=args.sheet,
        cache_dir=args.cache_dir,
        weapon_version_mode=args.weapon_version_mode,
        version_perk_policy=args.version_perk_policy,
        wishlist_title=args.title,
        wishlist_description=args.description,
        write_diagnostics=args.diagnostics,
    )


def run(config: BuilderConfig) -> int:
    input_path = config.input_path.expanduser().resolve()
    manifest_path = config.manifest_path.expanduser().resolve()
    output_dir = config.output_dir.expanduser().resolve()
    cache_dir = config.cache_dir.expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    if not config.write_diagnostics:
        for filename in (
            config.unresolved_filename,
            config.extracted_filename,
            config.resolved_audit_filename,
            config.weapon_candidates_filename,
            config.perk_candidates_filename,
        ):
            path = output_dir / filename
            if path.is_file():
                path.unlink()

    print("[CONFIG]")
    print(f"      input: {input_path}")
    print(f"      manifest: {manifest_path}")
    print(f"      output_dir: {output_dir}")
    print(f"      weapon_version_mode: {config.weapon_version_mode}")
    print(f"      version_perk_policy: {config.version_perk_policy}\n")

    print(f"[1/5] 准备 manifest: {manifest_path}")
    sqlite_path = ensure_sqlite_manifest(manifest_path, cache_dir)
    print(f"      SQLite: {sqlite_path}")

    print("[2/5] 读取 DestinyInventoryItemDefinition ...")
    items = load_inventory_items(sqlite_path)
    plug_sets = load_plug_sets(sqlite_path)
    index = ManifestIndex(items, plug_sets)
    print(f"      已读取 {len(items)} 个 InventoryItem")
    print(f"      已读取 {len(plug_sets)} 个 PlugSet")

    print(f"[3/5] 读取推荐表: {input_path}")
    headers, rows = read_table(input_path, sheet=config.sheet_name)
    weapon_col, perk_cols, note_col, tier_col, rank_col = detect_columns(headers)
    print(f"      武器列: {weapon_col}")
    print(f"      Perk列: {', '.join(perk_cols)}")
    expanded = expand_rows(rows, weapon_col, perk_cols, note_col, tier_col, rank_col)
    print(f"      展开后 roll 组合: {len(expanded)}")

    if config.write_diagnostics:
        extracted_path = output_dir / config.extracted_filename
        weapon_candidates_path = output_dir / config.weapon_candidates_filename
        perk_candidates_path = output_dir / config.perk_candidates_filename
        write_extracted_report(extracted_path, expanded)
        write_weapon_candidate_report(weapon_candidates_path, config, index, expanded)
        write_perk_candidate_report(perk_candidates_path, config, index, expanded)

    print("[4/5] 匹配 hash 并生成 DIM wishlist ...")
    lines, unresolved, audit = build_wishlist(config, index, expanded)
    write_generation_outputs(config, lines, unresolved, audit)

    resolved_count = sum(line.startswith("dimwishlist:") for line in lines)
    print("[5/5] 完成")
    print(f"      resolved lines: {resolved_count}")
    print(f"      unresolved rows: {len(unresolved)}")
    print(f"      output_dir: {output_dir}")
    return 0


def main(argv: Optional[Sequence[str]] = None) -> int:
    return run(config_from_args(argv))

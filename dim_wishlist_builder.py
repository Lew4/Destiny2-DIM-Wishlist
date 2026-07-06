#!/usr/bin/env python3
"""
DIM 中文推荐表转愿望单工具。

输入中文 Excel/CSV 推荐表 + Bungie manifest SQLite/zip content，输出 DIM wishlist txt。

核心规则：
- 武器列默认识别：名字 / 武器 / name / weapon
- perk 列默认识别：列名包含 Perk / perk / PERK
- 单元格多行 perk 自动展开组合
- 同名多版本武器会全部生成
- 未匹配武器/perk 输出到 unresolved CSV
"""
from __future__ import annotations

import argparse
import csv
import itertools
import json
import os
import re
import sqlite3
import sys
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

try:
    from openpyxl import load_workbook
except Exception:  # pragma: no cover
    load_workbook = None

UINT32 = 2 ** 32

WEAPON_COL_ALIASES = {"名字", "名称", "武器", "武器名", "name", "weapon", "item"}
NOTE_COL_ALIASES = {"注释", "备注", "说明", "notes", "note", "comment"}
TIER_COL_ALIASES = {"tier", "等级", "评级"}
RANK_COL_ALIASES = {"rank", "排序", "排名"}
EXCLUDE_PERK_HEADERS = {"原始特性", "起源特性", "origin trait", "origin"}

# 常见中文全角/半角分隔符。换行最重要。
SPLIT_RE = re.compile(r"[\n\r/／、，,;；|｜]+")


def norm_name(s: Any) -> str:
    if s is None:
        return ""
    text = str(s).strip()
    text = text.replace("\u3000", " ")
    text = text.lower()
    text = re.sub(r"[\s\-_'\"“”‘’·•:：()（）\[\]【】{}<>《》]+", "", text)
    return text


def clean_text(s: Any) -> str:
    if s is None:
        return ""
    return str(s).strip()




def make_unique_headers(raw_headers: Sequence[Any]) -> List[str]:
    """保留重复列名。CSV/Excel 里常见多个 Perk 列，直接 DictReader 会覆盖前面的 Perk。"""
    counts: Dict[str, int] = {}
    headers: List[str] = []
    for i, h in enumerate(raw_headers):
        base = clean_text(h) or f"col_{i+1}"
        counts[base] = counts.get(base, 0) + 1
        if counts[base] == 1:
            headers.append(base)
        else:
            headers.append(f"{base}__{counts[base]}")
    return headers


def looks_like_header(row: Sequence[Any]) -> bool:
    vals = [clean_text(x) for x in row]
    has_weapon = any(norm_name(v) in {norm_name(a) for a in WEAPON_COL_ALIASES} for v in vals)
    has_perk = any("perk" in v.lower() for v in vals)
    return has_weapon and has_perk


def row_to_dict(headers: Sequence[str], raw: Sequence[Any]) -> Dict[str, Any]:
    return {headers[i]: raw[i] if i < len(raw) else None for i in range(len(headers))}


def split_options(cell: Any) -> List[str]:
    text = clean_text(cell)
    if not text:
        return []
    # 去掉项目符号和多余空格。
    parts = [p.strip(" \t-—*·•") for p in SPLIT_RE.split(text)]
    return [p for p in parts if p]


def to_dim_hash(value: Any) -> int:
    """Bungie manifest 的 sqlite id 有时是 signed int32，DIM 需要 uint32 形式。"""
    x = int(value)
    if x < 0:
        x += UINT32
    return x


def sql_quote_identifier(name: str) -> str:
    return '"' + name.replace('"', '""') + '"'


def ensure_sqlite_manifest(path: Path, workdir: Optional[Path] = None) -> Path:
    """返回真正可查询的 SQLite 数据库路径。

    Bungie 的 .content 常见是 zip；有时 sqlite3 shell 直接打开只看到 zip 表。
    这里优先按 zip 解压；如果不是 zip，就直接返回原路径。
    """
    path = path.expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"manifest 不存在: {path}")

    if zipfile.is_zipfile(path):
        out_dir = workdir or Path(tempfile.mkdtemp(prefix="bungie_manifest_"))
        out_dir.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(path, "r") as zf:
            names = [n for n in zf.namelist() if not n.endswith("/")]
            if not names:
                raise RuntimeError(f"zip manifest 为空: {path}")
            zf.extractall(out_dir)
            # 通常只有一个文件。
            extracted = out_dir / names[0]
            return extracted.resolve()
    return path


@dataclass
class InvItem:
    hash: int
    name: str
    item_type: Optional[int]
    item_type_display: str
    has_plug: bool
    json_obj: Dict[str, Any]


def load_inventory_items(db_path: Path) -> List[InvItem]:
    con = sqlite3.connect(str(db_path))
    cur = con.cursor()
    tables = [r[0] for r in cur.execute("SELECT name FROM sqlite_master WHERE type='table'")]
    if "DestinyInventoryItemDefinition" not in tables:
        raise RuntimeError(
            "没有找到 DestinyInventoryItemDefinition。"
            "如果 .tables 只看到 zip，说明你传入的是外层压缩文件；脚本通常会自动解压，"
            "但也可以手动 unzip 后再传入解压出的文件。"
        )

    rows = cur.execute("SELECT id, json FROM DestinyInventoryItemDefinition").fetchall()
    items: List[InvItem] = []
    for raw_id, raw_json in rows:
        try:
            obj = json.loads(raw_json)
        except Exception:
            continue
        name = obj.get("displayProperties", {}).get("name", "") or ""
        if not name:
            continue
        items.append(
            InvItem(
                hash=to_dim_hash(raw_id),
                name=name,
                item_type=obj.get("itemType"),
                item_type_display=obj.get("itemTypeAndTierDisplayName", "") or "",
                has_plug=bool(obj.get("plug")),
                json_obj=obj,
            )
        )
    con.close()
    return items


class ManifestIndex:
    def __init__(self, items: List[InvItem]):
        self.items = items
        self.by_name: Dict[str, List[InvItem]] = {}
        for it in items:
            self.by_name.setdefault(norm_name(it.name), []).append(it)

    def find_weapons(self, name: str) -> List[InvItem]:
        key = norm_name(name)
        exact = self.by_name.get(key, [])
        weapons = [x for x in exact if x.item_type == 3]
        if weapons:
            return sorted(weapons, key=lambda x: x.hash)
        # 保守的模糊匹配：仅当归一化名互相包含，并且是武器。
        fuzzy = [x for k, vals in self.by_name.items() if (key and (key in k or k in key)) for x in vals if x.item_type == 3]
        return sorted(fuzzy, key=lambda x: (x.name, x.hash))

    def find_perk_hashes(self, name: str) -> List[InvItem]:
        key = norm_name(name)
        exact = self.by_name.get(key, [])
        # 优先选择 InventoryItem 中的 plug。DIM 文档也要求用 InventoryItem，而不是 SandboxPerk。
        plugs = [x for x in exact if x.has_plug]
        if plugs:
            return self._dedupe_perk_candidates(plugs)
        # 一些 perk 可能没有 plug 字段，次选 itemType 19。
        type19 = [x for x in exact if x.item_type == 19]
        if type19:
            return self._dedupe_perk_candidates(type19)
        # 最后才允许精确同名结果。
        if exact:
            return self._dedupe_perk_candidates(exact)
        return []

    @staticmethod
    def _dedupe_perk_candidates(cands: List[InvItem]) -> List[InvItem]:
        """同名 perk 多版本时，尽量选普通版，避免 enhanced/强化版膨胀。"""
        if not cands:
            return []
        # 中文/英文常见 enhanced 标记，普通版优先。
        bad_words = ("强化", "enhanced")
        normal = [c for c in cands if not any(w in c.name.lower() for w in bad_words)]
        selected = normal or cands
        seen = set()
        out = []
        for c in sorted(selected, key=lambda x: x.hash):
            if c.hash not in seen:
                out.append(c)
                seen.add(c.hash)
        return out


def read_table(input_path: Path, sheet: Optional[str] = None) -> Tuple[List[str], List[Dict[str, Any]]]:
    suffix = input_path.suffix.lower()

    def build_from_rows(all_rows: List[List[Any]]) -> Tuple[List[str], List[Dict[str, Any]]]:
        # 你的表是“武器大类行 + 表头行 + 数据行 + 下一个武器大类行 + 表头行...”。
        # 所以不能只用第一行当表头，要在前若干行里找真正含“名字”和“Perk”的表头。
        header_idx = None
        for i, row in enumerate(all_rows[:80]):
            if looks_like_header(row):
                header_idx = i
                break
        if header_idx is None:
            preview = [[clean_text(x) for x in r] for r in all_rows[:8]]
            raise RuntimeError(
                "没有识别到表头。请确保某一行同时包含 `名字` 列和至少一个 `Perk` 列。"
                f"\n前几行内容: {preview}"
            )

        headers = make_unique_headers(all_rows[header_idx])
        rows: List[Dict[str, Any]] = []
        for raw in all_rows[header_idx + 1:]:
            vals = [clean_text(x) for x in raw]
            if not any(vals):
                continue
            # 跳过后续武器类别之间重复出现的表头行。
            if looks_like_header(raw):
                continue
            row = row_to_dict(headers, raw)
            rows.append(row)
        return headers, rows

    if suffix in {".xlsx", ".xlsm"}:
        if load_workbook is None:
            raise RuntimeError("读取 xlsx 需要 openpyxl：python3 -m pip install openpyxl")
        wb = load_workbook(str(input_path), data_only=True)
        ws = wb[sheet] if sheet else wb.active
        all_rows = [[cell for cell in row] for row in ws.iter_rows(values_only=True)]
        return build_from_rows(all_rows)

    if suffix == ".csv":
        # 不用 csv.DictReader，因为你的文件第一行是“自动步枪”这种分类标题，
        # 而且存在多个同名 Perk 列，DictReader 会把重复列覆盖掉。
        with input_path.open("r", encoding="utf-8-sig", newline="") as f:
            all_rows = list(csv.reader(f))
        return build_from_rows(all_rows)

    raise RuntimeError(f"暂不支持输入格式: {suffix}")


def detect_columns(headers: Sequence[str]) -> Tuple[str, List[str], Optional[str], Optional[str], Optional[str]]:
    weapon_col = None
    for h in headers:
        if norm_name(h) in {norm_name(a) for a in WEAPON_COL_ALIASES}:
            weapon_col = h
            break
    if not weapon_col:
        raise RuntimeError(f"没有找到武器名列。可用列名: {headers}")

    perk_cols = []
    for h in headers:
        base_h = h.split("__", 1)[0]
        hn = norm_name(base_h)
        if base_h in EXCLUDE_PERK_HEADERS or hn in {norm_name(x) for x in EXCLUDE_PERK_HEADERS}:
            continue
        if "perk" in base_h.lower():
            perk_cols.append(h)
    if not perk_cols:
        raise RuntimeError("没有找到 Perk 列。列名需要包含 `Perk`。")

    note_col = next((h for h in headers if norm_name(h) in {norm_name(a) for a in NOTE_COL_ALIASES}), None)
    tier_col = next((h for h in headers if norm_name(h) in {norm_name(a) for a in TIER_COL_ALIASES}), None)
    rank_col = next((h for h in headers if norm_name(h) in {norm_name(a) for a in RANK_COL_ALIASES}), None)
    return weapon_col, perk_cols, note_col, tier_col, rank_col


def make_note(row: Dict[str, Any], note_col: Optional[str], tier_col: Optional[str], rank_col: Optional[str]) -> str:
    parts = []
    if tier_col and clean_text(row.get(tier_col)):
        parts.append(f"Tier {clean_text(row.get(tier_col))}")
    if rank_col and clean_text(row.get(rank_col)):
        parts.append(f"Rank {clean_text(row.get(rank_col))}")
    if note_col and clean_text(row.get(note_col)):
        parts.append(clean_text(row.get(note_col)).replace("\n", " "))
    return " | ".join(parts)


def expand_rows(rows: List[Dict[str, Any]], weapon_col: str, perk_cols: List[str], note_col: Optional[str], tier_col: Optional[str], rank_col: Optional[str]) -> List[Dict[str, Any]]:
    expanded = []
    for row_idx, row in enumerate(rows, start=1):
        weapon = clean_text(row.get(weapon_col))
        if not weapon:
            continue
        # 类目行通常没有 perk。
        perk_options_by_col: List[List[str]] = []
        raw_perk_columns: Dict[str, str] = {}
        parsed_perk_columns: Dict[str, List[str]] = {}
        for col in perk_cols:
            raw_value = clean_text(row.get(col))
            opts = split_options(row.get(col))
            if raw_value:
                raw_perk_columns[col] = raw_value
            if opts:
                parsed_perk_columns[col] = opts
                perk_options_by_col.append(opts)
        if not perk_options_by_col:
            continue
        for combo in itertools.product(*perk_options_by_col):
            expanded.append({
                "source_row": row_idx,
                "weapon": weapon,
                "perks": list(combo),
                "notes": make_note(row, note_col, tier_col, rank_col),
                "raw_perk_columns": raw_perk_columns,
                "parsed_perk_columns": parsed_perk_columns,
            })
    return expanded


def write_extracted_report(path: Path, expanded: List[Dict[str, Any]]) -> None:
    """输出从原始推荐表识别并展开后的内容，先于 hash 匹配用于人工检查。"""
    rows = []
    for rec in expanded:
        rows.append({
            "source_row": rec.get("source_row", ""),
            "weapon": rec.get("weapon", ""),
            "expanded_perks": " / ".join(rec.get("perks", [])),
            "notes": rec.get("notes", ""),
            "raw_perk_columns_json": json.dumps(rec.get("raw_perk_columns", {}), ensure_ascii=False),
            "parsed_perk_columns_json": json.dumps(rec.get("parsed_perk_columns", {}), ensure_ascii=False),
        })
    csv_write(
        path,
        rows,
        [
            "source_row",
            "weapon",
            "expanded_perks",
            "notes",
            "raw_perk_columns_json",
            "parsed_perk_columns_json",
        ],
    )


def csv_write(path: Path, rows: List[Dict[str, Any]], fields: Sequence[str]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(fields))
        writer.writeheader()
        for r in rows:
            writer.writerow({k: r.get(k, "") for k in fields})


def sanitize_comment(text: Any) -> str:
    """清理写入 DIM 注释行的文本，避免换行破坏 block notes。"""
    return clean_text(text).replace("\r", " ").replace("\n", " ").strip()


def build_group_header(weapon: str, notes: str) -> List[str]:
    """生成接近 Little Light / DIM 官方工具的分组注释样式。"""
    header = [f"// {sanitize_comment(weapon)} - recommended"]
    note_text = sanitize_comment(notes)
    if note_text:
        header.append(f"//notes: {note_text}")
    return header


def build_wishlist(index: ManifestIndex, expanded: List[Dict[str, Any]]) -> Tuple[List[str], List[Dict[str, Any]]]:
    lines = [
        "title:Converted Chinese Weapon Wishlist",
        "description:Generated by dim_wishlist_builder.py from a Chinese recommendation table.",
        "",
    ]
    unresolved: List[Dict[str, Any]] = []
    seen_lines = set()

    current_group_key: Optional[Tuple[int, str, str]] = None
    group_has_output = False

    for rec in expanded:
        row_id = rec["source_row"]
        weapon = rec["weapon"]
        perks = rec["perks"]
        notes = rec.get("notes", "")

        weapon_candidates = index.find_weapons(weapon)
        if not weapon_candidates:
            unresolved.append({
                "source_row": row_id,
                "type": "weapon",
                "name": weapon,
                "reason": "weapon_not_found",
                "perks": " / ".join(perks),
            })
            continue

        perk_hashes: List[int] = []
        bad = False
        for perk in perks:
            cands = index.find_perk_hashes(perk)
            if not cands:
                unresolved.append({
                    "source_row": row_id,
                    "type": "perk",
                    "name": perk,
                    "reason": "perk_not_found",
                    "weapon": weapon,
                    "perks": " / ".join(perks),
                })
                bad = True
                break
            # 普通版优先后只取第一个，避免同一 perk 生成过多重复规则。
            perk_hashes.append(cands[0].hash)
        if bad:
            continue

        group_key = (int(row_id), sanitize_comment(weapon), sanitize_comment(notes))
        group_lines: List[str] = []
        perk_part = ",".join(str(x) for x in perk_hashes)
        for wc in weapon_candidates:
            # notes 使用上方 //notes: block note，不再写在每条规则末尾。
            line = f"dimwishlist:item={wc.hash}&perks={perk_part}"
            if line not in seen_lines:
                group_lines.append(line)
                seen_lines.add(line)

        if not group_lines:
            continue

        if group_key != current_group_key:
            if group_has_output:
                lines.append("")
            lines.extend(build_group_header(weapon, notes))
            current_group_key = group_key
            group_has_output = True

        lines.extend(group_lines)

    return lines, unresolved


def main(argv: Optional[Sequence[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="Convert Chinese weapon recommendations to DIM wishlist txt.")
    ap.add_argument("--input", required=True, help="中文推荐表，支持 .xlsx/.csv")
    ap.add_argument("--manifest", required=True, help="Bungie manifest .content 文件；可传压缩包或解压后的 sqlite")
    ap.add_argument("--sheet", default=None, help="Excel 工作表名；默认读取第一个工作表")
    ap.add_argument("--out", default="dim_wishlist_resolved.txt", help="输出 DIM wishlist txt")
    ap.add_argument("--unresolved", default="dim_wishlist_unresolved.csv", help="未匹配 CSV")
    ap.add_argument("--extracted", default="dim_wishlist_extracted.csv", help="从原始推荐表识别并展开后的检查 CSV")
    ap.add_argument("--cache-dir", default=".manifest_cache", help="manifest 解压缓存目录")
    args = ap.parse_args(argv)

    input_path = Path(args.input).expanduser().resolve()
    manifest_path = Path(args.manifest).expanduser().resolve()
    cache_dir = Path(args.cache_dir).expanduser().resolve()

    print(f"[1/5] 准备 manifest: {manifest_path}")
    sqlite_path = ensure_sqlite_manifest(manifest_path, cache_dir)
    print(f"      SQLite: {sqlite_path}")

    print("[2/5] 读取 DestinyInventoryItemDefinition ...")
    items = load_inventory_items(sqlite_path)
    index = ManifestIndex(items)
    print(f"      已读取 {len(items)} 个 InventoryItem")

    print(f"[3/5] 读取推荐表: {input_path}")
    headers, rows = read_table(input_path, sheet=args.sheet)
    weapon_col, perk_cols, note_col, tier_col, rank_col = detect_columns(headers)
    print(f"      武器列: {weapon_col}")
    print(f"      Perk列: {', '.join(perk_cols)}")

    expanded = expand_rows(rows, weapon_col, perk_cols, note_col, tier_col, rank_col)
    print(f"      展开后 roll 组合: {len(expanded)}")
    write_extracted_report(Path(args.extracted), expanded)
    print(f"      识别检查文件: {Path(args.extracted).resolve()}")

    print("[4/5] 匹配 hash 并生成 DIM wishlist ...")
    lines, unresolved = build_wishlist(index, expanded)
    Path(args.out).write_text("\n".join(lines) + "\n", encoding="utf-8")
    csv_write(Path(args.unresolved), unresolved, ["source_row", "type", "name", "reason", "weapon", "perks"])

    resolved_count = sum(1 for line in lines if line.startswith("dimwishlist:"))
    print("[5/5] 完成")
    print(f"      resolved lines: {resolved_count}")
    print(f"      unresolved rows: {len(unresolved)}")
    print(f"      wishlist: {Path(args.out).resolve()}")
    print(f"      unresolved: {Path(args.unresolved).resolve()}")
    print(f"      extracted: {Path(args.extracted).resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

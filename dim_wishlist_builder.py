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

# =========================
# 固定配置区：只改这里
# =========================
# 输入推荐表，支持 .csv / .xlsx。
INPUT_PATH = r"./命运2-凯旋丰碑全种类武器推荐-Sheet1.csv"

# Bungie manifest 文件。可以是刚下载的外层 .content 压缩包，也可以是解压后的 SQLite。
MANIFEST_PATH = r"./world_sql_content_22b6eb96bbcaa631746b584b52bcc2a6.content"

# Excel 工作表名；CSV 无效。None 表示读取第一个 sheet。
SHEET_NAME = None

# 输出目录。建议单独放一个文件夹
OUTPUT_DIR = r"./outputs"

# 输出文件名。
OUT_WISHLIST = "dim_wishlist_resolved.txt"
OUT_UNRESOLVED = "dim_wishlist_unresolved.csv"
OUT_EXTRACTED = "dim_wishlist_extracted.csv"
OUT_RESOLVED_AUDIT = "dim_wishlist_resolved_audit.csv"

# manifest 解压缓存目录。
CACHE_DIR = "./.manifest_cache"

# 武器同名候选处理策略：
# - "single": 同名查到多个武器 hash 时，只使用一个。
# - "all": 同名查到多个武器 hash 时，全部写入，并按每个 item hash 独立分组输出。推荐默认。
WEAPON_VERSION_MODE = "all"

# 武器 hash 手动覆盖。用于 manifest 查出多个同名候选时指定准确版本。
# 写法示例：WEAPON_HASH_OVERRIDES = {"Yeartide Apex": [3293207827], "毫不迟疑": [123456789]}
WEAPON_HASH_OVERRIDES: Dict[str, List[int]] = {}

# 输出所有武器名匹配到的候选 hash，方便检查为什么组合数翻倍。
OUT_WEAPON_CANDIDATES = "dim_wishlist_weapon_candidates.csv"
OUT_PERK_CANDIDATES = "dim_wishlist_perk_candidates.csv"
OUT_PERK_CANDIDATES = "dim_wishlist_perk_candidates.csv"


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
    """SQLite id / raw manifest hash -> DIM 使用的 uint32 hash。

    Bungie manifest 的 SQLite 表 id 常用 signed int32 存储；
    DIM wishlist 必须使用 unsigned uint32。
    例如 SQLite id = -1001759469，对应 DIM hash = 3293207827。
    """
    x = int(value)
    if x < 0:
        x += UINT32
    return x


def to_sql_id(value: Any) -> int:
    """DIM 使用的 uint32 hash -> SQLite 查询用 signed int32 id。

    手动写 SQL 时，如果 hash > 2147483647，需要减去 4294967296。
    例如 DIM hash = 3293207827，对应 SQLite id = -1001759469。
    """
    x = int(value)
    if x >= 2 ** 31:
        x -= UINT32
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
    sql_id: int
    name: str
    item_type: Optional[int]
    item_type_display: str
    item_type_and_tier_display: str
    tier_type_name: str
    plug_category_identifier: str
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
        plug_obj = obj.get("plug") or {}
        # SQLite 的 id 可能是 signed int32；JSON 内部 hash 一般是 DIM 使用的 unsigned hash。
        # 优先使用 JSON hash，避免手动查 SQL 时的 signed/unsigned 混淆。
        dim_hash = obj.get("hash")
        if dim_hash is None:
            dim_hash = to_dim_hash(raw_id)
        items.append(
            InvItem(
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
        """按 DIM wishlist 需要的 InventoryItem plug hash 解析 perk。

        关键规则：
        - 只接受精确同名；
        - 优先 itemType == 19 且存在 plug；
        - 普通“特性”优先；
        - 排除“强化特征”、徽标等同名非目标条目。
        """
        key = norm_name(name)
        exact = self.by_name.get(key, [])
        if not exact:
            return []

        # DIM wishlist 的 perk 应使用 DestinyInventoryItemDefinition 里的 plug item hash。
        plug_traits = [x for x in exact if x.item_type == 19 and x.has_plug]
        if plug_traits:
            return self._rank_perk_candidates(plug_traits)

        # 极少数定义可能缺 plug 字段，仍限制 itemType==19。
        type19 = [x for x in exact if x.item_type == 19]
        if type19:
            return self._rank_perk_candidates(type19)

        # 不再回退到徽标、收藏品等同名非 perk 条目。
        return []

    @staticmethod
    def _is_enhanced_perk(c: InvItem) -> bool:
        text = " ".join([
            c.name,
            c.item_type_display,
            c.item_type_and_tier_display,
            c.tier_type_name,
        ]).lower()
        return ("强化" in text) or ("enhanced" in text)

    @staticmethod
    def _is_normal_trait(c: InvItem) -> bool:
        # 中文 manifest：itemTypeDisplayName="特性", tierTypeName="普通"。
        # 英文 manifest：itemTypeDisplayName 通常为 "Trait"。
        t = c.item_type_display.lower()
        tier = c.tier_type_name.lower()
        tt = c.item_type_and_tier_display.lower()
        is_trait_name = (c.item_type_display in {"特性", "Trait"}) or (" trait" in tt) or tt.endswith("trait")
        is_normal_tier = (c.tier_type_name in {"普通", "Common"}) or ("普通" in tt) or ("common" in tt)
        return is_trait_name and is_normal_tier and not ManifestIndex._is_enhanced_perk(c)

    @staticmethod
    def _perk_rank(c: InvItem) -> Tuple[int, int, int, int, int]:
        """候选排序：普通特性 > 非强化 plug > 强化 plug > 其他 itemType19 > 其他。"""
        if ManifestIndex._is_normal_trait(c):
            return (0, 0, 0, 0, c.hash)
        if c.item_type == 19 and c.has_plug and not ManifestIndex._is_enhanced_perk(c):
            return (1, 0, 0, 0, c.hash)
        if c.item_type == 19 and c.has_plug:
            return (2, 0, 0, 0, c.hash)
        if c.item_type == 19:
            return (3, 0, 0, 0, c.hash)
        return (9, 9, 9, 9, c.hash)

    @staticmethod
    def _rank_perk_candidates(cands: List[InvItem]) -> List[InvItem]:
        """同名候选排序：普通特性 > 非强化 plug > 其他 itemType19 plug。"""
        if not cands:
            return []

        seen = set()
        out: List[InvItem] = []
        for c in sorted(cands, key=ManifestIndex._perk_rank):
            if c.hash in seen:
                continue
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


def canonical_perk_slot(col: str, index: int) -> str:
    """把原始列名固定映射为 DIM 检查用的 slot 名称。

    你的表头是 Perk, Perk, Perk 1, Perk 2。第二个重复 Perk
    会被内部重命名为 Perk__2，但语义仍然是第二列。
    """
    base = col.split("__", 1)[0].strip()
    if index == 0:
        return "slot_1_barrel"
    if index == 1:
        return "slot_2_magazine"
    if base.lower() in {"perk 1", "trait 1", "特性1", "特性 1"}:
        return "slot_3_trait"
    if base.lower() in {"perk 2", "trait 2", "特性2", "特性 2"}:
        return "slot_4_trait"
    return f"slot_{index + 1}"


def expand_rows(rows: List[Dict[str, Any]], weapon_col: str, perk_cols: List[str], note_col: Optional[str], tier_col: Optional[str], rank_col: Optional[str]) -> List[Dict[str, Any]]:
    expanded = []
    for row_idx, row in enumerate(rows, start=1):
        weapon = clean_text(row.get(weapon_col))
        if not weapon:
            continue
        # 类目行通常没有 perk。
        perk_options_by_col: List[List[str]] = []
        slot_names_by_col: List[str] = []
        raw_perk_columns: Dict[str, str] = {}
        parsed_perk_columns: Dict[str, List[str]] = {}
        for col_index, col in enumerate(perk_cols):
            raw_value = clean_text(row.get(col))
            opts = split_options(row.get(col))
            slot_name = canonical_perk_slot(col, col_index)
            if raw_value:
                raw_perk_columns[slot_name] = raw_value
            if opts:
                parsed_perk_columns[slot_name] = opts
                perk_options_by_col.append(opts)
                slot_names_by_col.append(slot_name)
        if not perk_options_by_col:
            continue
        for combo in itertools.product(*perk_options_by_col):
            perk_slots = {slot: perk for slot, perk in zip(slot_names_by_col, combo)}
            expanded.append({
                "source_row": row_idx,
                "weapon": weapon,
                "perks": [perk_slots[slot] for slot in slot_names_by_col],
                "perk_slots": perk_slots,
                "slot_order": list(slot_names_by_col),
                "notes": make_note(row, note_col, tier_col, rank_col),
                "raw_perk_columns": raw_perk_columns,
                "parsed_perk_columns": parsed_perk_columns,
            })
    return expanded


def write_extracted_report(path: Path, expanded: List[Dict[str, Any]]) -> None:
    """输出从原始推荐表识别并展开后的内容，先于 hash 匹配用于人工检查。"""
    rows = []
    for rec in expanded:
        perk_slots = rec.get("perk_slots", {}) or {}
        rows.append({
            "source_row": rec.get("source_row", ""),
            "weapon": rec.get("weapon", ""),
            "slot_1_barrel": perk_slots.get("slot_1_barrel", ""),
            "slot_2_magazine": perk_slots.get("slot_2_magazine", ""),
            "slot_3_trait": perk_slots.get("slot_3_trait", ""),
            "slot_4_trait": perk_slots.get("slot_4_trait", ""),
            "expanded_perks_in_dim_order": " / ".join(rec.get("perks", [])),
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
            "slot_1_barrel",
            "slot_2_magazine",
            "slot_3_trait",
            "slot_4_trait",
            "expanded_perks_in_dim_order",
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



def describe_inv_item(c: InvItem) -> str:
    return f"{c.hash}:{c.name}:{c.item_type_display}:{c.tier_type_name}"


def write_perk_candidate_report(path: Path, index: ManifestIndex, expanded: List[Dict[str, Any]]) -> None:
    """输出每个请求 perk 的所有精确同名候选，方便检查是否误选强化/徽标。"""
    seen = set()
    rows: List[Dict[str, Any]] = []
    for rec in expanded:
        weapon = rec.get("weapon", "")
        row_id = rec.get("source_row", "")
        perk_slots = rec.get("perk_slots", {}) or {}
        slot_order = rec.get("slot_order", []) or []
        for slot in slot_order:
            perk = perk_slots.get(slot, "")
            key = (weapon, slot, perk)
            if not perk or key in seen:
                continue
            seen.add(key)
            exact = index.by_name.get(norm_name(perk), [])
            selected = index.find_perk_hashes(perk)
            rows.append({
                "source_row": row_id,
                "weapon": weapon,
                "slot": slot,
                "perk": perk,
                "selected_hash": selected[0].hash if selected else "",
                "selected_sqlite_id": selected[0].sql_id if selected else "",
                "selected_name": selected[0].name if selected else "",
                "selected_type_display": selected[0].item_type_display if selected else "",
                "selected_tier": selected[0].tier_type_name if selected else "",
                "candidate_count": len(exact),
                "candidate_hashes": ";".join(str(c.hash) for c in exact),
                "candidate_sqlite_ids": ";".join(str(c.sql_id) for c in exact),
                "candidate_types": ";".join(f"{c.item_type_display}/{c.tier_type_name}/itemType={c.item_type}/plug={c.has_plug}" for c in exact),
                "candidate_names": ";".join(c.name for c in exact),
            })
    csv_write(path, rows, [
        "source_row", "weapon", "slot", "perk",
        "selected_hash", "selected_sqlite_id", "selected_name", "selected_type_display", "selected_tier",
        "candidate_count", "candidate_hashes", "candidate_sqlite_ids", "candidate_types", "candidate_names",
    ])

def build_group_header(weapon: str, notes: str) -> List[str]:
    """生成接近 Little Light / DIM 官方工具的分组注释样式。"""
    header = [f"// {sanitize_comment(weapon)} - recommended"]
    note_text = sanitize_comment(notes)
    if note_text:
        header.append(f"//notes: {note_text}")
    return header



def select_weapon_candidates(weapon: str, candidates: List[InvItem]) -> List[InvItem]:
    """根据固定配置决定同名武器候选的写入方式。

    组合数量应首先由四个 perk 槽位决定。例如 2×2×3×3=36。
    如果同一武器名在 manifest 中匹配到 2 个 item hash，而又全部写入，最终会变成 36×2=72。
    """
    if not candidates:
        return []

    # 手动覆盖优先。支持中文名或英文名，只要表里 weapon 字段能 norm 后匹配。
    wkey = norm_name(weapon)
    for k, hashes in WEAPON_HASH_OVERRIDES.items():
        if norm_name(k) == wkey:
            allowed = {int(h) for h in hashes}
            picked = [c for c in candidates if c.hash in allowed]
            if picked:
                return picked
            # 覆盖写错时不要静默回退，避免误生成。
            return []

    mode = str(WEAPON_VERSION_MODE).strip().lower()
    if mode == "all":
        return candidates
    if mode == "single":
        # 稳定选择第一个候选。候选排序已在 find_weapons 中完成。
        return candidates[:1]
    raise RuntimeError(f"未知 WEAPON_VERSION_MODE: {WEAPON_VERSION_MODE!r}，只能是 'single' 或 'all'")


def write_weapon_candidate_report(path: Path, index: ManifestIndex, expanded: List[Dict[str, Any]]) -> None:
    """输出每个武器名在 manifest 中匹配到的 item hash 候选。"""
    seen_names = []
    seen_key = set()
    for rec in expanded:
        w = rec.get("weapon", "")
        k = norm_name(w)
        if w and k not in seen_key:
            seen_names.append(w)
            seen_key.add(k)

    rows = []
    for weapon in seen_names:
        candidates = index.find_weapons(weapon)
        selected = select_weapon_candidates(weapon, candidates)
        cand_hashes = [c.hash for c in candidates]
        selected_hashes = [c.hash for c in selected]
        for c in candidates:
            rows.append({
                "weapon": weapon,
                "candidate_count": len(candidates),
                "selected_count": len(selected),
                "selected": "yes" if c.hash in selected_hashes else "no",
                "hash": c.hash,
                "sqlite_id": c.sql_id,
                "manifest_name": c.name,
                "item_type_display": c.item_type_display,
                "item_type_and_tier_display": c.item_type_and_tier_display,
                "tier_type_name": c.tier_type_name,
                "item_type": c.item_type,
                "has_plug": c.has_plug,
                "all_candidate_hashes": ";".join(str(x) for x in cand_hashes),
                "selected_hashes": ";".join(str(x) for x in selected_hashes),
            })
        if not candidates:
            rows.append({
                "weapon": weapon,
                "candidate_count": 0,
                "selected_count": 0,
                "selected": "no",
                "hash": "",
                "sqlite_id": "",
                "manifest_name": "",
                "item_type_display": "",
                "item_type_and_tier_display": "",
                "tier_type_name": "",
                "item_type": "",
                "has_plug": "",
                "all_candidate_hashes": "",
                "selected_hashes": "",
            })
    csv_write(path, rows, [
        "weapon", "candidate_count", "selected_count", "selected", "hash", "sqlite_id", "manifest_name",
        "item_type_display", "item_type_and_tier_display", "tier_type_name", "item_type", "has_plug",
        "all_candidate_hashes", "selected_hashes",
    ])


def write_perk_candidate_report(path: Path, index: ManifestIndex, expanded: List[Dict[str, Any]]) -> None:
    """输出每个 perk 名称的候选与最终选择，专门排查同名/强化/徽标误匹配。"""
    seen = []
    seen_key = set()
    for rec in expanded:
        for slot, perk in (rec.get("perk_slots", {}) or {}).items():
            k = (slot, norm_name(perk))
            if perk and k not in seen_key:
                seen.append((slot, perk))
                seen_key.add(k)

    rows = []
    for slot, perk in seen:
        exact = index.by_name.get(norm_name(perk), [])
        selected = index.find_perk_hashes(perk)
        selected_hash = selected[0].hash if selected else ""
        for c in sorted(exact, key=lambda x: ManifestIndex._perk_rank(x) if x.item_type == 19 else (9, 9, 9, 9, x.hash)):
            rows.append({
                "slot": slot,
                "requested_perk": perk,
                "selected": "yes" if c.hash == selected_hash else "no",
                "hash": c.hash,
                "sqlite_id": c.sql_id,
                "manifest_name": c.name,
                "item_type": c.item_type,
                "item_type_display": c.item_type_display,
                "item_type_and_tier_display": c.item_type_and_tier_display,
                "tier_type_name": c.tier_type_name,
                "has_plug": c.has_plug,
                "plug_category_identifier": c.plug_category_identifier,
                "candidate_count_same_name": len(exact),
                "selected_hash": selected_hash,
            })
        if not exact:
            rows.append({
                "slot": slot,
                "requested_perk": perk,
                "selected": "no",
                "hash": "",
                "sqlite_id": "",
                "manifest_name": "",
                "item_type": "",
                "item_type_display": "",
                "item_type_and_tier_display": "",
                "tier_type_name": "",
                "has_plug": "",
                "plug_category_identifier": "",
                "candidate_count_same_name": 0,
                "selected_hash": "",
            })

    csv_write(path, rows, [
        "slot", "requested_perk", "selected", "hash", "sqlite_id", "manifest_name",
        "item_type", "item_type_display", "item_type_and_tier_display", "tier_type_name",
        "has_plug", "plug_category_identifier", "candidate_count_same_name", "selected_hash",
    ])

def build_wishlist(index: ManifestIndex, expanded: List[Dict[str, Any]]) -> Tuple[List[str], List[Dict[str, Any]], List[Dict[str, Any]]]:
    """生成 DIM wishlist。

    关键点：
    - perk 组合按固定槽位展开：slot_1 × slot_2 × slot_3 × slot_4。
    - 如果同名武器匹配到多个 item hash，每个 item hash 作为一个独立分组输出。
      例如 2 个版本、每个版本 36 个组合，应输出两个分组，各 36 条；
      不应在同一个分组里 2965... / 3293... 交替出现。
    """
    lines = [
        "title:Converted Chinese Weapon Wishlist",
        "description:Generated by dim_wishlist_builder.py from a Chinese recommendation table.",
        "",
    ]
    unresolved: List[Dict[str, Any]] = []
    resolved_audit: List[Dict[str, Any]] = []

    # groups 保持插入顺序：先按原始表行，再按武器候选 hash 分组。
    # key = (source_row, weapon, notes, weapon_hash, manifest_name)
    groups: Dict[Tuple[int, str, str, int, str], List[str]] = {}
    seen_lines_by_group: Dict[Tuple[int, str, str, int, str], set] = {}

    for rec in expanded:
        row_id = rec["source_row"]
        weapon = rec["weapon"]
        perks = rec["perks"]
        notes = rec.get("notes", "")

        all_weapon_candidates = index.find_weapons(weapon)
        weapon_candidates = select_weapon_candidates(weapon, all_weapon_candidates)
        if not all_weapon_candidates:
            unresolved.append({
                "source_row": row_id,
                "type": "weapon",
                "name": weapon,
                "reason": "weapon_not_found",
                "perks": " / ".join(perks),
            })
            continue
        if not weapon_candidates:
            unresolved.append({
                "source_row": row_id,
                "type": "weapon",
                "name": weapon,
                "reason": "weapon_override_or_selection_empty",
                "perks": " / ".join(perks),
            })
            continue

        perk_hashes: List[int] = []
        selected_perk_types: List[str] = []
        selected_perk_sqlite_ids: List[int] = []
        bad = False
        for perk in perks:
            cands = index.find_perk_hashes(perk)
            if not cands:
                same_name = index.by_name.get(norm_name(perk), [])
                unresolved.append({
                    "source_row": row_id,
                    "type": "perk",
                    "name": perk,
                    "reason": "perk_not_found_or_no_valid_inventory_plug",
                    "weapon": weapon,
                    "perks": " / ".join(perks),
                    "same_name_candidate_hashes": ";".join(str(x.hash) for x in same_name),
                    "same_name_candidate_types": ";".join(f"{x.hash}:{x.item_type_display}/{x.item_type_and_tier_display}" for x in same_name),
                })
                bad = True
                break
            # 普通版优先后只取第一个，避免同一 perk 生成过多重复规则。
            perk_hashes.append(cands[0].hash)
        if bad:
            continue

        perk_slots = rec.get("perk_slots", {}) or {}
        slot_order = rec.get("slot_order", []) or []
        slot_hashes = {slot: perk_hashes[i] if i < len(perk_hashes) else "" for i, slot in enumerate(slot_order)}
        slot_sqlite_ids = {slot: selected_perk_sqlite_ids[i] if i < len(selected_perk_sqlite_ids) else "" for i, slot in enumerate(slot_order)}
        slot_types = {slot: selected_perk_types[i] if i < len(selected_perk_types) else "" for i, slot in enumerate(slot_order)}
        perk_part = ",".join(str(x) for x in perk_hashes)

        for wc in weapon_candidates:
            group_key = (
                int(row_id),
                sanitize_comment(weapon),
                sanitize_comment(notes),
                int(wc.hash),
                sanitize_comment(wc.name),
            )
            if group_key not in groups:
                groups[group_key] = []
                seen_lines_by_group[group_key] = set()

            line = f"dimwishlist:item={wc.hash}&perks={perk_part}"
            if line in seen_lines_by_group[group_key]:
                continue
            groups[group_key].append(line)
            seen_lines_by_group[group_key].add(line)
            resolved_audit.append({
                "source_row": row_id,
                "weapon": weapon,
                "weapon_hash": wc.hash,
                "weapon_sqlite_id": wc.sql_id,
                "manifest_name": wc.name,
                "slot_1_barrel": perk_slots.get("slot_1_barrel", ""),
                "slot_1_hash": slot_hashes.get("slot_1_barrel", ""),
                "slot_1_sqlite_id": slot_sqlite_ids.get("slot_1_barrel", ""),
                "slot_1_type": slot_types.get("slot_1_barrel", ""),
                "slot_2_magazine": perk_slots.get("slot_2_magazine", ""),
                "slot_2_hash": slot_hashes.get("slot_2_magazine", ""),
                "slot_2_sqlite_id": slot_sqlite_ids.get("slot_2_magazine", ""),
                "slot_2_type": slot_types.get("slot_2_magazine", ""),
                "slot_3_trait": perk_slots.get("slot_3_trait", ""),
                "slot_3_hash": slot_hashes.get("slot_3_trait", ""),
                "slot_3_sqlite_id": slot_sqlite_ids.get("slot_3_trait", ""),
                "slot_3_type": slot_types.get("slot_3_trait", ""),
                "slot_4_trait": perk_slots.get("slot_4_trait", ""),
                "slot_4_hash": slot_hashes.get("slot_4_trait", ""),
                "slot_4_sqlite_id": slot_sqlite_ids.get("slot_4_trait", ""),
                "slot_4_type": slot_types.get("slot_4_trait", ""),
                "dim_line": line,
            })

    wrote_any_group = False
    for group_key, group_lines in groups.items():
        if not group_lines:
            continue
        _row_id, weapon_name, notes, weapon_hash, manifest_name = group_key
        if wrote_any_group:
            lines.append("")

        # 多版本时必须分组。标题里加 hash，方便肉眼检查不同版本没有混在一起。
        header_weapon = weapon_name
        if manifest_name and norm_name(manifest_name) != norm_name(weapon_name):
            header_weapon = f"{weapon_name} / {manifest_name}"
        header_weapon = f"{header_weapon} [{weapon_hash}]"
        lines.extend(build_group_header(header_weapon, notes))
        lines.extend(group_lines)
        wrote_any_group = True

    return lines, unresolved, resolved_audit


def main() -> int:
    input_path = Path(INPUT_PATH).expanduser().resolve()
    manifest_path = Path(MANIFEST_PATH).expanduser().resolve()
    output_dir = Path(OUTPUT_DIR).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    out_path = output_dir / OUT_WISHLIST
    unresolved_path = output_dir / OUT_UNRESOLVED
    extracted_path = output_dir / OUT_EXTRACTED
    resolved_audit_path = output_dir / OUT_RESOLVED_AUDIT
    weapon_candidates_path = output_dir / OUT_WEAPON_CANDIDATES
    perk_candidates_path = output_dir / OUT_PERK_CANDIDATES
    perk_candidates_path = output_dir / OUT_PERK_CANDIDATES
    cache_dir = Path(CACHE_DIR).expanduser().resolve()

    print("[CONFIG]")
    print(f"      input: {input_path}")
    print(f"      manifest: {manifest_path}")
    print(f"      output_dir: {output_dir}")
    print(f"      weapon_version_mode: {WEAPON_VERSION_MODE}")
    print("")

    print(f"[1/5] 准备 manifest: {manifest_path}")
    sqlite_path = ensure_sqlite_manifest(manifest_path, cache_dir)
    print(f"      SQLite: {sqlite_path}")

    print("[2/5] 读取 DestinyInventoryItemDefinition ...")
    items = load_inventory_items(sqlite_path)
    index = ManifestIndex(items)
    print(f"      已读取 {len(items)} 个 InventoryItem")

    print(f"[3/5] 读取推荐表: {input_path}")
    headers, rows = read_table(input_path, sheet=SHEET_NAME)
    weapon_col, perk_cols, note_col, tier_col, rank_col = detect_columns(headers)
    print(f"      武器列: {weapon_col}")
    print(f"      Perk列: {', '.join(perk_cols)}")

    expanded = expand_rows(rows, weapon_col, perk_cols, note_col, tier_col, rank_col)
    print(f"      展开后 roll 组合: {len(expanded)}")
    write_extracted_report(extracted_path, expanded)
    print(f"      识别检查文件: {extracted_path}")
    write_weapon_candidate_report(weapon_candidates_path, index, expanded)
    print(f"      武器候选检查文件: {weapon_candidates_path}")
    write_perk_candidate_report(perk_candidates_path, index, expanded)
    print(f"      Perk候选检查文件: {perk_candidates_path}")

    print("[4/5] 匹配 hash 并生成 DIM wishlist ...")
    lines, unresolved, resolved_audit = build_wishlist(index, expanded)
    out_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    csv_write(unresolved_path, unresolved, [
        "source_row", "type", "name", "reason", "weapon", "perks",
        "same_name_candidate_hashes", "same_name_candidate_types",
    ])
    csv_write(resolved_audit_path, resolved_audit, [
        "source_row", "weapon", "weapon_hash", "weapon_sqlite_id",
        "slot_1_barrel", "slot_1_hash", "slot_1_sqlite_id", "slot_1_type",
        "slot_2_magazine", "slot_2_hash", "slot_2_sqlite_id", "slot_2_type",
        "slot_3_trait", "slot_3_hash", "slot_3_sqlite_id", "slot_3_type",
        "slot_4_trait", "slot_4_hash", "slot_4_sqlite_id", "slot_4_type",
        "dim_line",
    ])

    resolved_count = sum(1 for line in lines if line.startswith("dimwishlist:"))
    print("[5/5] 完成")
    print(f"      resolved lines: {resolved_count}")
    print(f"      unresolved rows: {len(unresolved)}")
    print(f"      wishlist: {out_path}")
    print(f"      unresolved: {unresolved_path}")
    print(f"      extracted: {extracted_path}")
    print(f"      resolved audit: {resolved_audit_path}")
    print(f"      weapon candidates: {weapon_candidates_path}")
    print(f"      perk candidates: {perk_candidates_path}")
    print(f"      perk candidates: {perk_candidates_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

"""Read XLSX drawing anchors and map embedded icons to recommendation slots."""

from __future__ import annotations

import hashlib
import re
import zipfile
from pathlib import Path, PurePosixPath
from typing import Any, Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

from .icon_config import (
    ICON_COLUMNS_PER_SLOT,
    NS_A,
    NS_DRAW,
    NS_REL_DOC,
    NS_SHEET,
    REL_PVE_TRAIT3_START,
    REL_PVE_TRAIT4_START,
    REL_PVP_TRAIT3_START,
    REL_PVP_TRAIT4_START,
    REL_SPECIAL,
    REL_WEAPON_TYPE,
    IconBuilderConfig,
)
from .icon_models import DrawingImage, IconContext
from .utils import clean_text


def col_letters_to_zero_based(letters: str) -> int:
    number = 0
    for character in letters:
        number = number * 26 + ord(character.upper()) - ord("A") + 1
    return number - 1


def zero_based_to_col_letters(index: int) -> str:
    result = ""
    number = index + 1
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def read_shared_strings(archive: zipfile.ZipFile) -> List[str]:
    name = "xl/sharedStrings.xml"
    if name not in archive.namelist():
        return []
    root = ET.fromstring(archive.read(name))
    return [
        "".join(node.text or "" for node in item.iter(f"{{{NS_SHEET}}}t"))
        for item in root.findall(f"{{{NS_SHEET}}}si")
    ]


def read_sheet_cells(
    archive: zipfile.ZipFile,
    sheet_path: str = "xl/worksheets/sheet1.xml",
) -> Dict[Tuple[int, int], str]:
    shared = read_shared_strings(archive)
    root = ET.fromstring(archive.read(sheet_path))
    cells = {}
    for cell in root.findall(f".//{{{NS_SHEET}}}c"):
        match = re.fullmatch(r"([A-Z]+)(\d+)", cell.attrib.get("r", ""))
        if not match:
            continue
        col = col_letters_to_zero_based(match.group(1))
        row = int(match.group(2)) - 1
        cell_type = cell.attrib.get("t", "")
        if cell_type == "inlineStr":
            value = "".join(node.text or "" for node in cell.iter(f"{{{NS_SHEET}}}t"))
        else:
            value_node = cell.find(f"{{{NS_SHEET}}}v")
            raw = "" if value_node is None or value_node.text is None else value_node.text
            if cell_type == "s" and raw:
                try:
                    value = shared[int(raw)]
                except (ValueError, IndexError):
                    value = raw
            else:
                value = raw
        if value != "":
            cells[row, col] = value
    return cells


def resolve_zip_target(base_file: str, target: str) -> str:
    parts: List[str] = []
    for part in PurePosixPath(PurePosixPath(base_file).parent, target).parts:
        if part == "..":
            if parts:
                parts.pop()
        elif part not in (".", ""):
            parts.append(part)
    return "/".join(parts)


def read_drawing_images(archive: zipfile.ZipFile) -> List[DrawingImage]:
    drawing_path = "xl/drawings/drawing1.xml"
    relationship_path = "xl/drawings/_rels/drawing1.xml.rels"
    if drawing_path not in archive.namelist() or relationship_path not in archive.namelist():
        raise RuntimeError("XLSX 中没有找到 drawing1.xml，无法提取内嵌图标。")

    relationships = {
        node.attrib.get("Id", ""): node.attrib.get("Target", "")
        for node in list(ET.fromstring(archive.read(relationship_path)))
        if node.attrib.get("Id") and node.attrib.get("Target")
    }
    root = ET.fromstring(archive.read(drawing_path))
    result = []
    for anchor in list(root):
        start = anchor.find(f"{{{NS_DRAW}}}from")
        if start is None:
            continue
        try:
            row = int(start.findtext(f"{{{NS_DRAW}}}row", default="-1"))
            col = int(start.findtext(f"{{{NS_DRAW}}}col", default="-1"))
        except ValueError:
            continue
        blip = anchor.find(f".//{{{NS_A}}}blip")
        if blip is None:
            continue
        relationship_id = blip.attrib.get(f"{{{NS_REL_DOC}}}embed", "")
        target = relationships.get(relationship_id, "")
        media_path = resolve_zip_target(drawing_path, target) if target else ""
        if not media_path or media_path not in archive.namelist():
            continue
        data = archive.read(media_path)
        result.append(DrawingImage(
            row=row,
            col=col,
            target=target,
            media_path=media_path,
            data=data,
            sha256=sha256_bytes(data),
            extension=Path(media_path).suffix.lower() or ".bin",
        ))
    return result


def detect_sections(cells: Dict[Tuple[int, int], str]) -> List[int]:
    starts = sorted({
        col for (row, col), value in cells.items()
        if row == 0 and clean_text(value) == "武器名称"
    })
    if not starts:
        raise RuntimeError("第一行没有找到“武器名称”表头，无法识别横向武器区块。")
    return starts


def classify_icon_column(section_start: int, col: int) -> Optional[Tuple[str, str, int]]:
    offset = col - section_start
    for start, usage, slot in (
        (REL_PVE_TRAIT3_START, "pve", "trait_3"),
        (REL_PVE_TRAIT4_START, "pve", "trait_4"),
        (REL_PVP_TRAIT3_START, "pvp", "trait_3"),
        (REL_PVP_TRAIT4_START, "pvp", "trait_4"),
    ):
        if start <= offset < start + ICON_COLUMNS_PER_SLOT:
            return usage, slot, offset - start + 1
    return None


def export_unique_icon(icon_dir: Path, image: DrawingImage) -> str:
    icon_dir.mkdir(parents=True, exist_ok=True)
    filename = f"{image.sha256[:24]}{image.extension}"
    path = icon_dir / filename
    if not path.exists():
        path.write_bytes(image.data)
    return filename


def extract_icon_contexts(
    xlsx_path: Path,
    output_dir: Path,
    config: IconBuilderConfig,
) -> Tuple[List[IconContext], Dict[str, Any]]:
    if not xlsx_path.exists():
        raise FileNotFoundError(f"找不到输入文件：{xlsx_path}")
    icon_dir = output_dir / config.extracted_icon_dirname
    with zipfile.ZipFile(xlsx_path) as archive:
        cells = read_sheet_cells(archive)
        drawings = read_drawing_images(archive)
        sections = detect_sections(cells)

    contexts = []
    ignored = missing_weapon = 0
    for image in drawings:
        section_candidates = [start for start in sections if start <= image.col < start + 20]
        if not section_candidates:
            ignored += 1
            continue
        section_start = max(section_candidates)
        classified = classify_icon_column(section_start, image.col)
        if classified is None:
            ignored += 1
            continue
        usage, slot, slot_position = classified
        weapon_name = clean_text(cells.get((image.row, section_start), ""))
        weapon_type = clean_text(cells.get((image.row, section_start + REL_WEAPON_TYPE), ""))
        special_note = clean_text(cells.get((image.row, section_start + REL_SPECIAL), ""))
        if not weapon_name:
            for previous_row in range(image.row - 1, max(-1, image.row - 4), -1):
                weapon_name = clean_text(cells.get((previous_row, section_start), ""))
                if weapon_name:
                    weapon_type = weapon_type or clean_text(
                        cells.get((previous_row, section_start + REL_WEAPON_TYPE), "")
                    )
                    special_note = special_note or clean_text(
                        cells.get((previous_row, section_start + REL_SPECIAL), "")
                    )
                    break
        if not weapon_name:
            missing_weapon += 1
            continue
        exported = (
            export_unique_icon(icon_dir, image)
            if config.write_diagnostics else f"{image.sha256[:24]}{image.extension}"
        )
        contexts.append(IconContext(
            section_index=sections.index(section_start) + 1,
            excel_row=image.row + 1,
            weapon_name=weapon_name,
            weapon_type=weapon_type,
            usage=usage,
            slot=slot,
            slot_position=slot_position,
            source_col=image.col,
            source_cell=f"{zero_based_to_col_letters(image.col)}{image.row + 1}",
            special_note=special_note,
            media_path=image.media_path,
            icon_sha256=image.sha256,
            icon_extension=image.extension,
            icon_bytes=image.data,
            exported_icon=str(Path(config.extracted_icon_dirname) / exported),
        ))

    contexts.sort(key=lambda item: (
        item.excel_row, item.section_index, item.usage, item.slot, item.slot_position
    ))
    stats = {
        "drawing_count": len(drawings),
        "unique_media_count": len({image.sha256 for image in drawings}),
        "section_count": len(sections),
        "perk_icon_position_count": len(contexts),
        "unique_perk_icon_count": len({item.icon_sha256 for item in contexts}),
        "ignored_drawing_count": ignored,
        "missing_weapon_count": missing_weapon,
        "section_start_columns": sections,
    }
    return contexts, stats

"""Text, hash, filesystem, and CSV helpers."""

from __future__ import annotations

import csv
import re
import tempfile
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence


UINT32 = 2**32
SPLIT_RE = re.compile(r"[\n\r/／、，,;；|｜]+")


def norm_name(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().replace("\u3000", " ").lower()
    return re.sub(r"[\s\-_'\"“”‘’·•:：()（）\[\]【】{}<>《》]+", "", text)


def clean_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def split_options(cell: Any) -> List[str]:
    text = clean_text(cell)
    if not text:
        return []
    parts = [part.strip(" \t-—*·•") for part in SPLIT_RE.split(text)]
    return [part for part in parts if part]


def to_dim_hash(value: Any) -> int:
    """Convert a signed SQLite id to DIM's unsigned uint32 hash."""
    number = int(value)
    return number + UINT32 if number < 0 else number


def to_sql_id(value: Any) -> int:
    """Convert DIM's uint32 hash to the signed id used by SQLite."""
    number = int(value)
    return number - UINT32 if number >= 2**31 else number


def ensure_sqlite_manifest(path: Path, workdir: Optional[Path] = None) -> Path:
    """Return a queryable SQLite path, extracting a zipped manifest if needed."""
    path = path.expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"manifest 不存在: {path}")
    if not zipfile.is_zipfile(path):
        return path

    out_dir = workdir or Path(tempfile.mkdtemp(prefix="bungie_manifest_"))
    out_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "r") as archive:
        names = [name for name in archive.namelist() if not name.endswith("/")]
        if not names:
            raise RuntimeError(f"zip manifest 为空: {path}")
        archive.extractall(out_dir)
    return (out_dir / names[0]).resolve()


def csv_write(path: Path, rows: List[Dict[str, Any]], fields: Sequence[str]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=list(fields))
        writer.writeheader()
        for row in rows:
            writer.writerow({field: row.get(field, "") for field in fields})


def sanitize_comment(text: Any) -> str:
    return clean_text(text).replace("\r", " ").replace("\n", " ").strip()

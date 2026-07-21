"""Unified dispatcher for text and icon recommendation inputs."""

from __future__ import annotations

import sys
from typing import Optional, Sequence


HELP = """usage: dim_wishlist_builder.py [text|icon] [options]

双输入 DIM Wishlist 生成器：
  text    中文武器名/perk文字表（默认，兼容原命令）
  icon    perk以嵌入图片表示的XLSX

examples:
  python3 dim_wishlist_builder.py
  python3 dim_wishlist_builder.py text --help
  python3 dim_wishlist_builder.py icon --run-mode extract_only
  python3 dim_wishlist_builder.py icon --help
"""


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if args and args[0] == "icon":
        from .icon_cli import main as icon_main
        return icon_main(args[1:])
    if args and args[0] == "text":
        from .cli import main as text_main
        return text_main(args[1:])
    if args in (["-h"], ["--help"]):
        print(HELP)
        return 0
    from .cli import main as text_main
    return text_main(args)

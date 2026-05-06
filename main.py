#!/usr/bin/env python3
"""
Office 字体标准化工具 - 入口

GUI 模式:  uv run main.py
CLI 模式:  uv run main.py file1.pptx [file2.docx ...]
"""

import sys
from pathlib import Path

from core import DEFAULT_FONTS, HANDLERS, ensure_fonts_installed, process_file


def run_cli(args: list[str]) -> None:
    fonts = DEFAULT_FONTS.copy()
    for arg in args:
        p = Path(arg)
        if not p.exists():
            print(f'✗ 文件不存在: {p}')
            continue
        if p.suffix.lower() not in HANDLERS:
            print(f'✗ 不支持的格式: {p.suffix}（支持 .pptx .docx .xlsx）')
            continue
        try:
            print(f'处理: {p.name}')
            out = process_file(p, None, fonts, print)
            print(f'✓ {out}')
        except Exception as e:
            print(f'✗ {p.name}: {e}')
            import traceback
            traceback.print_exc()


def main() -> None:
    installed = ensure_fonts_installed()
    if installed:
        print(f'已安装字体: {", ".join(installed)}')

    if len(sys.argv) > 1:
        run_cli(sys.argv[1:])
    else:
        from gui import FontFormatApp
        app = FontFormatApp()
        app.mainloop()


if __name__ == '__main__':
    main()

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""writeExcel 入口脚本。

用法：
    python write_excel.py 模板.xlsx 数据.txt [输出.xlsx]
    python write_excel.py            # 进入交互模式（可拖拽文件）
"""

from __future__ import annotations

import sys
from pathlib import Path

# 让脚本无论从哪里双击/运行都能找到同目录的 wexcel 包
_here = Path(__file__).resolve().parent
if str(_here) not in sys.path:
    sys.path.insert(0, str(_here))

from wexcel.cli import main  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())

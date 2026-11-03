# -*- coding: utf-8 -*-
"""环境自检：Python 版本 + 各格式解析依赖。

用法：python doctor.py
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from lib.parser import dependency_report  # noqa: E402


def main() -> int:
    print("=" * 60)
    print("管理体系文件一致性追溯技能 · 环境自检")
    print("=" * 60)
    print(f"Python  : {sys.version.split()[0]}")
    print(f"解释器  : {sys.executable}")
    print()

    rows = dependency_report()
    missing = []
    print(f"{'能力':<20}{'模块':<16}{'状态'}")
    print("-" * 60)
    for label, mod, ok, cmd in rows:
        status = "可用" if ok else "缺失"
        print(f"{label:<20}{mod:<16}{status}")
        if not ok:
            missing.append(cmd)

    print()
    if missing:
        print("缺失依赖会导致对应格式无法解析，标记为盲区（不会中断运行）。")
        print("建议安装：")
        print(f"  {sys.executable} -m pip install " + " ".join(
            c.replace("pip install ", "") for c in missing))
        print()
        print("安装后重新运行本脚本确认。")
        return 1

    print("全部依赖就绪，可完整解析 docx / xlsx / pdf / pptx。")
    return 0


if __name__ == "__main__":
    sys.exit(main())

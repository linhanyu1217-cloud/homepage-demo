# -*- coding: utf-8 -*-
"""检查本项目的 Python 依赖是否就位（只检查，绝不自动安装）。

用法:
    python 脚本\检查依赖.py

输出（纯 ASCII，便于 PowerShell 解析）:
    OK|python-pptx|1.0.2
    MISS|openpyxl|

退出码:
    0 = 全部就位
    1 = 有缺失
"""
import importlib
import sys

# (显示名, import 名, 用途)
DEPS = [
    ("python-docx", "docx",   "读取 Word（公告模板、周报）"),
    ("python-pptx", "pptx",   "生成 PPT 页"),
    ("openpyxl",    "openpyxl", "Excel 结构化数据更新"),
    ("lxml",        "lxml",   "XML 解析支撑"),
    ("Pillow",      "PIL",    "图片处理支撑"),
]


def main():
    missing = []
    for display, mod, _purpose in DEPS:
        try:
            m = importlib.import_module(mod)
            ver = getattr(m, "__version__", "?")
            print("OK|%s|%s" % (display, ver))
        except ImportError:
            print("MISS|%s|" % display)
            missing.append(display)

    if missing:
        print("MISSING_TOTAL|%d" % len(missing))
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())

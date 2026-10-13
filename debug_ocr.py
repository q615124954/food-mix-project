# -*- coding: utf-8 -*-
"""临时调试：导出某张采购单的 OCR 原文，用于定位「员工餐剩余」未识别问题。
用法：python debug_ocr.py "D:\\采购单\\图片.jpg"
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pipeline_min as core


def main(path):
    if not os.path.exists(path):
        print(f"图片不存在：{path}")
        sys.exit(1)
    text = core.call_ocr(path)
    out = path + ".ocr.txt"
    with open(out, "w", encoding="utf-8") as f:
        f.write(text)
    print(f"OCR 原文已保存到: {out}\n")
    print("=" * 60)
    print(text)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print('用法：python debug_ocr.py "D:\\采购单\\图片.jpg"')
        sys.exit(1)
    main(sys.argv[1])

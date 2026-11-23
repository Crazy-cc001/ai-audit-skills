#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""探测 PDF 是否有文字层，决定解析路线。

用法:
    python probe_pdf.py <pdf路径>

判断:
    total_chars > 200  -> 有文字层，用 parse_statement.py 按坐标解析
    total_chars == 0   -> 纯扫描件，用 render_page.py 转图后直接看
"""
import sys
import argparse


def main():
    ap = argparse.ArgumentParser(description="探测 PDF 文字层")
    ap.add_argument("pdf", help="PDF 文件路径")
    a = ap.parse_args()

    try:
        import pymupdf
    except ImportError:
        sys.exit("缺少依赖: pip install pymupdf")

    d = pymupdf.open(a.pdf)
    print(f"文件: {a.pdf}")
    print(f"页数: {len(d)}")
    print()
    print(f"{'页':>4} {'字符数':>8} {'内嵌图':>6} {'页面尺寸':>14}")
    print("-" * 40)
    total = 0
    for i, p in enumerate(d):
        t = p.get_text().strip()
        total += len(t)
        print(f"{i+1:>4} {len(t):>8} {len(p.get_images()):>6} "
              f"{round(p.rect.width):>6}x{round(p.rect.height):<6}")

    print()
    print(f"全文字符总数: {total}")
    print("-" * 40)
    if total == 0:
        print("→ 【纯扫描件，无文字层】")
        print("  1. 先确认同目录有没有电子版对账单（文件名常含账号或日期区间）")
        print("  2. 没有电子版则用 render_page.py 转图，再用看图工具直接看")
        print("  3. ⚠️ 看图时必须读「大写金额」栏做交叉校验（铁律 2）")
    elif total < 200:
        print("→ 【文字层很薄】可能是部分页有文字、部分页是图")
        print("  建议逐页用 render_page.py 转图复核关键词号")
    else:
        print("→ 【有文字层】用 parse_statement.py 按坐标解析")
        print("  注意：文字是碎片化的，必须按 x 列 + y 行重组才能还原成表")


if __name__ == "__main__":
    main()

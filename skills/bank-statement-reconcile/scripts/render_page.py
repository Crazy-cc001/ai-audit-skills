#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 PDF 页渲染成 PNG，支持旋转与左右裁半，供「直接看图」使用。

扫描件常见两种版式：
  A. 整页竖排但方向转了 90°      -> 用 --rotate 90 / 270 转正
  B. 一页并排两张回单（各转 90°） -> 先用 --half L/R 裁半，再各 --rotate 转正

用法:
    python render_page.py <pdf> --page 1                        # 整页
    python render_page.py <pdf> --page 6 --rotate -90           # 转正
    python render_page.py <pdf> --page 6 --half R --rotate -90  # 右半张并转正
    python render_page.py <pdf> --all --dpi 60                  # 全部页缩略图，先摸布局

参数:
    --page N      指定页（从 1 起）；与 --all 二选一
    --all         渲染全部页
    --dpi         分辨率，默认 130（密集表格用 150-200；先摸布局用 55-70）
    --rotate      90 / 180 / 270 / -90，转正用
    --half        L 左半 / R 右半（用于一页两栏版式）
    --out         输出路径（单页时有效），默认 render_p<N>.png
"""
import sys
import argparse


def main():
    ap = argparse.ArgumentParser(description="PDF 页 -> PNG，支持旋转/裁半")
    ap.add_argument("pdf")
    ap.add_argument("--page", type=int, help="页号，从 1 起")
    ap.add_argument("--all", action="store_true", help="渲染全部页")
    ap.add_argument("--dpi", type=int, default=130)
    ap.add_argument("--rotate", type=int, default=0, help="90/180/270/-90")
    ap.add_argument("--half", choices=["L", "R"], help="只取左半或右半")
    ap.add_argument("--out", help="输出文件路径")
    a = ap.parse_args()

    if not a.page and not a.all:
        sys.exit("需指定 --page N 或 --all")

    try:
        import pymupdf
    except ImportError:
        sys.exit("缺少依赖: pip install pymupdf")

    d = pymupdf.open(a.pdf)
    pages = range(len(d)) if a.all else [a.page - 1]
    if not a.all and not (0 <= a.page - 1 < len(d)):
        sys.exit(f"页号越界：共 {len(d)} 页")

    for pno in pages:
        p = d[pno]
        R = p.rect
        clip = None
        if a.half == "L":
            clip = pymupdf.Rect(0, 0, R.width * 0.48, R.height)
        elif a.half == "R":
            clip = pymupdf.Rect(R.width * 0.52, 0, R.width, R.height)

        mat = pymupdf.Matrix(1, 1).prerotate(a.rotate) if a.rotate else None
        kw = {"dpi": a.dpi}
        if clip is not None:
            kw["clip"] = clip
        if mat is not None:
            kw["matrix"] = mat

        px = p.get_pixmap(**kw)
        out = a.out if (a.out and not a.all) else f"render_p{pno+1}.png"
        px.save(out)
        print(f"页{pno+1} -> {out}  ({px.width}x{px.height})")

    print("\n提示：图渲染出来后，用「看图」工具直接看，不要走 OCR —— 密集表格 OCR 会失效。")
    print("⚠️ 看图时务必读「大写金额」栏，与阿拉伯数字双向校验（铁律 2）。")


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""
步骤1: Excel解析筛选
读取目标表格, 识别带填充色标记的数据行, 支持:
- 多机构筛选(--orgs, 对应sheet名, 如 市交易中心)
- 自定义填充色(--colors, 默认 黄FFFFFF00+绿FF92D050)
- 自定义序号区间(--seq-start/--seq-end)与项目编号区间(--code-start/--code-end)
输出匹配清单JSON, 供后续步骤使用。
"""
import argparse
import json
import os
import sys

import openpyxl

sys.stdout.reconfigure(encoding="utf-8")

DEFAULT_COLORS = ["FFFFFF00", "FF92D050"]  # 黄色 + 绿色


def norm_color(c):
    return c.replace("#", "").upper()


def cell_marked(cell, colors):
    fill = cell.fill
    if fill is None or (fill.patternType or "") != "solid":
        return False
    rgb = fill.fgColor.rgb
    if rgb and isinstance(rgb, str):
        return norm_color(rgb) in colors
    return False


def find_col_index(ws, header_row, name):
    for i, c in enumerate(header_row, 1):
        if c.value is not None and str(c.value).strip() == name:
            return i
    return None


def main():
    ap = argparse.ArgumentParser(description="Excel解析筛选(填充色+机构+编号区间)")
    ap.add_argument("--excel", required=True, help="目标Excel路径")
    ap.add_argument("--orgs", default="市交易中心", help="机构(sheet名)列表, 逗号分隔; ALL=全部sheet")
    ap.add_argument("--colors", default=",".join(DEFAULT_COLORS), help="填充色hex, 逗号分隔")
    ap.add_argument("--seq-start", type=int, default=None, help="序号下限(含)")
    ap.add_argument("--seq-end", type=int, default=None, help="序号上限(含)")
    ap.add_argument("--code-start", default=None, help="项目编号下限(含, 字符串比较)")
    ap.add_argument("--code-end", default=None, help="项目编号上限(含, 字符串比较)")
    ap.add_argument("--out", default=None, help="匹配清单JSON输出路径(默认: Excel同目录下 匹配清单.json)")
    args = ap.parse_args()

    colors = {norm_color(c) for c in args.colors.split(",") if c.strip()}
    out = args.out or os.path.join(os.path.dirname(os.path.abspath(args.excel)), "匹配清单.json")

    wb = openpyxl.load_workbook(args.excel, data_only=True)
    orgs = wb.sheetnames if args.orgs.strip().upper() == "ALL" else [o.strip() for o in args.orgs.split(",") if o.strip()]

    results = []
    for org in orgs:
        if org not in wb.sheetnames:
            print(f"⚠️ 机构sheet不存在: {org}, 跳过")
            continue
        ws = wb[org]
        header = ws[1]
        idx_seq = find_col_index(ws, header, "序号")
        idx_code = find_col_index(ws, header, "项目编号")
        idx_name = find_col_index(ws, header, "项目名称")
        if idx_code is None:
            print(f"⚠️ [{org}] 未找到'项目编号'列, 跳过该sheet")
            continue
        headers = [str(c.value).replace("\n", "") if c.value is not None else f"列{c.column}" for c in header]
        n = 0
        for r in range(2, ws.max_row + 1):
            row_cells = [ws.cell(r, c) for c in range(1, ws.max_column + 1)]
            marked = any(cell_marked(c, colors) for c in row_cells)
            if not marked:
                continue
            seq = str(ws.cell(r, idx_seq).value).strip() if idx_seq and ws.cell(r, idx_seq).value is not None else ""
            code = str(ws.cell(r, idx_code).value).strip() if ws.cell(r, idx_code).value is not None else ""
            name = str(ws.cell(r, idx_name).value).strip() if idx_name and ws.cell(r, idx_name).value is not None else ""
            # 序号区间
            if seq and (args.seq_start is not None or args.seq_end is not None):
                try:
                    seq_i = int(float(seq))
                except ValueError:
                    continue
                if args.seq_start is not None and seq_i < args.seq_start:
                    continue
                if args.seq_end is not None and seq_i > args.seq_end:
                    continue
            # 编号区间(字符串比较)
            if code:
                if args.code_start is not None and code < args.code_start:
                    continue
                if args.code_end is not None and code > args.code_end:
                    continue
            item = {"机构": org, "序号": seq, "项目编号": code, "项目名称": name}
            for c in range(1, ws.max_column + 1):
                v = ws.cell(r, c).value
                item[headers[c - 1]] = "" if v is None else str(v).strip()
            results.append(item)
            n += 1
        print(f"📋 [{org}] 匹配 {n} 行")

    with open(out, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\n✅ 共匹配 {len(results)} 条 → {out}")
    for it in results:
        print(f"   {it['机构']} | 序号{it['序号']} | {it['项目编号']} | {it['项目名称']}")


if __name__ == "__main__":
    main()

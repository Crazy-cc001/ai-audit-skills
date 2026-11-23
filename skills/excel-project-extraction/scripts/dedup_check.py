# -*- coding: utf-8 -*-
"""
步骤5: 表格数据去重校验
对指定文件夹内每张Excel, 对指定列单列查重(不跨列比对), 输出查重报告;
可选: --action mark 用颜色标记重复单元格 / --action remove 生成去重后副本(保留首条)。
"""
import argparse
import os
import sys

import openpyxl
import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")

ALL_COLS = ["文件名", "投标人名称", "注册地址", "联系人", "联系人电话",
            "电子邮件", "法定代表人", "法定代表人电话", "技术负责人", "技术负责人电话"]


def scan(folder, columns):
    files = sorted(f for f in os.listdir(folder)
                   if f.lower().endswith(".xlsx") and not f.startswith("~$") and f != "列内查重结果.xlsx")
    summary_rows, detail_rows, per_file = [], [], {}
    for f in files:
        path = os.path.join(folder, f)
        df = pd.read_excel(path, dtype=str)
        dup_info = {}
        for col in columns:
            if col not in df.columns:
                continue
            s = df[col].fillna("").str.strip()
            s = s[s != ""]
            counts = s.value_counts()
            dups = counts[counts > 1]
            dup_info[col] = s
            for val, cnt in dups.items():
                summary_rows.append({"项目表": f, "所属列": col, "重复值": val, "出现次数": cnt})
                rows = df[s == val]
                for _, r in rows.iterrows():
                    detail_rows.append({
                        "项目表": f, "所属列": col, "重复值": val,
                        **{c: r.get(c, "") for c in ALL_COLS if c in df.columns},
                    })
        per_file[f] = (df, dup_info)
    return files, summary_rows, detail_rows, per_file


def write_report(folder, summary_rows, detail_rows):
    out = os.path.join(folder, "列内查重结果.xlsx")
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        if summary_rows:
            pd.DataFrame(summary_rows).to_excel(w, sheet_name="查重汇总", index=False)
            pd.DataFrame(detail_rows).to_excel(w, sheet_name="重复明细(完整行)", index=False)
        else:
            pd.DataFrame([{"结果": "未发现任何列内重复内容"}]).to_excel(w, sheet_name="查重报告(无重复)", index=False)
    return out


def mark_duplicates(folder, per_file, mark_color):
    """在原表上以指定颜色标记所有重复值单元格(仅标记重复列内的值, 空值除外)"""
    color = mark_color.replace("#", "").upper()
    if not (len(color) == 8 or len(color) == 6):
        color = "FFFF0000"
    for f, (df, dup_info) in per_file.items():
        path = os.path.join(folder, f)
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        header = {str(c.value).strip(): c.column for c in ws[1] if c.value is not None}
        for col, s in dup_info.items():
            if col not in header:
                continue
            # 找出重复值集合
            counts = s.value_counts()
            dup_vals = set(counts[counts > 1].index)
            col_idx = header[col]
            for r in range(2, ws.max_row + 1):
                v = ws.cell(r, col_idx).value
                if v is None:
                    continue
                if str(v).strip() in dup_vals:
                    cell = ws.cell(r, col_idx)
                    cell.fill = openpyxl.styles.PatternFill(fill_type="solid", fgColor=color)
        wb.save(path)
        print(f"   ✅ 已标记: {f}")


def remove_duplicates(folder, per_file):
    """生成去重后副本 <原名>_去重后.xlsx(重复值仅保留首次出现的行)"""
    for f, (df, dup_info) in per_file.items():
        mask = pd.Series(True, index=df.index)
        for col, s in dup_info.items():
            counts = s.value_counts()
            dup_vals = counts[counts > 1].index
            if len(dup_vals) == 0:
                continue
            first_seen = {}
            for i, v in s.items():
                if v in dup_vals:
                    if v in first_seen:
                        mask.loc[i] = False
                    else:
                        first_seen[v] = i
        cleaned = df[mask]
        out = os.path.join(folder, os.path.splitext(f)[0] + "_去重后.xlsx")
        cleaned.to_excel(out, index=False)
        print(f"   ✅ 去重副本: {out} ({len(df)}→{len(cleaned)}行)")


def main():
    ap = argparse.ArgumentParser(description="表格单列查重(标记/剔除)")
    ap.add_argument("--folder", required=True, help="存放Excel汇总表的文件夹")
    ap.add_argument("--columns", default="文件名,注册地址,联系人电话,电子邮件",
                    help="查重列, 逗号分隔")
    ap.add_argument("--action", default="report", choices=["report", "mark", "remove"],
                    help="report=仅报告 / mark=报告+颜色标记 / remove=报告+生成去重副本")
    ap.add_argument("--mark-color", default="FFFF0000", help="标记颜色hex(默认红色)")
    args = ap.parse_args()

    columns = [c.strip() for c in args.columns.split(",") if c.strip()]
    files, summary_rows, detail_rows, per_file = scan(args.folder, columns)
    print(f"扫描 {len(files)} 张表, 查重列: {columns}\n")

    if not summary_rows:
        print("🎉 未发现任何列内重复内容")
    else:
        sm = pd.DataFrame(summary_rows)
        print(f"查重汇总 ({len(sm)} 组重复):")
        print(sm.to_string(max_colwidth=40))

    out = write_report(args.folder, summary_rows, detail_rows)
    print(f"\n📁 查重报告: {out}")

    if summary_rows and args.action == "mark":
        print("\n标记重复单元格...")
        mark_duplicates(args.folder, per_file, args.mark_color)
    elif summary_rows and args.action == "remove":
        print("\n生成去重副本...")
        remove_duplicates(args.folder, per_file)


if __name__ == "__main__":
    main()

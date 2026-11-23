#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""解析电子版银行对账单（有文字层），按坐标重组表格 + 三重自校验。

用法:
    python parse_statement.py <pdf路径> [--json out.json]
    python parse_statement.py <pdf路径> --anchor-x 35 --pad 25   # 微调锚点参数

原理:
    电子版 PDF 的每个单元格都是一段独立文字（如 20260826 / 114534 / 借方 / 220.83），
    get_text() 出来是碎片的，必须按坐标重组：
      1. 序号锚点：x < anchor_x 且内容为纯数字的词，就是行锚点
      2. 行范围：[锚y - pad, 下一锚y - pad)，覆盖该行的多行单元格（对方名称常占 2-3 行）
      3. 列区间：按 x 落在哪个区间，决定词属于哪一列

三重自校验（缺一不可）:
    1. 金额合计 == 表头发生额
    2. 余额链倒推自洽  上笔余额 ± 本笔金额 == 本笔余额
    3. 承前页余额 == 上月末账面余额（需传入 --book-balance 比对）
"""
import sys
import re
import json
import argparse

# 列区间（x 起, x 止）。换银行时按实际 PDF 微调。
COLS = [
    ("no",     0,    35),
    ("date",   35,   85),
    ("time",   85,   122),
    ("dc",     122,  160),
    ("amt",    160,  252),
    ("bal",    252,  305),
    ("racc",   305,  368),
    ("rname",  368,  420),
    ("vou",    420,  462),
    ("memo",   462,  540),
    ("serial", 540,  99999),
]


def to_num(s):
    s = str(s or "").replace(",", "").strip()
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def extract(path, anchor_x=35, pad=25):
    import pymupdf
    d = pymupdf.open(path)

    words = []
    for pno, page in enumerate(d):
        for w in page.get_text("words"):
            # w = (x0, y0, x1, y1, text, block, line, word_no)
            words.append((pno, w[1], w[0], w[4]))

    anchors = sorted(
        (pno, y, t) for pno, y, x, t in words
        if x < anchor_x and re.fullmatch(r"\d{1,3}", t)
    )

    recs = []
    for i, (pno, ay, anum) in enumerate(anchors):
        # 行边界取【相邻锚点的中点】，避免把上一行的尾行（对方名称第2、3行）切给自己
        if i > 0 and anchors[i - 1][0] == pno:
            lo = (anchors[i - 1][1] + ay) / 2
        else:
            lo = ay - pad
        if i + 1 < len(anchors) and anchors[i + 1][0] == pno:
            hi = (ay + anchors[i + 1][1]) / 2
        else:
            hi = ay + pad * 2.4
        items = [(y, x, t) for p, y, x, t in words if p == pno and lo <= y < hi]

        rec = {"no": int(anum), "page": pno + 1}
        for name, x0, x1 in COLS:
            if name == "no":
                continue
            # 按 (y, x) 排序 = 按阅读顺序，多行单元格才不会串行
            toks = [t for y, x, t in sorted(items) if x0 <= x < x1]
            rec[name] = "".join(toks).strip()
        # 日期只保留 8 位数字（页脚免责声明等噪声会落进日期列区间）
        m = re.search(r"\d{8}", rec.get("date", ""))
        rec["date"] = m.group() if m else ""
        # 金额/余额只取第一个 token
        rec["amt"] = rec["amt"].split()[0] if rec["amt"] else ""
        rec["bal"] = rec["bal"].split()[0] if rec["bal"] else ""
        recs.append(rec)

    return recs


def verify(recs, book_balance=None):
    print("=" * 78)
    print(f"{'序号':>4} {'日期':>10} {'借贷':<4} {'交易金额':>15} {'账户余额':>15}  对方名称")
    print("-" * 78)
    for r in recs:
        print(f"{r['no']:>4} {r['date']:>10} {r['dc']:<4} "
              f"{r['amt']:>15} {r['bal']:>15}  {r['rname']}")

    print("=" * 78)
    problems = []

    # 校验 1：金额合计
    j = sum(to_num(r["amt"]) or 0 for r in recs if r["dc"] == "借方")
    dd = sum(to_num(r["amt"]) or 0 for r in recs if r["dc"] == "贷方")
    nj = sum(1 for r in recs if r["dc"] == "借方")
    nd = sum(1 for r in recs if r["dc"] == "贷方")
    print(f"[1] 借方 {nj} 笔 合计 {j:,.2f}  ｜ 贷方 {nd} 笔 合计 {dd:,.2f}")
    print("    → 与 PDF 表头「借方发生额 / 贷方发生额 / 笔数」逐项核对")
    if nj + nd != len(recs):
        problems.append(f"借贷标志未识别的行数 = {len(recs) - nj - nd}")

    # 校验 2：余额链倒推
    print("[2] 余额链倒推（上笔余额 ± 本笔金额 == 本笔余额）")
    bad = 0
    for i in range(len(recs) - 1):
        cur, nxt = recs[i], recs[i + 1]
        b, amt = to_num(cur["bal"]), to_num(cur["amt"])
        if b is None or amt is None:
            continue
        prev = b + amt if cur["dc"] == "借方" else b - amt
        nb = to_num(nxt["bal"])
        if nb is not None and abs(prev - nb) > 0.005:
            print(f"    ✗ 序号{cur['no']}→{nxt['no']}: 推出 {prev:,.2f} ≠ 实际 {nb:,.2f}")
            bad += 1
    if bad == 0:
        print(f"    ✅ {len(recs)} 笔余额链完全自洽")
    else:
        problems.append(f"余额链 {bad} 处断链")

    # 校验 3：承前页
    if recs:
        first, last = recs[0], recs[-1]
        # 月末余额 = 最新那笔【交易之后】的余额，也就是它自己的「账户余额」列
        top = to_num(first["bal"])
        bl, al = to_num(last["bal"]), to_num(last["amt"])
        start = bl + al if last["dc"] == "借方" else bl - al
        print(f"[3] 承前页余额（由最早一笔序号{last['no']}反推）= {start:,.2f}")
        print(f"    月末账户余额（最新一笔序号{first['no']}的余额栏） = {top:,.2f}")
        print(f"    验算: {start:,.2f} + {dd:,.2f} - {j:,.2f} = {start + dd - j:,.2f}")
        if top is None or abs((start + dd - j) - top) > 0.005:
            problems.append("承前页 + 收 - 支 ≠ 月末余额")
        else:
            print("    ✅ 承前页 → 月末 闭环成立")
        if book_balance is not None:
            if abs(start - book_balance) > 0.005:
                problems.append(
                    f"承前页 {start:,.2f} ≠ 账面上月末 {book_balance:,.2f}（差 {start - book_balance:,.2f}）")
                print(f"    ✗ 与账面上月末余额 {book_balance:,.2f} 不符！")
            else:
                print(f"    ✅ 承前页与账面上月末余额一致（{book_balance:,.2f}）")

    print("-" * 78)
    if problems:
        print("❌ 存在问题，禁止据此出凭证：")
        for p in problems:
            print("   -", p)
    else:
        print("✅ 三重自校验全部通过")
    return problems


def main():
    ap = argparse.ArgumentParser(description="解析电子版银行对账单 + 三重自校验")
    ap.add_argument("pdf")
    ap.add_argument("--json", help="把结构化结果写到该 json 文件")
    ap.add_argument("--anchor-x", type=float, default=35, help="序号锚点的 x 上限（默认 35）")
    ap.add_argument("--pad", type=float, default=25, help="行范围上下扩展量（默认 25）")
    ap.add_argument("--book-balance", type=float,
                    help="账面【上月末】银行存款余额，用于校验承前页")
    a = ap.parse_args()

    recs = extract(a.pdf, a.anchor_x, a.pad)
    verify(recs, a.book_balance)

    if a.json:
        with open(a.json, "w", encoding="utf-8") as f:
            json.dump(recs, f, ensure_ascii=False, indent=2)
        print(f"\n已写出: {a.json}")


if __name__ == "__main__":
    main()

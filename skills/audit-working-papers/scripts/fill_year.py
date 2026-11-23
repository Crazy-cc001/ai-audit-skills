# -*- coding: utf-8 -*-
"""按年度将账套数据填入底稿各表对应年列
用法: python fill_year.py <年度>
- 货币资金(现金/银行/5家银行), 其他应收款/应付款明细, 非限定性明细, 预提/长期借款明细
- 2025额外: 明细账龄表、基准日表、Sheet1现金银行、124行非限定性资产表重建
"""
import json, os, sys
import openpyxl
from openpyxl.utils import get_column_letter

YEAR = sys.argv[1]
BASE = r'C:\path\to\temp\xsj_pt'
TPL = r'C:\path\to\底稿模板.xlsx'
p = os.path.join(BASE, f'out{YEAR}', f'{YEAR}_datasets.json')
if not os.path.exists(p):
    print(f'[{YEAR}] 无数据集, 跳过')
    sys.exit(0)
ds = json.load(open(p, encoding='utf-8'))
acc = ds['accounts']
YEARS = list(range(2016, 2026))
COL = 3 + (int(YEAR) - 2016)   # 该年列

def qm(codes):
    return round(abs(sum(acc[c]['agg_qm'] for c in codes if c in acc)), 2)

def sub_codes(prefix):
    return sorted([c for c in acc if c == prefix or c.startswith(prefix)])

def setv(ws, row, col, v, fmt='#,##0.00'):
    if v is None:
        return
    try:
        c = ws.cell(row=row, column=col)
        c.value = round(v, 2)
        c.number_format = fmt
    except (AttributeError, TypeError):
        pass

def sets(ws, row, col, v):
    try:
        ws.cell(row=row, column=col, value=v)
    except AttributeError:
        pass

wb = openpyxl.load_workbook(TPL)
done = []

# ---- 货币资金: 现金/银行/各行 ----
ws = wb['货币资金']
setv(ws, 8, COL, qm(['1001.']))
setv(ws, 9, COL, qm(['1002.']))
banks = [c for c in sub_codes('1002.') if c != '1002.']
banks.sort()
if len(banks) > 5:   # 行10-14共5行, 超出并入'其他'
    extra = banks[4:]
    banks = banks[:4] + [None]
for i, c in enumerate(banks):
    r = 10 + i
    if c is None:
        sets(ws, r, 2, '其他')
        setv(ws, r, COL, qm(extra))
    else:
        sets(ws, r, 2, acc[c]['name'])
        setv(ws, r, COL, qm([c]))
done.append(f'货币资金{len(banks)}家银行')

# ---- 其他应收款审定表. (应收账款/其他应收款, 带符号) ----
ws = wb['其他应收款审定表.']
def signed_qm(codes):
    return round(sum(acc[c]['agg_qm'] for c in codes if c in acc), 2)
setv(ws, 7, COL, signed_qm(['1121.']))
setv(ws, 8, COL, signed_qm(['1122.']))

# ---- 其他应付款审定表(十年) ----
ws = wb['其他应付款审定表']
setv(ws, 7, COL, signed_qm(['2202.']))
setv(ws, 8, COL, signed_qm(['2209.']))

# ---- 非限定性资产审定表 明细 ----
ws = wb['非限定性资产审定表']
for r, code in [(8, '3101.01.'), (9, '3101.02.'), (10, '3101.02.01.'),
                (11, '3101.02.02.'), (12, '3101.02.03.'), (13, '3101.02.04.'),
                (14, '3101.03.')]:
    setv(ws, r, COL, qm([code]))

# ---- 应交税费: 二级科目行 ----
ws = wb['应交税费审定表']
subs = [c for c in sub_codes('2206.') if c != '2206.']
subs.sort()
for i, c in enumerate(subs[:10]):
    sets(ws, 8 + i, 2, acc[c]['name'])
    setv(ws, 8 + i, COL, qm([c]))
if subs:
    end_row = 7 + len(subs)
    for col in range(3, 13):
        L = get_column_letter(col)
        ws.cell(row=9, column=col, value=f'=SUM({L}7:{L}{end_row})')
done.append(f'应交税费{len(subs)}个二级')

# ---- 2025年专属: 基准日表 ----
if YEAR == '2025':
    # 明细账龄表: 其他应收款(1122)/预付(1141)/应付(2202)/预收(2203)/其他应付(2209) 末级明细
    def aging_rows(sheet, prefix, col):
        wsx = wb[sheet]
        def is_leaf(c):
            return not any(k != c and k.startswith(c) for k in acc)
        codes = [c for c in acc if c.startswith(prefix) and c != prefix and is_leaf(c)]
        detail = [(c, acc[c]['name'], acc[c]['agg_qm']) for c in codes
                  if acc[c]['agg_qm'] and abs(acc[c]['agg_qm']) > 0.005]
        total_row = None
        for r in range(8, 60):
            b = wsx.cell(row=r, column=2).value
            if b and '合计' in str(b):
                total_row = r
                break
        # 清旧明细(8..合计前一行)
        if total_row:
            for r in range(8, total_row):
                sets(wsx, r, 2, None)
                setv(wsx, r, col, None)
        n = 0
        for i, (c, name, v) in enumerate(detail):
            r = 8 + i
            sets(wsx, r, 2, name)
            setv(wsx, r, col, abs(v))
            n = i + 1
        if total_row:
            L = get_column_letter(col)
            wsx.cell(row=total_row, column=col, value=f'=SUM({L}8:{L}{total_row-1})')
        return n
    n1 = aging_rows('其他应收款明细及账龄分析表', '1122.', 3)
    n2 = aging_rows('其他应收款明细及账龄分析表 ', '1122.', 3)
    n3 = aging_rows('预付账款明细及账龄分析表', '1141.', 4)
    n4 = aging_rows('应付账款明细及账龄分析表', '2202.', 4)
    n5 = aging_rows('预收账款明细及账龄分析表', '2203.', 4)
    n6 = aging_rows('其他应付款明细及账龄分析表', '2209.', 4)
    done.append(f'账龄表明细 {n1}/{n2}/{n3}/{n4}/{n5}/{n6}')
    # Sheet1 现金/银行
    ws = wb['Sheet1']
    setv(ws, 7, 2, qm(['1001.']))
    setv(ws, 7, 3, qm(['1001.']))
    setv(ws, 8, 2, qm(['1002.']))
    setv(ws, 8, 3, qm(['1002.']))
    # 应交税费右表(年初/本年/期末)
    ws = wb['应交税费审定表']
    for i, c in enumerate(subs):
        if i >= 9:
            break
        sets(ws, 9 + i, 16, acc[c]['name'])
        setv(ws, 9 + i, 19, abs(acc[c]['ncye']))
        setv(ws, 9 + i, 20, acc[c]['agg_jf'] + acc[c]['agg_df'])
        setv(ws, 9 + i, 21, abs(acc[c]['agg_qm']))
    done.append('应交税费右表')

wb.save(TPL)
print(f'[{YEAR}] 年度填充完成: ' + '; '.join(done))

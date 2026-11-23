# -*- coding: utf-8 -*-
r"""账套余额 vs 协会报表 核对表 (按年度参数化)
用法: python recon_year.py <年度>
输出: out<年度>\<年度>_账套与报表核对表.xlsx + 控制台摘要
"""
import sys, json, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side

YEAR = sys.argv[1]
ENTITY = '某行业协会'   # ← 换客户时修改单位名称
BASE = r'C:\path\to\temp\xsj_pt'
OUT = os.path.join(BASE, f'out{YEAR}')

d = json.load(open(os.path.join(OUT, f'{YEAR}_datasets.json'), encoding='utf-8'))
acc = d['accounts']
st = json.load(open(os.path.join(BASE, 'statements.json'), encoding='utf-8'))[YEAR]['bs']

def qm(codes):
    return round(abs(sum(acc[c]['agg_qm'] for c in codes if c in acc)), 2)

def nc(codes):
    return round(abs(sum(acc[c]['ncye'] for c in codes if c in acc)), 2)

MAP = [
    ('货币资金', ['1001.', '1002.'], '资产'),
    ('短期投资', ['1201.'], '资产'),
    ('应收款项', ['1121.', '1122.'], '资产'),
    ('预付账款', ['1141.'], '资产'),
    ('存  货', ['1301.'], '资产'),
    ('待摊费用', ['1401.'], '资产'),
    ('长期股权投资', ['1801.'], '资产'),
    ('长期债权投资', ['1802.'], '资产'),
    ('固定资产原价', ['1501.'], '资产'),
    ('减：累计折旧', ['1502.'], '资产'),
    ('在建工程', ['1601.'], '资产'),
    ('文物文化资产', ['1701.'], '资产'),
    ('无形资产', ['1901.'], '资产'),
    ('受托代理资产', ['1902.'], '资产'),
    ('短期借款', ['2101.'], '负债'),
    ('应付款项', ['2202.', '2209.'], '负债'),
    ('应付工资', ['2204.'], '负债'),
    ('应交税金', ['2206.'], '负债'),
    ('预收账款', ['2203.'], '负债'),
    ('预提费用', ['2601.'], '负债'),
    ('预计负债', ['2701.'], '负债'),
    ('长期借款', ['2501.'], '负债'),
    ('长期应付款', ['2205.'], '负债'),
    ('受托代理负债', ['2801.'], '负债'),
    ('非限定性净资产', ['3101.'], '净资产'),
    ('限定性净资产', ['3102.'], '净资产'),
]

rows = []
for name, codes, cat in MAP:
    b_nc, b_qm = st.get(name, (None, None))
    if name == '减：累计折旧':
        b_nc = None if b_nc is None else -b_nc
        b_qm = None if b_qm is None else -b_qm
    k_nc, k_qm = nc(codes), qm(codes)
    rows.append(dict(name=name, cat=cat, b_nc=b_nc, b_qm=b_qm, k_nc=k_nc, k_qm=k_qm,
                     diff_nc=None if b_nc is None else round(b_nc - k_nc, 2),
                     diff_qm=None if b_qm is None else round(b_qm - k_qm, 2)))
b_nc, b_qm = st.get('固定资产净值', (None, None))
fnc = round(nc(['1501.']) - nc(['1502.']), 2)
fqm = round(qm(['1501.']) - qm(['1502.']), 2)
rows.append(dict(name='固定资产净值', cat='资产', b_nc=b_nc, b_qm=b_qm,
                 k_nc=fnc, k_qm=fqm,
                 diff_nc=None if b_nc is None else round(b_nc - fnc, 2),
                 diff_qm=None if b_qm is None else round(b_qm - fqm, 2)))

# ---- 写 xlsx ----
wb = Workbook()
ws = wb.active
ws.title = f'{YEAR}核对表'
title_font = Font(name='微软雅黑', size=12, bold=True)
head_font = Font(name='微软雅黑', size=9, bold=True)
cell_font = Font(name='微软雅黑', size=9)
red_font = Font(name='微软雅黑', size=9, color='CC0000', bold=True)
thin = Side(style='thin')
border = Border(left=thin, right=thin, top=thin, bottom=thin)
head_fill = PatternFill('solid', fgColor='D9E1F2')
diff_fill = PatternFill('solid', fgColor='FFC7CE')

ws.append([f'{ENTITY} {YEAR}年度 账套总账与报表核对表'])
ws.merge_cells('A1:I1')
ws['A1'].font = title_font
ws.append(['报表项目', '类别', '报表年初数', '账套年初数', '年初差异',
           '报表期末数', '账套期末数', '期末差异', '核对结论'])
for i in range(1, 10):
    c = ws.cell(row=2, column=i)
    c.font = head_font
    c.fill = head_fill
    c.border = border
n = 3
diffs = []
for r in rows:
    ws.cell(row=n, column=1, value=r['name']).font = cell_font
    ws.cell(row=n, column=2, value=r['cat']).font = cell_font
    for col, key in [(3, 'b_nc'), (4, 'k_nc'), (6, 'b_qm'), (7, 'k_qm')]:
        v = r[key]
        if v is not None:
            ws.cell(row=n, column=col, value=v).number_format = '#,##0.00'
    for col, key in [(5, 'diff_nc'), (8, 'diff_qm')]:
        v = r[key]
        if v is not None:
            cell = ws.cell(row=n, column=col, value=v)
            cell.number_format = '#,##0.00'
            if abs(v) > 0.005:
                cell.fill = diff_fill
                cell.font = red_font
    d1, d2 = r['diff_nc'], r['diff_qm']
    if d1 is not None and d2 is not None:
        if abs(d1) <= 0.005 and abs(d2) <= 0.005:
            concl = '相符'
        else:
            concl = f'差异{d2 if d2 else d1:,.2f}'
            diffs.append((r['name'], r['cat'], d1 or 0.0, d2 or 0.0))
        ws.cell(row=n, column=9, value=concl).font = red_font if concl != '相符' else cell_font
    for col in range(1, 10):
        ws.cell(row=n, column=col).border = border
    n += 1
for col, w in zip('ABCDEFGHI', [18, 8, 14, 14, 12, 14, 14, 12, 14]):
    ws.column_dimensions[col].width = w
wb.save(os.path.join(OUT, f'{YEAR}_账套与报表核对表.xlsx'))

# ---- 控制台摘要 ----
print(f'[{YEAR}] 核对: 共{len(rows)}项')
for r in rows:
    d1, d2 = r['diff_nc'], r['diff_qm']
    if d1 is not None and (abs(d1) > 0.005 or (d2 is not None and abs(d2) > 0.005)):
        print(f'  !! {r["name"]}: 年初差{d1:,.2f} 期末差{d2 if d2 is not None else 0:,.2f}')
if diffs:
    print(f'[{YEAR}] !! 差异项 {len(diffs)} 个(详见核对表)')
else:
    print(f'[{YEAR}] 全部相符')

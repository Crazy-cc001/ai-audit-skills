# -*- coding: utf-8 -*-
"""解析协会财务报表 2016-2025（资产负债表+业务活动表）→ 统一JSON数据集"""
import os, re, json, glob
import openpyxl

ROOT = r'C:\path\to\协会财务报表'
OUT = r'C:\path\to\temp\xsj_pt\statements.json'

def num(v):
    if v is None or v == '':
        return None
    try:
        return round(float(v), 2)
    except (TypeError, ValueError):
        return None

data = {}

# ---------- 资产负债表 ----------
for f in sorted(glob.glob(os.path.join(ROOT, '16-25资产负债表', '*.xlsx'))):
    m = re.search(r'(20\d\d)资产负债表', os.path.basename(f))
    if not m:
        continue
    year = int(m.group(1))
    wb = openpyxl.load_workbook(f, data_only=True)
    ws = wb['表页_1']
    bs = {}
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        if not row or len(row) < 8:
            continue
        left_name, left_nc, left_qm = row[0], row[2], row[3]
        right_name, right_nc, right_qm = row[4], row[6], row[7]
        for name, nc, qm in ((left_name, left_nc, left_qm), (right_name, right_nc, right_qm)):
            if name is None:
                continue
            n = str(name).strip().replace('\u3000', '').strip()
            if not n:
                continue
            if n in ('资产', '负债和净资产', '行次'):
                continue
            bs[n] = (num(nc), num(qm))
    wb.close()
    data[year] = {'bs': bs}
    zc = num([v for k, v in bs.items() if k == '资产总计'][0][1]) if '资产总计' in bs else None
    print(f'{year} 资产负债表: {len(bs)} 项, 资产总计期末={zc}')

# ---------- 业务活动表 ----------
for f in sorted(glob.glob(os.path.join(ROOT, '16-25年业务活动表', '*.xlsx'))):
    m = re.search(r'(20\d\d)业务活动表', os.path.basename(f))
    if not m:
        continue
    year = int(m.group(1))
    wb = openpyxl.load_workbook(f, data_only=True)
    ws = wb['表页_1']
    ba = {}
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True):
        if not row or len(row) < 6:
            continue
        name = row[0]
        if name is None:
            continue
        n = str(name).strip().replace('\u3000', '').strip()
        if not n:
            continue
        benyue = num(row[3])   # 本月数合计
        bennian = num(row[5])  # 本年数合计
        if n in ('项目', '行次', '其中：', '一、收入', '二、费用', '三、限定性净资产转为非限定性净资产', '四、净资产变动额（若为减少以-号填列）'):
            ba.setdefault(n, []).append((benyue, bennian))
            continue
        # 同名列可能多行, 取第一行有效值
        if n not in ba:
            ba[n] = (benyue, bennian)
    wb.close()
    data.setdefault(year, {})['ba'] = ba
    sr = ba.get('收入合计')
    print(f'{year} 业务活动表: {len(ba)} 项, 收入合计本年={sr}')

json.dump(data, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('已写出', OUT)

# ---------- 校验: 各年资产负债表平衡 ----------
print()
for y in sorted(data):
    bs = data[y]['bs']
    zc = bs.get('资产总计')
    fz = bs.get('负债和净资产总计')
    if zc and fz:
        flag = '✓' if abs(zc[1] - fz[1]) < 0.01 else f'✗ 差{zc[1]-fz[1]:,.2f}'
        print(f'{y}: 资产期末 {zc[1]:>16,.2f} vs 负债净资产 {fz[1]:>16,.2f}  {flag}')

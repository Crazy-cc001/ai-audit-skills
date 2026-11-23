# -*- coding: utf-8 -*-
r"""
新纪元账套年度数据提取（可复用）
用法: python extract_year.py <MyTemp0快照目录> <年度> <输出目录>
输出: <输出目录>\<年度>_datasets.json + <年度>_总账科目余额表.xlsx + <年度>_明细账_<科目>.xlsx
"""
import sys, os, json, collections
sys.path.insert(0, r'C:\path\to\temp\xsj_pt\pypx\src\pypxlib-2.5')
from pypxlib import Table

ENTITY = '某行业协会'   # ← 换客户时修改单位名称
SNAP, YEAR = sys.argv[1], sys.argv[2]
OUT = sys.argv[3]
os.makedirs(OUT, exist_ok=True)

def fnum(b):
    """pypxlib BCD字段返回bytes(b'5178.00')或float"""
    if b is None:
        return 0.0
    if isinstance(b, bytes):
        s = b.decode('ascii', 'replace').replace('\x00', '').strip()
        try:
            return float(s) if s else 0.0
        except ValueError:
            return 0.0
    if isinstance(b, str):
        try:
            return float(b.strip()) if b.strip() else 0.0
        except ValueError:
            return 0.0
    try:
        return float(b)
    except (TypeError, ValueError):
        return 0.0

def read_table(name):
    p = os.path.join(SNAP, name)
    if not os.path.exists(p):
        return None
    t = Table(p, encoding='gbk')
    rows = []
    for i in range(len(t)):
        r = t[i]
        rows.append({k: r[k] for k in t.fields})
    t.close()
    return rows

print(f'[{YEAR}] 读取快照 {SNAP}')

# ---------- 科目字典 ----------
km_rows = read_table('Km.DB') or []
km_by_id = {r['FAccountID']: r for r in km_rows}
print(f'[{YEAR}] 科目 {len(km_rows)} 个')

# ---------- 期初余额 ----------
ye_rows = read_table('Kmye.DB') or []
ye_by_km = {}
for r in ye_rows:
    ye_by_km[r['Kmdm'].strip()] = r
print(f'[{YEAR}] 余额行 {len(ye_rows)}')

# ---------- 凭证 ----------
jz_rows = read_table('Jzpz.DB') or []
vouchers = []
for r in jz_rows:
    vouchers.append(dict(
        inc=r['IncNo'], pzrq=str(r['Pzrq']).strip(), date=str(r['Pz_Date'])[:10],
        pzlx=(r['Pzlx_Mc'] or '').strip(), pzbh=(r['Pzbh'] or '').strip(),
        jd=(r['Jd'] or '').strip(), km=(r['Kmdm'] or '').strip(),
        bmdm=(r['Bmdm'] or '').strip(), wldm=(r['Wldm'] or '').strip(),
        xmdm=(r['Xmdm'] or '').strip(), zy=(r['Zy'] or '').strip(),
        sl=fnum(r['Sl']), rmb=fnum(r['Rmb']), wbje=fnum(r['Wbje']), hl=fnum(r['Hl']),
        dfkm=(r['Dfkm'] or '').strip() if isinstance(r['Dfkm'], (str, bytes)) else str(r['Dfkm'] or ''),
        kjqj=r['Kjqj'], fjs=(r['Fjzs'] if r['Fjzs'] else None),
        sr=(r['Sr'] or '').strip() if r['Sr'] else '', sh=(r['Sh'] or '').strip() if r['Sh'] else '',
        jzr=(r['Jzr'] or '').strip() if r['Jzr'] else '',
    ))
print(f'[{YEAR}] 凭证分录 {len(vouchers)}')

# ---------- 分月汇总 ----------
ps_rows = read_table('P_SUM.DB') or []
print(f'[{YEAR}] 月度汇总 {len(ps_rows)}')

# ---------- 项目字典 ----------
xm_rows = read_table('Xm.DB') or []
print(f'[{YEAR}] 项目 {len(xm_rows)}')

# ---------- 科目类型统计 ----------
types = collections.Counter(r['KM_TYPE'] for r in km_rows)
print(f'[{YEAR}] KM_TYPE 分布: {dict(types)}')
yefx = collections.Counter(r['KM_YEFX'] for r in km_rows if r['KM_YEFX'])
print(f'[{YEAR}] KM_YEFX 分布: {dict(yefx)}')

# ---------- 借贷平衡校验 ----------
jf = sum(v['rmb'] for v in vouchers if v['jd'] == '借')
df = sum(v['rmb'] for v in vouchers if v['jd'] == '贷')
print(f'[{YEAR}] 借合计 {jf:,.2f} 贷合计 {df:,.2f} 差 {jf-df:,.2f}')

# 月度汇总 vs 凭证 勾稽
ps_sum = collections.defaultdict(float)
for r in ps_rows:
    ps_sum[(r['Kjqj'], str(r['Kmdm']).strip(), str(r['Jd']).strip())] += float(r['RMB'] or 0)
jz_sum = collections.defaultdict(float)
for v in vouchers:
    jz_sum[(v['kjqj'], v['km'], v['jd'])] += v['rmb']
mism = sum(1 for k in set(ps_sum) | set(jz_sum) if abs(ps_sum.get(k, 0) - jz_sum.get(k, 0)) > 0.01)
print(f'[{YEAR}] P_SUM与凭证不一致组合数: {mism}')

# ---------- 计算账户级发生额/余额 ----------
# 带符号口径: 借方余额为正, 贷方余额为负; 期末(符号) = 期初(符号) + 借方发生 - 贷方发生 (统一公式)
# 余额方向: 优先 Km.KM_YEFX ('1'=借,'2'=贷), 缺失按科目代码首位
def acct_dir(kmdm, km_yefx):
    if km_yefx:
        return '借' if str(km_yefx).strip() == '1' else '贷'
    return '借' if kmdm[0] in '1458' else '贷'

acc = {}
for km in km_rows:
    code = (km['Kmdm'] or '').strip()
    acc[code] = dict(code=code, name=(km['Kmmc'] or '').strip(), level=km['Kmjb'],
                     parent=km['FParentID'], accid=km['FAccountID'],
                     ktype=str(km['KM_TYPE'] or '').strip(),
                     ncye=0.0, jf=0.0, df=0.0, qm=0.0,
                     dirn=acct_dir(code, km.get('KM_YEFX')),
                     monthly=collections.defaultdict(lambda: [0.0, 0.0]))
for r in ye_rows:
    c = (r['Kmdm'] or '').strip()
    if c in acc:
        acc[c]['ncye'] = float(r['Ncye'] or 0)
for v in vouchers:
    c = v['km']
    if c in acc:
        if v['jd'] == '借':
            acc[c]['jf'] += v['rmb']
        elif v['jd'] == '贷':
            acc[c]['df'] += v['rmb']
        m = acc[c]['monthly'][int(v['kjqj'])]
        if v['jd'] == '借':
            m[0] += v['rmb']
        else:
            m[1] += v['rmb']
for c, a in acc.items():
    a['qm'] = a['ncye'] + a['jf'] - a['df']

# 父科目汇总口径（自身+全部下级）: 期初余额只取自身, 发生额取自身+下级合计
# 自底向上(按代码长度降序), 只加直接下级的agg(避免跨级重复计入)
def direct_children(c):
    return [ch for ch in acc if ch != c and ch.startswith(c)
            and ch[len(c):].count('.') == 1]

for c in sorted(acc, key=lambda x: (-len(x), x)):
    a = acc[c]
    a['agg_jf'] = a['jf'] + sum(acc[ch]['agg_jf'] for ch in direct_children(c))
    a['agg_df'] = a['df'] + sum(acc[ch]['agg_df'] for ch in direct_children(c))
    a['agg_qm'] = a['ncye'] + a['agg_jf'] - a['agg_df']

# ---------- 输出JSON ----------
datasets = dict(
    year=YEAR, accounts=acc, ye_rows=ye_rows, vouchers=vouchers,
    ps_rows=[{**r} for r in ps_rows], xm_rows=xm_rows,
    km_rows=km_rows,
)
json_path = os.path.join(OUT, f'{YEAR}_datasets.json')
json.dump(datasets, open(json_path, 'w', encoding='utf-8'), ensure_ascii=False, default=str)
print(f'[{YEAR}] 已写出 {json_path}')

# ---------- 输出: 总账科目余额表 ----------
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

wb = Workbook()
ws = wb.active
ws.title = f'{YEAR}总账余额表'
title_font = Font(name='微软雅黑', size=12, bold=True)
head_font = Font(name='微软雅黑', size=9, bold=True)
cell_font = Font(name='微软雅黑', size=9)
thin = Side(style='thin')
border = Border(left=thin, right=thin, top=thin, bottom=thin)
head_fill = PatternFill('solid', fgColor='D9E1F2')
total_fill = PatternFill('solid', fgColor='FCE4D6')

ws.append([f'{ENTITY} {YEAR}年度 总账科目余额表'])
ws.merge_cells('A1:H1')
ws['A1'].font = title_font
ws.append(['科目代码', '科目名称', '方向', '年初余额',
           '本年借方累计', '本年贷方累计', '期末余额', '期末方向'])
for i, c in enumerate('ABCDEFGH', 1):
    cell = ws.cell(row=2, column=i)
    cell.font = head_font
    cell.fill = head_fill
    cell.border = border
    cell.alignment = Alignment(horizontal='center')

# 按代码排序, 层级显示
def code_key(c):
    parts = c.rstrip('.').split('.')
    return [int(p) if p.isdigit() else 9999 for p in parts]

codes = sorted(acc, key=code_key)
n = 3
for c in codes:
    a = acc[c]
    indent = '　' * (a['level'] - 1) if a['level'] else ''
    ws.cell(row=n, column=1, value=c).font = cell_font
    ws.cell(row=n, column=2, value=indent + a['name']).font = cell_font
    ws.cell(row=n, column=3, value=('借' if a['ncye'] >= 0 else '贷')).font = cell_font
    ws.cell(row=n, column=4, value=round(abs(a['ncye']), 2)).number_format = '#,##0.00'
    ws.cell(row=n, column=5, value=round(a['agg_jf'], 2)).number_format = '#,##0.00'
    ws.cell(row=n, column=6, value=round(a['agg_df'], 2)).number_format = '#,##0.00'
    ws.cell(row=n, column=7, value=round(abs(a['agg_qm']), 2)).number_format = '#,##0.00'
    ws.cell(row=n, column=8, value=('借' if a['agg_qm'] >= 0 else '贷')).font = cell_font
    for col in range(1, 9):
        ws.cell(row=n, column=col).border = border
    n += 1
# 合计行（一级科目汇总）
lvl1 = [a for a in acc.values() if a['level'] == 1]
s_ncye = sum(a['ncye'] for a in lvl1)
s_qm = sum(a['agg_qm'] for a in lvl1)
ws.cell(row=n, column=1, value='合计').font = head_font
ws.cell(row=n, column=3, value=('借' if s_ncye >= 0 else '贷')).font = head_font
ws.cell(row=n, column=4, value=round(abs(s_ncye), 2)).number_format = '#,##0.00'
ws.cell(row=n, column=5, value=round(sum(a['agg_jf'] for a in lvl1), 2)).number_format = '#,##0.00'
ws.cell(row=n, column=6, value=round(sum(a['agg_df'] for a in lvl1), 2)).number_format = '#,##0.00'
ws.cell(row=n, column=7, value=round(abs(s_qm), 2)).number_format = '#,##0.00'
ws.cell(row=n, column=8, value=('借' if s_qm >= 0 else '贷')).font = head_font
for col in range(1, 9):
    cc = ws.cell(row=n, column=col)
    cc.border = border
    cc.fill = total_fill
    cc.font = head_font
ws.column_dimensions['A'].width = 16
ws.column_dimensions['B'].width = 42
for col in 'CDEFGH':
    ws.column_dimensions[col].width = 14
ws.freeze_panes = 'A3'

# ---------- 输出: 明细账 (每个有发生额的一级科目) ----------
def build_mx(km_code):
    a = acc.get(km_code)
    if not a:
        return None
    # 该科目及下级所有明细
    sub = [c for c in acc if c == km_code or c.startswith(km_code)]
    sub.sort(key=code_key)
    rows = []
    for c in sub:
        a2 = acc[c]
        qcye = a2['ncye']
        bal = qcye
        # 期初行
        rows.append(dict(date='', pzlx='', pzbh='', zy='上年结转', dfkm='',
                         jf=0.0, df=0.0, jd=('借' if qcye >= 0 else '贷'),
                         bal=abs(qcye), is_qq=True))
        for v in vouchers:
            if v['km'] != c:
                continue
            if v['jd'] == '借':
                bal += v['rmb']
            else:
                bal -= v['rmb']
            rows.append(dict(date=v['date'], pzlx=v['pzlx'], pzbh=v['pzbh'],
                             zy=v['zy'], dfkm=v['dfkm'],
                             jf=v['rmb'] if v['jd'] == '借' else 0.0,
                             df=v['rmb'] if v['jd'] == '贷' else 0.0,
                             jd=v['jd'],
                             baldir=('借' if bal >= 0 else '贷'), bal=abs(bal)))
        # 期末
        rows.append(dict(date='', pzlx='', pzbh='', zy='期末余额', dfkm='',
                         jf=a2['jf'], df=a2['df'],
                         jd=('借' if a2['qm'] >= 0 else '贷'), bal=abs(a2['qm']), is_qm=True))
    return rows

# 试点明细账: 库存现金 1001
for code, label in [('1001.', '库存现金'), ('1002.', '银行存款'), ('3101.', None)]:
    if code not in acc:
        continue
    name = acc[code]['name'] or label
    if not any(v['km'].startswith(code) for v in vouchers) and acc[code]['ncye'] == 0:
        continue
    rows = build_mx(code)
    if not rows:
        continue
    ws2 = wb.create_sheet(f'明细账_{code.rstrip(".")}')
    ws2.append([f'某行业协会 {YEAR}年度 明细账 — {name} ({code})'])
    ws2.merge_cells('A1:I1')
    ws2['A1'].font = title_font
    ws2.append(['日期', '凭证字号', '摘要', '对方科目', '借方金额', '贷方金额', '方向', '余额', '经办/制单'])
    for i in range(1, 10):
        c2 = ws2.cell(row=2, column=i)
        c2.font = head_font
        c2.fill = head_fill
        c2.border = border
    n2 = 3
    for r in rows:
        ws2.cell(row=n2, column=1, value=r['date']).font = cell_font
        ws2.cell(row=n2, column=2, value=(str(r['pzlx']) + '-' + str(r['pzbh'])) if r['pzlx'] else '').font = cell_font
        ws2.cell(row=n2, column=3, value=r['zy']).font = cell_font
        ws2.cell(row=n2, column=4, value=r['dfkm']).font = cell_font
        ws2.cell(row=n2, column=5, value=round(r['jf'], 2) if r['jf'] else None).number_format = '#,##0.00'
        ws2.cell(row=n2, column=6, value=round(r['df'], 2) if r['df'] else None).number_format = '#,##0.00'
        ws2.cell(row=n2, column=7, value=r.get('baldir', r['jd'])).font = cell_font
        ws2.cell(row=n2, column=8, value=round(r['bal'], 2)).number_format = '#,##0.00'
        for col in range(1, 10):
            ws2.cell(row=n2, column=col).border = border
        if r.get('is_qq') or r.get('is_qm'):
            for col in range(1, 10):
                ws2.cell(row=n2, column=col).fill = total_fill
        n2 += 1
    for col, w in zip('ABCDEFGHI', [12, 10, 40, 12, 14, 14, 6, 14, 10]):
        ws2.column_dimensions[col].width = w
    ws2.freeze_panes = 'A3'

xlsx_path = os.path.join(OUT, f'{YEAR}_总账科目余额表.xlsx')
wb.save(xlsx_path)
print(f'[{YEAR}] 已写出 {xlsx_path}')

# ---------- 控制台摘要 ----------
print(f'[{YEAR}] 科目数 {len(acc)}')
hbg = [a for a in acc.values() if a['level'] == 1]
print(f'[{YEAR}] 一级科目 {len(hbg)} 个, 期末余额不为0的 {sum(1 for a in hbg if abs(a["qm"])>0.005)} 个')

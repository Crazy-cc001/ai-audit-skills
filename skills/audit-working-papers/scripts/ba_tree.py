# -*- coding: utf-8 -*-
"""按业务活动表(ba)层级重建业务活动成本/管理费用审定表明细
1) 解析各年ba: 递归(子项金额加总=父项)重建树
2) 业务活动成本: 37二级(原型值保留)下按2025 ba树插入明细, 逐年按名取值
3) 管理费用: 清除错误明细行, 恢复原型27项
"""
import os, re, json
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font

BASE = r'C:\path\to\temp\xsj_pt'
TPL = r'C:\path\to\底稿模板.xlsx'
BAPATH = r'C:\path\to\协会财务报表\16-25年业务活动表'
YEARS = list(range(2016, 2026))

# ---------- 1. 解析各年业务活动表 ----------
def parse_ba(year):
    import openpyxl as xl
    p1 = os.path.join(BAPATH, f'{year}业务活动表1.xlsx')
    p2 = os.path.join(BAPATH, f'{year}业务活动表.xlsx')
    wb = xl.load_workbook(p1 if os.path.exists(p1) else p2, data_only=True)
    ws = wb['表页_1']
    rows = []
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True):
        if not row or row[0] is None:
            continue
        name = str(row[0]).strip()
        if not name:
            continue
        v = row[5]   # 本年数合计
        val = None if v is None or v == '' else round(float(v), 2)
        rows.append((name, val))
    wb.close()
    # 切分节
    sections = {}
    cur = None
    for name, val in rows:
        if name.startswith(('（', '一、', '二、', '三、', '四、')):
            cur = name
            sections[cur] = []
            continue
        if cur and name != '其中：':
            sections.setdefault(cur, []).append((name, val))
    return sections, rows

def build_tree(items):
    """items: [(name,val)] 递归建树(记忆化, 每位置只解析一次)"""
    memo = {}

    def parse(idx):
        if idx in memo:
            return memo[idx]
        name, V = items[idx]
        if V is None:
            memo[idx] = ({'n': name, 'v': V, 'c': []}, idx + 1)
            return memo[idx]
        j = idx + 1
        children = []
        total = 0.0
        while j < len(items):
            child, j2 = parse(j)
            j = j2
            children.append(child)
            if child['v'] is not None:
                total += child['v']
            if abs(total - V) < 0.005:
                break
            if total > V + 0.005:
                # 不构成父项: 自身为叶(子项作废, 从idx+1重新开始)
                memo[idx] = ({'n': name, 'v': V, 'c': []}, idx + 1)
                return memo[idx]
        memo[idx] = ({'n': name, 'v': V, 'c': children}, j)
        return memo[idx]

    out = []
    i = 0
    while i < len(items):
        node, i = parse(i)
        out.append(node)
    return out

ba_trees = {}
ba_flat = {}
for y in YEARS:
    try:
        secs, rows = parse_ba(y)
    except Exception as e:
        print(f'{y} 解析失败: {e}')
        continue
    trees = {}
    for sec in ('（一）业务活动成本', '（二）管理费用'):
        if sec in secs:
            trees[sec] = build_tree(secs[sec])
    ba_trees[y] = trees
    flat = {}
    for sec, items in secs.items():
        for n, v in items:
            if v is not None:
                flat[n] = v
    ba_flat[y] = flat
    n_bc = len(trees.get('（一）业务活动成本', []))
    n_gl = len(trees.get('（二）管理费用', []))
    print(f'{y}: 业务活动成本顶级{n_bc}项, 管理费用顶级{n_gl}项')

# ---------- 2. 映射与写入 ----------
SYN = {  # ← 换客户时调整: 底稿二级分类名 → 业务活动表明细名
    '维权戒惩工作': '纪律维权工作经费',
    '律师管理与服务经费': '服务与管理经费',
    '党团建设': '党建工作经费',
    '上缴上级协会会费': '交纳上级行业协会会费',
    '上缴社会组织联合会会费': '交纳省社会组织联合会会费',
    '下拨省协会县联络组经费': '下拨县域律师联络组经费',
    '青年律师人才培养基金': '青年律师人才培养经费',
    '县域律师人才培养基金': '县域律师人才培养经费',
    '涉外律师培养基金': '涉外律师人才培养经费',
    '执业责任保险': '全省律师“一揽子”保险',
    '其他不可预见性支出': '其他不可预见支出',
    '"1+1"法律援助及小岗村法律服务活动': '"1+1"法律援助及小岗村公益法律服务',
    '安徽律师行业发展陈列室': '安徽律师行业发展陈列室维护费',
    '午餐补贴': '误餐补贴',
    '绩效考核、奖金及专项补贴': '绩效考核、奖金及专项表彰奖等',
}

def find_node(tree, lvl2):
    for node in tree:
        if node['n'] == lvl2 or node['n'] == SYN.get(lvl2):
            return node
    return None

def flatten(node, depth, out):
    for ch in node['c']:
        out.append((depth + 1, ch['n']))
        flatten(ch, depth + 1, out)

bold = Font(bold=True)

def setv(ws, row, col, v, fmt='#,##0.00'):
    if v is None:
        return
    try:
        c = ws.cell(row=row, column=col)
        c.value = round(v, 2)
        c.number_format = fmt
    except (AttributeError, TypeError):
        pass

wb = openpyxl.load_workbook(TPL)

def rebuild_sheet(sheet, sec, keep_rows):
    """重建: 二级行按原型值, 明细按2025 ba树, 合计=二级和"""
    ws = wb[sheet]
    tree = ba_trees[2025][sec]
    # 读取现有二级行(序号列有值, B不以空格开头): 保存名称与年列值
    lvl2 = []
    r = 7
    while r < 300:
        a = ws.cell(row=r, column=1).value
        b = ws.cell(row=r, column=2).value
        if b and '合计' in str(b):
            break
        if a is not None and b:
            vals = []
            for c in range(3, 13):
                v = ws.cell(row=r, column=c).value
                if isinstance(v, str):
                    v = None if v == '-' else v
                vals.append(v)
            lvl2.append((r, str(b).strip(), vals))
        r += 1
    # 清空数据区
    tr = r
    for rr in range(7, tr):
        for c in range(1, 15):
            try:
                ws.cell(row=rr, column=c).value = None
            except AttributeError:
                pass
    # 重写: 二级(保留原型年列值) + 明细
    lvl2_cells = []
    rr = 7
    for (_, name, vals) in lvl2:
        ws.cell(row=rr, column=2, value=name).font = bold
        for i, v in enumerate(vals):
            setv(ws, rr, 3 + i, v)
        lvl2_cells.append(rr)
        rr += 1
        node = find_node(tree, name)
        if node and node['c']:
            details = []
            flatten(node, 0, details)
            for depth, dn in details:
                ws.cell(row=rr, column=2, value='　' * depth + dn)
                for y in YEARS:
                    setv(ws, rr, 3 + (y - 2016), ba_flat.get(y, {}).get(dn))
                rr += 1
    # 合计
    ws.cell(row=rr, column=2, value='合计').font = bold
    for i in range(10):
        L = get_column_letter(3 + i)
        refs = '+'.join(f'{L}{x}' for x in lvl2_cells)
        ws.cell(row=rr, column=3 + i, value=f'={refs}')
    ws.cell(row=rr, column=13, value=f'=SUM(M7:M{rr-1})')
    ws.cell(row=rr, column=14, value=f'=SUM(N7:N{rr-1})')
    return rr

tr1 = rebuild_sheet('业务活动成本审定表', '（一）业务活动成本', 37)
print(f'业务活动成本审定表: 合计行{tr1}')
tr2 = rebuild_sheet('管理费用审定表', '（二）管理费用', 27)
print(f'管理费用审定表: 合计行{tr2}')

wb.save(TPL)
print('明细已按业务活动表层级重建')

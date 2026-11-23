# -*- coding: utf-8 -*-
"""损益类勾稽: 账套损益科目发生额 vs 业务活动表本年数 (2016-2025)"""
import json, os

BASE = r'C:\path\to\temp\xsj_pt'
sts = json.load(open(os.path.join(BASE, 'statements.json'), encoding='utf-8'))

# 报表项 → 账套科目(费用取借方累计=agg_jf, 收入取贷方累计=agg_df)
MAP = [
    ('会费收入', '4201.', 'df'),
    ('其他收入', '4901.', 'df'),
    ('捐赠收入', '4101.', 'df'),
    ('提供服务收入', '4301.', 'df'),
    ('商品销售收入', '4401.', 'df'),
    ('政府补助收入', '4501.', 'df'),
    ('投资收益', '4601.', 'df'),
    ('（一）业务活动成本', '5101.', 'jf'),
    ('（二）管理费用', '5201.', 'jf'),
    ('（四）筹资费用', '5301.', 'jf'),
    ('（五）其他费用', '5401.', 'jf'),
    ('(三）其他不可预见支出', '5202.', 'jf'),
]

def ba_item(y, item):
    b = sts[str(y)]['ba'].get(item)
    if not b or b[1] is None:
        return None
    return round(b[1], 2)

print(f'{"年度":<6}{"报表项目":<14}{"业务活动表":>14}{"账套科目":>14}{"差额":>12}  结论')
all_diffs = []
for y in range(2016, 2026):
    p = os.path.join(BASE, f'out{y}', f'{y}_datasets.json')
    if not os.path.exists(p):
        continue
    ds = json.load(open(p, encoding='utf-8'))
    acc = ds['accounts']
    for item, code, side in MAP:
        rpt = ba_item(y, item)
        if rpt is None:
            continue
        # 账套: 取该一级科目agg(已含下级)
        a = acc.get(code)
        if a is None:
            continue
        kjm = round(a['agg_jf'] if side == 'jf' else a['agg_df'], 2)
        diff = round(rpt - kjm, 2)
        flag = '相符' if abs(diff) < 0.005 else f'!!差异{diff:,.2f}'
        if abs(diff) >= 0.005:
            all_diffs.append((y, item, rpt, kjm, diff))
            print(f'{y:<6}{item:<14}{rpt:>14,.2f}{kjm:>14,.2f}{diff:>12,.2f}  {flag}')
print()
if all_diffs:
    print(f'共 {len(all_diffs)} 项差异')
else:
    print('损益类全部相符')

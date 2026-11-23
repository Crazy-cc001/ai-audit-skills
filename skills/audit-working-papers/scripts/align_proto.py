# -*- coding: utf-8 -*-
"""按原型对齐底稿: 将原型对应sheet数据区整体复制到当前底稿(保留我的目录/报表/往来2025明细等)"""
import openpyxl
from copy import copy as style_copy

TPL = r'C:\path\to\底稿模板.xlsx'  # 目标底稿路径(按项目修改)
PROTO = r'C:\path\to\原型底稿.xlsx'

# 需要从原型复制的sheet(整体替换)
COPY_SHEETS = [
    '货币资金', '其他应收款审定表.', '固定资产审定表', '预提费用审定表',
    '长期借款审定表', '会费收入审定表', '其他收入审定表', '业务活动成本审定表',
    '其他不可预见支出审定表', '管理费用审定表', '其他费用审定表', '筹资费用审定表',
    '其他应付款审定表',   # 126行十年表
]
# 新增sheet
NEW_SHEETS = ['固定资产明细表']

wb = openpyxl.load_workbook(TPL)
pwb = openpyxl.load_workbook(PROTO)

def copy_sheet(src_ws, dst_ws):
    for row in src_ws.iter_rows():
        for cell in row:
            try:
                nc = dst_ws.cell(row=cell.row, column=cell.column)
                nc.value = cell.value
                if cell.number_format:
                    nc.number_format = cell.number_format
                if cell.font and cell.font.bold:
                    nc.font = openpyxl.styles.Font(bold=True)
                if cell.alignment:
                    nc.alignment = style_copy(cell.alignment)
            except AttributeError:
                pass
    for col, dim in src_ws.column_dimensions.items():
        if dim.width:
            dst_ws.column_dimensions[col].width = dim.width
    # 复制合并
    for m in src_ws.merged_cells.ranges:
        try:
            dst_ws.merge_cells(str(m))
        except Exception:
            pass

for sn in COPY_SHEETS:
    if sn not in pwb.sheetnames or sn not in wb.sheetnames:
        print(f'跳过 {sn}')
        continue
    # 清空现有sheet数据区后复制
    dst = wb[sn]
    for row in dst.iter_rows():
        for cell in row:
            try:
                cell.value = None
            except AttributeError:
                pass
    copy_sheet(pwb[sn], dst)
    print(f'已按原型对齐: {sn}')

for sn in NEW_SHEETS:
    if sn in wb.sheetnames:
        del wb[sn]
    dst = wb.create_sheet(sn)
    copy_sheet(pwb[sn], dst)
    print(f'已新增: {sn}')

pwb.close()
wb.save(TPL)
print('对齐完成')

# -*- coding: utf-8 -*-
"""把表头序列数年份(42735等)改为'YYYY年'命名(参照原型)"""
import openpyxl

TPL = r'C:\path\to\底稿模板.xlsx'

SERIAL_YEAR = {
    42735: '2016年', 43100: '2017年', 43465: '2018年', 43830: '2019年', 44196: '2020年',
    44561: '2021年', 44926: '2022年', 45291: '2023年', 45657: '2024年', 46022: '2025年',
}
SERIAL_ALT = {42370: '2016年初'}

wb = openpyxl.load_workbook(TPL)
fixed = {}
for sn in wb.sheetnames:
    ws = wb[sn]
    cnt = 0
    # 扫描表头区(1-8行)的序列数年份
    for r in range(1, 9):
        for c in range(1, 60):
            cell = ws.cell(row=r, column=c)
            v = cell.value
            if v is None:
                continue
            if isinstance(v, str) and v.strip() in {str(k) for k in SERIAL_YEAR}:
                cell.value = SERIAL_YEAR[int(v.strip())]
                cnt += 1
            elif isinstance(v, str) and v.strip() in {str(k) for k in SERIAL_ALT}:
                cell.value = SERIAL_ALT[int(v.strip())]
                cnt += 1
            elif isinstance(v, (int, float)) and v in SERIAL_YEAR:
                cell.value = SERIAL_YEAR[int(v)]
                cell.number_format = 'General'
                cnt += 1
            elif isinstance(v, (int, float)) and v in SERIAL_ALT:
                cell.value = SERIAL_ALT[int(v)]
                cell.number_format = 'General'
                cnt += 1
    if cnt:
        fixed[sn] = cnt
for sn, n in sorted(fixed.items()):
    print(f'{sn}: {n}处')
wb.save(TPL)
print('年份命名修正完成')

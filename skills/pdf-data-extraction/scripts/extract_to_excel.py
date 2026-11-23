# -*- coding: utf-8 -*-
"""
步骤4: 从Word文档的《投标人基本情况表》中提取关键信息, 汇总为Excel
由用户现有脚本参数化改造而来
"""
import argparse
import os
import re
import sys

import pandas as pd
from docx import Document

COLUMN_ORDER = [
    '文件名', '投标人名称', '注册地址',
    '联系人', '联系人电话', '电子邮件',
    '法定代表人', '法定代表人电话',
    '技术负责人', '技术负责人电话',
]


def _fix_stdout():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


def extract_bidder_info_from_word(docx_path):
    """
    从Word文档的《投标人基本情况表》中提取关键信息
    支持手机号和座机号
    """
    doc = Document(docx_path)
    info = {
        '文件名': os.path.basename(docx_path),
        '投标人名称': None,
        '注册地址': None,
        '联系人': None,
        '联系人电话': None,
        '电子邮件': None,
        '法定代表人': None,
        '法定代表人电话': None,
        '技术负责人': None,
        '技术负责人电话': None,
    }

    # ========== 辅助函数 ==========
    def extract_phone_from_row(row_texts):
        """
        从一行单元格中提取电话号码
        支持: 11位手机号、带区号的座机号(如0551-12345678)
        """
        for text in row_texts:
            # 尝试匹配座机号格式: 区号-号码 或 区号-号码-分机
            # 如 0551-12345678 或 0551-12345678-808
            landline_match = re.search(r'(\d{3,4}-\d{7,8}(?:-\d+)?)', text)
            if landline_match:
                return landline_match.group(1)

            # 尝试匹配11位手机号
            mobile_match = re.search(r'(\d{11})', text)
            if mobile_match:
                return mobile_match.group(1)

            # 上面都没匹配到, 尝试5-8位数字(可能是座机号不含区号)
            phone_match = re.search(r'(\d{5,8})', text)
            if phone_match:
                return phone_match.group(1)

        return None

    def get_cell_value(row_texts, field_name):
        """在行中查找字段名(忽略空格), 跳过空单元格, 返回后面第一个有效值"""
        for i, text in enumerate(row_texts):
            cleaned = re.sub(r'\s+', '', text)
            if field_name in cleaned:
                for j in range(i + 1, len(row_texts)):
                    value = row_texts[j].strip()
                    if not value:
                        continue
                    if not any(f in value for f in ['联系人', '电话', '传真', '电子邮件', '姓名', '技术职称']):
                        return value
                    break
        return None

    def is_address_field(text):
        """
        检查文本是否包含地址相关的字段名
        支持: 注册地址、详细地址、注册地 等
        """
        cleaned = re.sub(r'\s+', '', text)
        for keyword in ('注册地址', '详细地址', '注册地'):
            if keyword in cleaned:
                return True
        return False

    for table in doc.tables:
        table_text = ''
        for row in table.rows:
            for cell in row.cells:
                table_text += cell.text

        if '投标人名称' not in table_text and not is_address_field(table_text):
            continue

        for row in table.rows:
            row_texts = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
            row_combined = ' '.join(row_texts)

            if not row_combined.strip():
                continue

            # ========== 判断行类型 ==========
            has_contact = '联系方式' in row_combined or '联系人' in row_combined
            has_legal = '法定代表人' in row_combined
            has_tech = '技术负责人' in row_combined

            # =========================================================
            # 投标人名称
            # =========================================================
            if '投标人名称' in row_combined and info['投标人名称'] is None:
                for i, text in enumerate(row_texts):
                    if '投标人名称' in text:
                        for j in range(i + 1, len(row_texts)):
                            value = row_texts[j].strip()
                            if not value:
                                continue
                            if value != '投标人名称':
                                info['投标人名称'] = value
                                break
                        break

            # =========================================================
            # 注册地址(支持: 注册地址、详细地址、注册地)
            # =========================================================
            if is_address_field(row_combined) and info['注册地址'] is None:
                for i, text in enumerate(row_texts):
                    if is_address_field(text):
                        for j in range(i + 1, len(row_texts)):
                            value = row_texts[j].strip()
                            if not value:
                                continue
                            if value == '邮政编码':  # 纯标签单元格, 跳过
                                continue
                            # 过滤掉"邮政编码"等无关内容
                            if '邮政编码' in value:
                                value = value.split('邮政编码')[0].strip()
                            if value:
                                info['注册地址'] = value
                            break
                        break

                # 如果上面的方式没取到, 尝试从整行正则提取
                if info['注册地址'] is None:
                    # 匹配 "注册地址/详细地址/注册地" 后面的内容
                    match = re.search(r'(?:注册地址|详细地址|注册地)\s*[：:]\s*(.+)', row_combined)
                    if match:
                        value = match.group(1).strip()
                        if '邮政编码' in value:
                            value = value.split('邮政编码')[0].strip()
                        if value and '邮政编码' not in value:
                            info['注册地址'] = value

            # =========================================================
            # 联系方式行(联系人 + 联系人电话 + 电子邮件)
            # =========================================================
            if has_contact and not has_legal and not has_tech:
                # 提取联系人
                if info['联系人'] is None:
                    value = get_cell_value(row_texts, '联系人')
                    if value:
                        info['联系人'] = value

                # 提取联系人电话
                if info['联系人电话'] is None:
                    phone = extract_phone_from_row(row_texts)
                    if phone:
                        info['联系人电话'] = phone

                # 提取电子邮件
                if info['电子邮件'] is None:
                    value = get_cell_value(row_texts, '电子邮件')
                    if value and '@' in value:
                        info['电子邮件'] = value

            # =========================================================
            # 法定代表人行
            # =========================================================
            if has_legal and info['法定代表人'] is None:
                # 提取法定代表人姓名
                value = get_cell_value(row_texts, '姓名')
                if value and '技术职称' not in value:
                    info['法定代表人'] = value

                # 提取法定代表人电话
                if info['法定代表人电话'] is None:
                    phone = extract_phone_from_row(row_texts)
                    if phone:
                        info['法定代表人电话'] = phone

            # =========================================================
            # 技术负责人行
            # =========================================================
            if has_tech and info['技术负责人'] is None:
                # 提取技术负责人姓名
                value = get_cell_value(row_texts, '姓名')
                if value and '技术职称' not in value:
                    info['技术负责人'] = value

                # 提取技术负责人电话
                if info['技术负责人电话'] is None:
                    phone = extract_phone_from_row(row_texts)
                    if phone:
                        info['技术负责人电话'] = phone

        if info['投标人名称']:
            break

    return info


def run(folder_path, excel_path=None):
    """批量提取并汇总为Excel。返回 (成功数量, 失败文件列表)"""
    if excel_path is None:
        excel_path = os.path.join(folder_path, '投标人信息汇总.xlsx')

    all_results = []
    failed_files = []
    old_doc_files = []

    print("=" * 60)
    print("开始提取信息...")
    print("=" * 60)

    # 先统计总文件数
    total_files = 0
    for filename in os.listdir(folder_path):
        if filename.startswith('~$'):
            continue
        if filename.lower().endswith('.docx'):
            total_files += 1
        elif filename.lower().endswith('.doc'):
            old_doc_files.append(filename)

    if old_doc_files:
        print(f"⚠️ 跳过 {len(old_doc_files)} 个 .doc 旧格式文件(不支持), 请先用Word/WPS另存为.docx")
        for f in old_doc_files:
            print(f"   - {f}")

    print(f"📁 共发现 {total_files} 个Word文档\n")

    processed = 0
    for filename in os.listdir(folder_path):
        if filename.startswith('~$'):
            continue
        if not filename.lower().endswith('.docx'):
            continue
        file_path = os.path.join(folder_path, filename)
        processed += 1

        try:
            result = extract_bidder_info_from_word(file_path)
        except Exception as e:
            failed_files.append(filename)
            print(f"  [{processed}/{total_files}] ❌ {filename} 处理异常: {e}")
            continue

        if result and result['投标人名称']:
            all_results.append(result)
            print(f"  [{processed}/{total_files}] ✅ {result['投标人名称']}")
        else:
            failed_files.append(filename)
            print(f"  [{processed}/{total_files}] ⚠️ {filename} 提取失败")

    # 保存结果
    if all_results:
        df = pd.DataFrame(all_results)
        df = df[COLUMN_ORDER]
        df.to_excel(excel_path, index=False)

        print("\n" + "=" * 60)
        print(f"✅ 提取完成!")
        print(f"   📊 成功提取: {len(all_results)} 个文件")
        if failed_files:
            print(f"   ⚠️ 提取失败: {len(failed_files)} 个文件")
        print(f"   📁 结果已保存至: {excel_path}")
        print("=" * 60)
    else:
        print("\n❌ 未能提取到任何信息")

    return len(all_results), failed_files


def main():
    _fix_stdout()
    parser = argparse.ArgumentParser(description="从Word提取表格信息并汇总为Excel")
    parser.add_argument("--input", required=True, help="包含Word文档的文件夹")
    parser.add_argument("--output", default=None,
                        help="汇总Excel路径(默认: 输入文件夹下 投标人信息汇总.xlsx)")
    args = parser.parse_args()
    run(args.input, args.output)


if __name__ == "__main__":
    main()

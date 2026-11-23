# -*- coding: utf-8 -*-
"""
步骤3: 批量清洗Word文档——拉宽表格 + 去除空格
默认彻底去除所有空格(修复被拆散的数字, 如 "17730225 964" → "17730225964")
"""
import argparse
import os
import re
import sys

from docx import Document
from docx.shared import Inches


def _fix_stdout():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


def remove_all_spaces_from_document(doc):
    """
    去除Word文档中所有空格(包括数字中间的空格)
    彻底清除所有空白字符
    """
    modified = False

    def clean_text(text):
        """彻底去除文本中的所有空格和空白字符"""
        if not text:
            return text
        return re.sub(r'\s+', '', text)

    # 处理所有段落
    for paragraph in doc.paragraphs:
        if paragraph.text:
            if re.search(r'\s', paragraph.text):
                if len(paragraph.runs) > 1:
                    full_text = ''.join(run.text for run in paragraph.runs)
                    cleaned_text = clean_text(full_text)
                    if cleaned_text != full_text:
                        paragraph.runs[0].text = cleaned_text
                        for run in paragraph.runs[1:]:
                            run.text = ''
                        modified = True
                else:
                    original_text = paragraph.text
                    cleaned_text = clean_text(original_text)
                    if cleaned_text != original_text:
                        paragraph.text = cleaned_text
                        modified = True

    # 处理所有表格
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if paragraph.text and re.search(r'\s', paragraph.text):
                        if len(paragraph.runs) > 1:
                            full_text = ''.join(run.text for run in paragraph.runs)
                            cleaned_text = clean_text(full_text)
                            if cleaned_text != full_text:
                                paragraph.runs[0].text = cleaned_text
                                for run in paragraph.runs[1:]:
                                    run.text = ''
                                modified = True
                        else:
                            original_text = paragraph.text
                            cleaned_text = clean_text(original_text)
                            if cleaned_text != original_text:
                                paragraph.text = cleaned_text
                                modified = True

    return modified


def remove_spaces_keep_one_space(doc):
    """去除多余空格, 但保留单词之间的单个空格"""
    modified = False

    def clean_text_keep_one(text):
        """将多个空格合并为一个, 去除首尾空格"""
        if not text:
            return text
        return re.sub(r' +', ' ', text.strip())

    # 处理所有段落
    for paragraph in doc.paragraphs:
        if paragraph.text:
            if len(paragraph.runs) > 1:
                full_text = ''.join(run.text for run in paragraph.runs)
                cleaned_text = clean_text_keep_one(full_text)
                if cleaned_text != full_text:
                    paragraph.runs[0].text = cleaned_text
                    for run in paragraph.runs[1:]:
                        run.text = ''
                    modified = True
            else:
                original_text = paragraph.text
                cleaned_text = clean_text_keep_one(original_text)
                if cleaned_text != original_text:
                    paragraph.text = cleaned_text
                    modified = True

    # 处理所有表格
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if paragraph.text:
                        if len(paragraph.runs) > 1:
                            full_text = ''.join(run.text for run in paragraph.runs)
                            cleaned_text = clean_text_keep_one(full_text)
                            if cleaned_text != full_text:
                                paragraph.runs[0].text = cleaned_text
                                for run in paragraph.runs[1:]:
                                    run.text = ''
                                modified = True
                        else:
                            original_text = paragraph.text
                            cleaned_text = clean_text_keep_one(original_text)
                            if cleaned_text != original_text:
                                paragraph.text = cleaned_text
                                modified = True

    return modified


def widen_tables(doc, col_width_inches=4.0):
    """拉宽文档中的所有表格"""
    modified = False

    for table in doc.tables:
        num_cols = len(table.columns)
        for col_idx in range(num_cols):
            for cell in table.columns[col_idx].cells:
                cell.width = Inches(col_width_inches)
        table.autofit = True
        modified = True

    return modified


def process_document(docx_path, col_width_inches=4.0, remove_all_spaces=True):
    """
    处理单个Word文档
    """
    try:
        doc = Document(docx_path)
        modified = False

        # 1. 拉宽表格
        if widen_tables(doc, col_width_inches):
            modified = True

        # 2. 去除空格
        if remove_all_spaces:
            if remove_all_spaces_from_document(doc):
                modified = True
        else:
            if remove_spaces_keep_one_space(doc):
                modified = True

        if modified:
            doc.save(docx_path)
            return True
        return False

    except Exception as e:
        print(f"  ❌ 处理失败: {str(e)}")
        return False


def run(folder_path, col_width_inches=4.0, remove_all_spaces=True):
    """批量处理文件夹中的所有Word文档。返回 (成功数量, 无需修改数量)"""
    print("=" * 70)
    print("批量处理Word文档")
    print("=" * 70)
    print(f"📋 操作内容:")
    print(f"   ✅ 拉宽表格(每列 {col_width_inches} 英寸)")
    if remove_all_spaces:
        print(f"   ✅ 去除所有空格(包括数字中间的空格)")
    else:
        print(f"   ✅ 去除多余空格(保留单个空格)")
    print("=" * 70)

    word_files = []
    for filename in os.listdir(folder_path):
        if filename.startswith('~$'):
            continue
        if filename.lower().endswith('.docx'):
            word_files.append(filename)

    if len(word_files) == 0:
        print("❌ 未找到任何Word文档")
        return 0, 0

    print(f"\n📁 共发现 {len(word_files)} 个Word文档\n")

    processed = 0
    success_count = 0
    no_change_files = []

    for filename in word_files:
        file_path = os.path.join(folder_path, filename)
        processed += 1

        print(f"[{processed}/{len(word_files)}] 处理: {filename}")

        success = process_document(file_path, col_width_inches, remove_all_spaces)

        if success:
            success_count += 1
            print(f"  ✅ 处理成功")
        else:
            no_change_files.append(filename)
            print(f"  ⚠️ 无需修改")

    print("\n" + "=" * 70)
    print("✅ 批量处理完成!")
    print("=" * 70)
    print(f"📊 处理统计:")
    print(f"   ✅ 成功处理: {success_count} 个文件")
    if no_change_files:
        print(f"   ⚠️ 无需修改: {len(no_change_files)} 个文件")
    print("=" * 70)

    return success_count, len(no_change_files)


def main():
    _fix_stdout()
    parser = argparse.ArgumentParser(description="批量清洗Word文档(拉宽表格 + 去空格, 原地修改)")
    parser.add_argument("--input", required=True, help="包含Word文档的文件夹(原地修改)")
    parser.add_argument("--col-width", type=float, default=4.0, help="表格列宽(英寸), 默认4.0")
    parser.add_argument("--mode", choices=["all", "single"], default="all",
                        help="all=去除所有空格(默认); single=仅合并多余空格, 保留单词间单个空格")
    args = parser.parse_args()
    run(args.input, col_width_inches=args.col_width, remove_all_spaces=(args.mode == "all"))


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""
步骤2: 将提取出的PDF页面批量转换为Word文档(.docx)
"""
import argparse
import os
import sys

from pdf2docx import Converter


def _fix_stdout():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


def convert_pdf_to_word(pdf_path, docx_path):
    """单个PDF转Word"""
    cv = Converter(pdf_path)
    try:
        cv.convert(docx_path)
    finally:
        cv.close()


def run(pdf_folder, output_folder):
    """转换文件夹内所有PDF为Word。返回 (成功数量, 总数)"""
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"📁 已创建输出文件夹: {output_folder}")

    pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith(".pdf")]
    if not pdf_files:
        print("❌ 未找到任何PDF文件")
        return 0, 0

    print("=" * 50)
    print(f"开始转换 {len(pdf_files)} 个PDF为Word...")
    print("=" * 50)

    success = 0
    for idx, filename in enumerate(sorted(pdf_files), 1):
        pdf_path = os.path.join(pdf_folder, filename)
        docx_path = os.path.join(output_folder, filename.replace(".pdf", ".docx"))
        print(f"\n⏳ [{idx}/{len(pdf_files)}] {filename}")
        try:
            convert_pdf_to_word(pdf_path, docx_path)
            success += 1
            print(f"   ✅ 已转换")
        except Exception as e:
            print(f"   ❌ 转换失败: {e}")

    print("\n" + "=" * 50)
    print(f"📊 转换完成! 成功 {success} 个, 失败 {len(pdf_files) - success} 个")
    print(f"   📁 结果保存在: {output_folder}")
    print("=" * 50)
    return success, len(pdf_files)


def main():
    _fix_stdout()
    parser = argparse.ArgumentParser(description="将PDF批量转换为Word")
    parser.add_argument("--input", required=True, help="包含PDF的文件夹")
    parser.add_argument("--output", required=True, help="Word输出文件夹")
    args = parser.parse_args()
    run(args.input, args.output)


if __name__ == "__main__":
    main()

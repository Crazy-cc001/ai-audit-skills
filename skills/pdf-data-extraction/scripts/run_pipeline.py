# -*- coding: utf-8 -*-
"""
一键流水线: 提取PDF关键词页面 → 转Word → 清洗Word → 汇总Excel
"""
import argparse
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import clean_word
import extract_pdf_pages
import extract_to_excel
import fallback_fill
import pdf_to_word

DEFAULT_KEYWORDS = extract_pdf_pages.DEFAULT_KEYWORDS


def _fix_stdout():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


def main():
    _fix_stdout()
    parser = argparse.ArgumentParser(description="PDF信息提取一键流水线")
    parser.add_argument("--pdf-folder", required=True, help="包含PDF的文件夹")
    parser.add_argument("--output", default=None,
                        help="汇总Excel路径(默认: PDF文件夹下 投标人信息汇总.xlsx)")
    parser.add_argument("--work-dir", default=None,
                        help="中间文件目录(默认: 输出Excel同目录下 流水线中间文件)")
    parser.add_argument("--keywords", default=",".join(DEFAULT_KEYWORDS),
                        help="逗号分隔的关键词列表")
    parser.add_argument("--suffix", default="_基本情况表", help="步骤1输出文件名后缀")
    parser.add_argument("--workers", type=int, default=6, help="步骤1并行进程数, 默认6")
    parser.add_argument("--skip-clean", action="store_true", help="跳过步骤3清洗")
    args = parser.parse_args()

    output_xlsx = args.output or os.path.join(args.pdf_folder, "投标人信息汇总.xlsx")
    output_xlsx = os.path.abspath(output_xlsx)
    work_dir = args.work_dir or os.path.join(os.path.dirname(output_xlsx), "流水线中间文件")
    keywords = [k.strip() for k in args.keywords.split(",") if k.strip()]

    p_pages = os.path.join(work_dir, "01_提取页面")
    p_word = os.path.join(work_dir, "02_转Word")

    print("=" * 60)
    print("PDF信息提取流水线")
    print(f"   PDF文件夹: {args.pdf_folder}")
    print(f"   汇总输出 : {output_xlsx}")
    print(f"   中间文件 : {work_dir}")
    print("=" * 60)

    # =========================================================
    # 阶段1: 提取关键词页面
    # =========================================================
    print("\n【阶段 1/4】提取关键词页面")
    success1, total1, not_found1 = extract_pdf_pages.run(args.pdf_folder, p_pages, keywords, args.suffix, args.workers)
    if success1 == 0:
        print("\n❌ 未提取到任何页面, 流水线终止。若PDF为扫描件(无文字层), 请先OCR。")
        sys.exit(1)

    # =========================================================
    # 阶段2: PDF转Word
    # =========================================================
    print("\n【阶段 2/4】PDF转Word")
    success2, total2 = pdf_to_word.run(p_pages, p_word)
    if success2 == 0:
        print("\n❌ 无PDF成功转换为Word, 流水线终止。")
        sys.exit(1)

    # =========================================================
    # 阶段3: 清洗Word(原地修改中间文件)
    # =========================================================
    if not args.skip_clean:
        print("\n【阶段 3/4】清洗Word(拉宽表格 + 去除空格)")
        clean_word.run(p_word, col_width_inches=4.0, remove_all_spaces=True)
    else:
        print("\n【阶段 3/4】已跳过清洗")

    # =========================================================
    # 阶段4: 提取信息汇总Excel
    # =========================================================
    print("\n【阶段 4/4】提取信息汇总Excel")
    success4, failed4 = extract_to_excel.run(p_word, output_xlsx)

    # =========================================================
    # 阶段5(兜底): 对失败投标人, 从源PDF其他位置提取信息
    #   覆盖两类: 步骤1未找到关键词页 / 步骤4提取失败(表格转换丢失、标签不符)
    # =========================================================
    failed_names = list(not_found1)
    for f in failed4:
        name = f.replace('_基本情况表.docx', '').replace('.docx', '')
        if name not in failed_names:
            failed_names.append(name)
    if failed_names:
        print(f"\n【阶段 5/5】兜底提取({len(failed_names)}个失败投标人)")
        fallback_ok = fallback_fill.run(args.pdf_folder, p_word, failed_names)
        if fallback_ok:
            print("\n【阶段 5/5】兜底后重新汇总Excel")
            success4, failed4 = extract_to_excel.run(p_word, output_xlsx)
        else:
            print("\n【阶段 5/5】兜底未生成任何Word文件, 跳过重新汇总")
    else:
        print("\n【阶段 5/5】无失败投标人, 跳过兜底")

    print("\n" + "=" * 60)
    print("✅ 流水线执行完成!")
    print(f"   步骤1 提取页面: {success1}/{total1}")
    print(f"   步骤2 转Word  : {success2}/{total2}")
    print(f"   步骤4 汇总Excel: 成功 {success4} 个")
    if failed4:
        print(f"   剩余失败文件: {', '.join(failed4)}")
    print(f"   📁 汇总结果: {output_xlsx}")
    print(f"   📁 中间文件: {work_dir}")
    print("=" * 60)


if __name__ == "__main__":
    main()

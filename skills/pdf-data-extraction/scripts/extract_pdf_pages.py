# -*- coding: utf-8 -*-
"""
步骤1: 从PDF中提取包含关键字的页面, 保存为新PDF
高速版: PyMuPDF 扫描文本 + 多进程并行
"""
import argparse
import concurrent.futures
import os
import re
import sys

import pymupdf  # PyMuPDF

DEFAULT_KEYWORDS = [
    "投标人基本情况表",
    "投标基本情况表",
    "基本情况表",
    "投标人基本情况",
    "注册地址",
]


def _fix_stdout():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


def _find_keyword_page(doc, keywords, name_hint=None):
    """
    在PDF中查找关键词页
    - 提供 name_hint(文件名中的公司名)时: 优先返回"同时含关键词和公司名"的页,
      避免联合体投标时误取联合体成员的表;
    - 否则返回第一个含关键词的页。
    """
    name_cleaned = re.sub(r'\s+', '', name_hint) if name_hint else None
    fallback_page = None
    for i in range(len(doc)):
        text = doc[i].get_text()
        if not text:
            continue
        text_cleaned = re.sub(r'\s+', '', text)
        hit = False
        for keyword in keywords:
            keyword_cleaned = re.sub(r'\s+', '', keyword)
            if keyword_cleaned in text_cleaned:
                hit = True
                break
        if not hit:
            continue
        if fallback_page is None:
            fallback_page = i
        if name_cleaned and name_cleaned in text_cleaned:
            return i
    return fallback_page


def _extract_one(pdf_path, output_path, keywords):
    """
    单个PDF: 查找关键词页并写出新PDF(子进程内执行)
    返回 (是否成功, 文件名, 错误信息或None)
    """
    try:
        doc = pymupdf.open(pdf_path)
        try:
            name_hint = os.path.splitext(os.path.basename(pdf_path))[0]
            page_num = _find_keyword_page(doc, keywords, name_hint)
            if page_num is None:
                return False, os.path.basename(pdf_path), None

            dst = pymupdf.open()
            dst.insert_pdf(doc, from_page=page_num, to_page=page_num)
            dst.save(output_path)
            dst.close()
            return True, os.path.basename(pdf_path), None
        finally:
            doc.close()
    except Exception as e:
        return False, os.path.basename(pdf_path), str(e)


def _extract_folder_task(folder_path, output_folder, keywords, suffix):
    """
    子文件夹模式任务(子进程内执行): 逐个检查文件夹内PDF, 找到第一个即停
    返回 (文件夹名, 是否成功, 命中的源文件名, PDF总数)
    """
    folder_name = os.path.basename(folder_path)
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]
    output_path = os.path.join(output_folder, f"{folder_name}{suffix}.pdf")
    for pdf_path in pdf_files:
        ok, name, err = _extract_one(os.path.join(folder_path, pdf_path), output_path, keywords)
        if ok:
            return folder_name, True, name, len(pdf_files)
        if err:
            return folder_name, False, f"{name}(出错: {err})", len(pdf_files)
    return folder_name, False, None, len(pdf_files)


def find_pdfs_in_root(root_folder):
    """查找根目录下的PDF文件(不进入子文件夹)"""
    pdf_files = []
    for filename in os.listdir(root_folder):
        file_path = os.path.join(root_folder, filename)
        if os.path.isfile(file_path) and filename.lower().endswith('.pdf'):
            pdf_files.append(file_path)
    return pdf_files


def find_pdfs_by_folder(root_folder, exclude_dir=None):
    """按子文件夹分组查找PDF文件。exclude_dir(输出文件夹)及其子目录不参与扫描。"""
    pdf_files_by_folder = {}
    for dirpath, dirnames, filenames in os.walk(root_folder):
        if dirpath == root_folder:
            continue
        if exclude_dir and os.path.commonpath([dirpath, exclude_dir]) == os.path.abspath(exclude_dir):
            dirnames[:] = []
            continue
        pdf_files = []
        for filename in filenames:
            if filename.lower().endswith('.pdf'):
                pdf_files.append(os.path.join(dirpath, filename))
        if pdf_files:
            pdf_files_by_folder[dirpath] = pdf_files
    return pdf_files_by_folder


def run(pdf_folder, output_folder, keywords=None, suffix="_基本情况表", workers=6):
    """
    执行整步提取。返回 (成功数量, 处理总数, 未找到列表)
    未找到列表: 平铺模式为PDF文件名(不含.pdf), 子文件夹模式为文件夹名
    """
    if keywords is None:
        keywords = DEFAULT_KEYWORDS
    if workers < 1:
        workers = 1

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"📁 已创建输出文件夹: {output_folder}")

    print("=" * 50)
    print("开始扫描PDF文件...")
    print(f"搜索关键词: {', '.join(keywords)}")
    print(f"并行进程数: {workers}")
    print("=" * 50)

    root_pdfs = find_pdfs_in_root(pdf_folder)
    success = 0
    total = 0
    not_found = []

    if root_pdfs:
        # =========================================================
        # 模式一: 根目录有PDF → 逐个提取(并行), 不限制数量
        # =========================================================
        print(f"📁 发现根目录下有 {len(root_pdfs)} 个PDF文件")
        print(f"📄 模式: 逐个提取(并行, 不限制数量)\n")

        total = len(root_pdfs)
        tasks = [
            (pdf_path,
             os.path.join(output_folder,
                          os.path.basename(pdf_path).replace('.pdf', f'{suffix}.pdf')))
            for pdf_path in root_pdfs
        ]

        done = 0
        with concurrent.futures.ProcessPoolExecutor(max_workers=workers) as executor:
            futures = {executor.submit(_extract_one, p, out, keywords): p
                       for p, out in tasks}
            for fut in concurrent.futures.as_completed(futures):
                pdf_path = futures[fut]
                ok, name, err = fut.result()
                done += 1
                if ok:
                    success += 1
                    print(f"   ✅ [{done}/{total}] {name} 已提取")
                elif err:
                    print(f"   ❌ [{done}/{total}] {name} 处理出错: {err}")
                    not_found.append(os.path.basename(pdf_path).replace('.pdf', ''))
                else:
                    print(f"   ⚠️ [{done}/{total}] {name} 未找到目标信息表")
                    not_found.append(os.path.basename(pdf_path).replace('.pdf', ''))
    else:
        # =========================================================
        # 模式二: 根目录没有PDF → 按子文件夹分组(并行), 每个只提取第一个
        # =========================================================
        pdf_files_by_folder = find_pdfs_by_folder(pdf_folder, exclude_dir=output_folder)
        total_folders = len(pdf_files_by_folder)

        if total_folders == 0:
            print("❌ 根目录和子文件夹均未找到任何PDF文件")
            return 0, 0, []

        print(f"📁 根目录下无PDF, 发现 {total_folders} 个子文件夹")
        print(f"📄 模式: 每个子文件夹只提取第一个符合条件的PDF(并行)\n")

        done = 0
        with concurrent.futures.ProcessPoolExecutor(max_workers=workers) as executor:
            futures = {executor.submit(_extract_folder_task, folder_path,
                                       output_folder, keywords, suffix): folder_path
                       for folder_path in pdf_files_by_folder.keys()}
            for fut in concurrent.futures.as_completed(futures):
                folder_name, ok, source_name, pdf_count = fut.result()
                done += 1
                if ok:
                    success += 1
                    print(f"   ✅ [{done}/{total_folders}] {folder_name} 找到! 已提取: {source_name}")
                else:
                    print(f"   ⚠️ [{done}/{total_folders}] {folder_name} 未找到(检查了{pdf_count}个PDF)"
                          + (f", {source_name}" if source_name else ""))
                    not_found.append(folder_name)

        total = total_folders

    print("\n" + "=" * 50)
    print(f"📊 处理完成!")
    print(f"   📁 共处理: {total} 个{'PDF文件' if root_pdfs else '子文件夹'}")
    print(f"   ✅ 成功提取: {success} 个")
    print(f"   ⚠️ 未找到: {total - success} 个")
    print(f"   📁 提取结果保存在: {output_folder}")
    print("=" * 50)

    return success, total, not_found


def main():
    _fix_stdout()
    parser = argparse.ArgumentParser(description="从PDF中提取包含关键词的页面(高速并行版)")
    parser.add_argument("--input", required=True, help="包含PDF的文件夹")
    parser.add_argument("--output", required=True, help="提取页面的输出文件夹")
    parser.add_argument("--keywords", default=",".join(DEFAULT_KEYWORDS),
                        help="逗号分隔的关键词列表")
    parser.add_argument("--suffix", default="_基本情况表", help="输出文件名后缀")
    parser.add_argument("--workers", type=int, default=6, help="并行进程数, 默认6")
    args = parser.parse_args()

    keywords = [k.strip() for k in args.keywords.split(",") if k.strip()]
    run(args.input, args.output, keywords, args.suffix, args.workers)


if __name__ == "__main__":
    main()

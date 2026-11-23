# -*- coding: utf-8 -*-
"""
步骤2-4: 按匹配清单, 从源目录检索复制项目文件 → 逐项目跑PDF提取流水线 → 生成汇总表并输出到指定路径。
- 步骤2: 按项目编号在源目录匹配文件夹并复制到目标目录
- 步骤3: 自动识别PDF布局(根目录平铺/子文件夹平铺/公司文件夹分组/混合), 调用pdf信息提取流水线, 合并结果
- 步骤4: 清理数据 → 按"项目编号-项目名称"重命名 → 加序号前缀复制到汇总目录
"""
import argparse
import json
import os
import re
import shutil
import subprocess
import sys

import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")

DEFAULT_PIPELINE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "..", "..", "pdf-data-extraction", "scripts", "run_pipeline.py",
)


def find_folder_by_code(source_dir, code):
    """源目录中按项目编号(子串)匹配文件夹, 返回文件夹名或None"""
    if not os.path.isdir(source_dir):
        return None
    hits = []
    for d in os.listdir(source_dir):
        p = os.path.join(source_dir, d)
        if os.path.isdir(p) and code in d:
            hits.append(d)
    return hits


ARTIFACT_DIRS = {"流水线中间文件", "输出信息", "01_提取页面", "02_转Word"}


def _is_artifact(dirpath):
    name = os.path.basename(dirpath)
    if name in ARTIFACT_DIRS:
        return True
    return name.startswith("流水线中间文件") or name.startswith("_临时汇总")


def detect_pdf_layout(root):
    """返回 (flat_dirs, company_dirs):
    - flat_dirs: 直接含PDF的目录(平铺模式, 逐个PDF提取);
    - company_dirs: 含≥2个"直接含PDF的子目录"(即≥2个公司文件夹)的目录(分组模式);
    排除历史运行遗留的 流水线中间文件/输出信息 等工件目录。
    公司文件夹目录本身不再作为平铺目录处理(其父目录用分组模式覆盖)。"""
    flat, company = [], []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if not _is_artifact(os.path.join(dirpath, d))]
        if _is_artifact(dirpath):
            continue
        files_pdf = [f for f in filenames if f.lower().endswith(".pdf")]
        if files_pdf:
            flat.append(dirpath)
        pdf_subdirs = []
        for d in dirnames:
            sub = os.path.join(dirpath, d)
            try:
                has = any(f.lower().endswith(".pdf") for f in os.listdir(sub)
                          if os.path.isfile(os.path.join(sub, f)))
            except OSError:
                continue
            if has:
                pdf_subdirs.append(sub)
        if len(pdf_subdirs) >= 2:  # ≥2个公司文件夹才算分组目录
            company.append(dirpath)
    # 公司文件夹本身不再作为平铺目录
    flat = [f for f in flat if os.path.dirname(f) not in company]
    return sorted(set(flat)), sorted(set(company))


def run_pipeline(pdf_folder, output, work_dir, pipeline):
    cmd = [sys.executable, "-X", "utf8", pipeline,
           "--pdf-folder", pdf_folder, "--output", output, "--work-dir", work_dir]
    print(f"   ▶ 流水线: {os.path.relpath(pdf_folder, os.path.dirname(output)) or pdf_folder}", flush=True)
    r = subprocess.run(cmd)
    return r.returncode


def extract_project(proj_root, pipeline):
    """对单个项目文件夹跑提取, 生成 投标人信息汇总.xlsx。返回行数或-1"""
    flat_dirs, company_dirs = detect_pdf_layout(proj_root)
    if not flat_dirs and not company_dirs:
        print(f"   ❌ 未找到任何PDF")
        return -1
    temps, parts = [], []
    i = 0
    for f in flat_dirs:
        parts.append(("flat", f))
    for c in company_dirs:
        parts.append(("company", c))
    for kind, d in parts:
        i += 1
        out = os.path.join(proj_root, f"_临时汇总_{i}.xlsx")
        wd = os.path.join(proj_root, f"流水线中间文件" if i == 1 else f"流水线中间文件_{i}")
        if run_pipeline(d, out, wd, pipeline) == 0 and os.path.exists(out):
            temps.append(out)
        else:
            print(f"   ⚠️ [{kind}] {d} 流水线失败")
    if not temps:
        return -1
    frames = [pd.read_excel(t, dtype=str) for t in temps if os.path.exists(t)]
    for t in temps:
        try:
            os.remove(t)
        except OSError:
            pass
    if not frames:
        return -1
    merged = pd.concat(frames, ignore_index=True)
    if "文件名" in merged.columns:
        merged = merged.drop_duplicates(subset=["文件名"], keep="first")
    final = os.path.join(proj_root, "投标人信息汇总.xlsx")
    merged.to_excel(final, index=False)
    return len(merged)


def finalize_project(proj_root, item, summary_dir, prefix, order_index):
    """清理数据, 重命名为 项目编号-项目名称.xlsx, 加前缀复制到汇总目录"""
    src = os.path.join(proj_root, "投标人信息汇总.xlsx")
    if not os.path.exists(src):
        return None
    df = pd.read_excel(src, dtype=str)
    if "文件名" in df.columns:
        df["文件名"] = df["文件名"].str.replace(r"_基本情况表\.docx$", "", regex=True)
    if "投标人名称" in df.columns:
        df["投标人名称"] = df["投标人名称"].fillna("").str.replace(r"^称?", "", regex=True)
        df["投标人名称"] = df["投标人名称"].str.replace(r"^名称", "", regex=True)
    code = item.get("项目编号", "")
    name = item.get("项目名称", "")
    if not name:
        name = os.path.basename(proj_root)
        name = re.sub(r"^\d+、", "", name)
    base = f"{code}-{name}" if code else name
    base = re.sub(r'[\\/:*?"<>|]', "", base).strip()
    renamed = os.path.join(proj_root, base + ".xlsx")
    df.to_excel(renamed, index=False)
    # 前缀: seq=Excel序号 / folder=文件夹前导数字 / index=顺序号
    if prefix == "folder":
        m = re.match(r"^(\d+)", os.path.basename(proj_root))
        pre = m.group(1).zfill(2) if m else str(order_index).zfill(2)
    elif prefix == "index":
        pre = str(order_index).zfill(2)
    else:
        seq = str(item.get("序号", "")).strip()
        pre = seq.zfill(2) if seq.isdigit() else str(order_index).zfill(2)
    os.makedirs(summary_dir, exist_ok=True)
    dest_file = os.path.join(summary_dir, f"{pre}-{base}.xlsx")
    shutil.copy2(renamed, dest_file)
    return len(df)


def main():
    ap = argparse.ArgumentParser(description="按匹配清单批量复制+提取+汇总")
    ap.add_argument("--list", required=True, help="匹配清单JSON(screen_excel.py的输出)")
    ap.add_argument("--source-dir", required=True, help="源文件根目录")
    ap.add_argument("--dest-dir", required=True, help="项目文件复制目标目录")
    ap.add_argument("--summary-dir", required=True, help="投标人信息汇总表输出目录")
    ap.add_argument("--prefix", default="seq", choices=["seq", "folder", "index"],
                    help="汇总表文件名前缀: seq=Excel序号 / folder=文件夹前导数字 / index=顺序号")
    ap.add_argument("--pipeline", default=DEFAULT_PIPELINE, help="pdf信息提取流水线脚本路径")
    ap.add_argument("--skip-copy", action="store_true")
    ap.add_argument("--skip-extract", action="store_true")
    ap.add_argument("--skip-finalize", action="store_true")
    ap.add_argument("--dry-run", action="store_true", help="只打印计划不执行")
    args = ap.parse_args()

    with open(args.list, encoding="utf-8") as f:
        items = json.load(f)
    print(f"匹配清单共 {len(items)} 条\n")

    # ---------- 步骤2: 复制 ----------
    copied = []
    if not args.skip_copy:
        os.makedirs(args.dest_dir, exist_ok=True)
        print("【步骤2】源目录检索并复制")
        for it in items:
            code = it.get("项目编号", "")
            hits = find_folder_by_code(args.source_dir, code)
            if not hits:
                print(f"   ❌ 序号{it.get('序号')} {code} 未找到匹配文件夹")
                continue
            if len(hits) > 1:
                print(f"   ⚠️ 序号{it.get('序号')} {code} 多个匹配: {hits}, 取第一个")
            folder = hits[0]
            src = os.path.join(args.source_dir, folder)
            dst = os.path.join(args.dest_dir, folder)
            if args.dry_run:
                print(f"   [dry-run] 将复制: {folder}")
            elif os.path.exists(dst):
                print(f"   ⏭️ 已存在: {folder}")
            else:
                print(f"   复制: {folder}")
                shutil.copytree(src, dst)
            copied.append(os.path.join(args.dest_dir, folder))
        print()
    else:
        copied = [os.path.join(args.dest_dir, d) for d in os.listdir(args.dest_dir)
                  if os.path.isdir(os.path.join(args.dest_dir, d))] if os.path.isdir(args.dest_dir) else []

    # ---------- 步骤3: 提取 ----------
    if not args.skip_extract:
        print("【步骤3】逐项目提取投标人信息")
        for it in items:
            code = it.get("项目编号", "")
            proj = next((p for p in copied if code in os.path.basename(p)), None)
            if proj is None:
                print(f"   ⚠️ 序号{it.get('序号')} {code} 无对应目录, 跳过提取")
                continue
            print(f"   项目: {os.path.basename(proj)}")
            if args.dry_run:
                continue
            n = extract_project(proj, args.pipeline)
            print(f"   → 提取 {n} 行" if n >= 0 else "   → 提取失败")
        print()

    # ---------- 步骤4: 汇总输出 ----------
    if not args.skip_finalize:
        print("【步骤4】重命名并输出汇总表")
        total = 0
        for i, it in enumerate(items, 1):
            code = it.get("项目编号", "")
            proj = next((p for p in copied if code in os.path.basename(p)), None)
            if proj is None:
                print(f"   ⚠️ 序号{it.get('序号')} {code} 无对应目录")
                continue
            if args.dry_run:
                print(f"   [dry-run] {os.path.basename(proj)}")
                continue
            n = finalize_project(proj, it, args.summary_dir, args.prefix, i)
            if n is not None:
                total += n
                print(f"   ✅ [{n}行] {os.path.basename(proj)}")
            else:
                print(f"   ⚠️ 无汇总表: {os.path.basename(proj)}")
        print(f"\n✅ 汇总完成: 共 {total} 行 → {args.summary_dir}")


if __name__ == "__main__":
    main()

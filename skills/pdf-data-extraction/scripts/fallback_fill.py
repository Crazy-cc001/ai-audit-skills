# -*- coding: utf-8 -*-
"""
兜底提取: 当投标文件中找不到《投标人基本情况表》(或表格转换失败)时,
从PDF其他位置提取投标人信息, 生成与步骤2产物同构的Word表格, 供步骤4汇总。

数据来源(按优先级):
  1. 表格重建: 源PDF中含"投标人名称/供应商名称 + 注册地址"的表格页,
     用pdfplumber find_tables 按表格线重建(修复pdf2docx丢表格/丢行问题);
  2. 文本解析: 《唱标信息》《法定代表人身份证明书》《授权书》:
     - 投标人名称: 唱标信息"投标人名称：X"/"供应商名称：X",
       或法代证明书"单位名称：X"/"投标人：X"
     - 注册地址: 法代证明书"地 址：X"(截到成立时间/经营期限)
     - 法定代表人+电话: 法代证明书"姓名：X" + "手机号码/联系电话：X"
     - 联系人+电话: 授权书"授权代表联系方式：X"(联系人栏填授权代表)
  解析前去除全部空格, 兼容"手 机 号 码"等拆字写法;
  只取"已填写"的页(含真实电话数字), 忽略空白模板页。

用法:
  python fallback_fill.py --pdf-folder <源PDF文件夹> --word-dir <02_转Word> \
      --failed <失败投标人名1> <失败投标人名2> ...
  或: --auto 自动查找 word-dir 中缺失的投标人(按源文件夹/文件名反查)

返回值: 生成docx数量。调用方(流水线)随后重跑步骤4即可。
"""
import argparse
import os
import re
import sys

from docx import Document

if False:  # 仅类型提示
    import pdfplumber


def _fix_stdout():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


# =========================================================
# 文本解析
# =========================================================

def _strip(t):
    """去除全部空白, 处理'手 机 号 码'等拆字"""
    return re.sub(r'\s+', '', t)


def _find_name(pages):
    """
    投标人名称: 唱标信息/法代证明书, 排除空白模板值(兼容无冒号写法'供应商名称XX公司')。
    捕获用"懒惰+字段名边界"策略: 值截止到下一个字段标签(投标报价/谈判范围等),
    避免无冒号时把后续字段吞进名称; 再校验公司类型后缀。
    """
    tail = r'(公司|集团|中心|事务所|研究院|合作社|银行|学校|院|厂|站|部|社)$'
    boundary = (r'投标报价|报价信息|谈判范围|统一社会|法定代表|注册地址|成立时间|联系人|'
                r'单位性质|地\s*址|电话|传真|电子|单位名称|投标人名称|供应商名称|授权|系|的')
    pat = re.compile(r'(?:投标人名称|供应商名称)[:：]?([^\s：:（）()]{4,40}?)(?=' + boundary + r'|$)')
    # 封面写法: "投标人：XX公司（盖单位章）" / "供应商：XX公司"
    # 边界含\d: 防止把日期("2024年11月15日")吞进名称
    pat_cover = re.compile(
        r'(?:投标人|供应商)[:：]([^\s：:（）()]{4,40}?)(?=\d|盖单位章|盖章|\(|（|$)')
    for t in pages:
        for m in pat_cover.finditer(_strip(t)):
            v = m.group(1)
            if v in ('none', 'None') or '名称' in v or '报价' in v:
                continue
            if not re.search(tail, v):
                continue
            return v
        for m in pat.finditer(_strip(t)):
            v = m.group(1)
            if v in ('none', 'None') or '名称' in v or '报价' in v:
                continue
            if not re.search(tail, v):
                continue
            return v
    return None


def _find_leg_page(pages):
    """找'已填写'的法定代表人身份证明页(含真实电话), 取最后一个"""
    found = None
    for t in pages:
        c = _strip(t)
        if ('法定代表人身份证明' in c or '单位名称' in c) and \
           re.search(r'(?:手机号码|联系电话)[:：](?:\d{3,4}-?\d{7,8}|1\d{10})', c):
            found = c
    return found


def parse_from_text(pages):
    """从文本页解析投标人信息, 返回按步骤4要求排列的行列表"""
    info = {}
    cleaned = [_strip(t) for t in pages]
    info['name'] = _find_name(cleaned)

    leg_page = _find_leg_page(cleaned)
    if leg_page:
        if not info['name']:
            boundary = (r'单位性质|地\s*址|地址|成立时间|经营期限|姓名|性别|年龄|职务|'
                        r'联系|手机|系|盖章|日期|$')
            for m in re.finditer(r'(?:单位名称|投标人)[:：]([^\s：:（）()]{4,40}?)(?=' + boundary + r')', leg_page):
                v = m.group(1)
                if '性质' in v or '盖章' in v or '名称' in v:
                    continue
                if not re.search(r'(公司|集团|中心|事务所|研究院|合作社|银行|学校|院|厂|站|部|社)$', v):
                    continue
                info['name'] = v
                break
        m = re.search(r'地\s*址[:：](.+?)(?:成立时间|经营期限)', leg_page)
        if m:
            addr = m.group(1)
            if addr and addr != '年月日' and '成立' not in addr:
                info['addr'] = addr
        m = re.search(r'姓\s*名[:：]([^：:（）()\d]{2,4})', leg_page)
        if m and m.group(1) not in ('性别',):
            leg = m.group(1)
            # 姓名可能粘连'性别'两字(姓名:狄猛性别:男 → 狄猛性别)
            if leg.endswith('性别'):
                leg = leg[:-2]
            elif leg.endswith('性') and len(leg) > 2:
                leg = leg[:-1]
            info['leg'] = leg
        m = re.search(r'手机号码[:：](1\d{10})', leg_page)
        if m:
            info['legphone'] = m.group(1)
        else:
            m = re.search(r'联系电话[:：](1\d{10}|\d{3,4}-?\d{7,8})', leg_page)
            if m:
                info['legphone'] = m.group(1)

    # 授权书: 联系人栏填授权代表
    for t in cleaned:
        m = re.search(r'授权代表联系方式[:：](1\d{10})', t)
        if m:
            info['repphone'] = m.group(1)
            m2 = re.search(r'授权([^\s（）()，,：:]{2,4})\s*[（(]?供应商授权', t)
            if m2:
                info['rep'] = m2.group(1)
            break

    rows = []
    if info.get('name'):
        rows.append(['投标人名称', info['name']])
    if info.get('addr'):
        rows.append(['注册地址', info['addr']])
    if info.get('rep') or info.get('repphone'):
        rows.append(['联系方式', '联系人', info.get('rep') or '', '', '', '电话', info.get('repphone') or ''])
    if info.get('leg'):
        rows.append(['法定代表人', '姓名', info['leg'], '技术职称', '', '', '电话', info.get('legphone') or ''])
    return rows, info


# =========================================================
# 表格重建
# =========================================================

def _grid_from_page(p, t):
    grid = []
    for row in t.rows:
        cells = []
        for cell in row.cells:
            if cell is None:
                cells.append('')
                continue
            try:
                cells.append((p.within_bbox(cell).extract_text() or '').replace('\n', ''))
            except Exception:
                cells.append('')
        grid.append(cells)
    return grid


def _find_form_grid(pdf_path):
    """在源PDF中找表格页: 含'投标人名称/供应商名称'且含'注册地址', 返回(find_tables网格, 页对象)"""
    import pdfplumber
    try:
        pdf = pdfplumber.open(pdf_path)
    except Exception:
        # 源PDF损坏(如0KB/15字节空文件)时跳过, 交给后续文本解析兜底(其内部已处理)
        return None
    with pdf:
        for p in pdf.pages:
            t = p.extract_text() or ''
            if ('投标人名称' in t or '供应商名称' in t) and '注册地址' in t:
                tables = p.find_tables()
                for tb in tables:
                    head = ''
                    for c in tb.rows[0].cells:
                        if c is None:
                            continue
                        try:
                            head += (p.within_bbox(c).extract_text() or '').replace('\n', '')
                        except Exception:
                            pass
                    head = _strip(head)
                    if ('投标人名称' in head or '供应商名称' in head) and len(tb.rows) >= 5:
                        return _grid_from_page(p, tb)
    return None


def _save_docx(word_dir, fname, grid):
    """保存重建表格; 标签归一化 + 传真/电子邮件行首列补'联系方式'(合并单元格标签)"""
    ncol = max(len(r) for r in grid)
    # 标签归一化: 部分表格用"供应商名称/项目总工", 与步骤4提取规则对齐
    for row in grid:
        for i, c in enumerate(row):
            if c.strip() == '供应商名称':
                row[i] = '投标人名称'
            elif c.strip() == '项目总工':
                row[i] = '技术负责人'
    for row in grid:
        if any('传真' in c or '电子邮件' in c for c in row):
            row[0] = '联系方式'
            break
    doc = Document()
    tb = doc.add_table(rows=len(grid), cols=ncol)
    for ri, row in enumerate(grid):
        for ci in range(ncol):
            tb.rows[ri].cells[ci].text = row[ci] if ci < len(row) else ''
    doc.save(os.path.join(word_dir, fname))
    print(f"   [表格重建] {fname} ({len(grid)}行x{ncol}列)")


def _save_rows(word_dir, fname, rows, info):
    doc = Document()
    tb = doc.add_table(rows=len(rows), cols=8)
    for ri, row in enumerate(rows):
        for ci in range(8):
            tb.rows[ri].cells[ci].text = row[ci] if ci < len(row) else ''
    doc.save(os.path.join(word_dir, fname))
    detail = {k: v for k, v in info.items() if v}
    print(f"   [文本解析] {fname} -> {detail}")


# =========================================================
# 源PDF定位
# =========================================================

def locate_source_pdf(pdf_folder, bidder_name):
    r"""
    按投标人名定位源PDF:
    - 平铺模式: <文件夹>\<名>.pdf (名含'投标文件'时用原名, 否则补'投标文件')
    - 子文件夹模式: 遍历各子文件夹
    返回文件路径或None
    """
    direct = os.path.join(pdf_folder, bidder_name)
    if os.path.isfile(direct):
        return direct
    direct = os.path.join(pdf_folder, bidder_name + '.pdf')
    if os.path.isfile(direct):
        return direct
    # 平铺模式: 名一般形如"XX公司投标文件"
    for f in os.listdir(pdf_folder):
        if f.lower().endswith('.pdf') and f.startswith(bidder_name):
            return os.path.join(pdf_folder, f)
    # 子文件夹模式: 按子文件夹名查找, 取文件夹内第一个PDF的表格/解析
    for dirpath, dirnames, filenames in os.walk(pdf_folder):
        for d in dirnames:
            if d == bidder_name:
                folder = os.path.join(dirpath, d)
                pdfs = [f for f in os.listdir(folder) if f.lower().endswith('.pdf')]
                return [os.path.join(folder, f) for f in pdfs]
    return None


def fill_one(pdf_folder, word_dir, bidder_name):
    """
    处理一个失败投标人。bidder_name为:
    - 平铺模式: "XX公司投标文件"(无.pdf) 或 "XX公司"
    - 子文件夹模式: 文件夹名(即公司名)
    成功生成docx返回True
    """
    target = bidder_name.replace('.pdf', '')
    fname = f'{target}_基本情况表.docx'
    if os.path.exists(os.path.join(word_dir, fname)):
        # 已存在docx但提取失败(表格空/标签不符): 删除后重建
        pass

    src = locate_source_pdf(pdf_folder, target)
    if src is None:
        print(f"   [兜底] 未定位到源PDF: {target}")
        return False
    if isinstance(src, list):
        if not src:
            print(f"   [兜底] 源文件夹无PDF: {target}")
            return False
        # 子文件夹模式: 先逐个找表格页, 再退回文本解析
        import pdfplumber
        all_pages = []
        for sp in src:
            try:
                with pdfplumber.open(sp) as pdf:
                    all_pages.extend([p.extract_text() or '' for p in pdf.pages])
            except Exception:
                continue
            grid = _find_form_grid(sp)
            if grid:
                _save_docx(word_dir, fname, grid)
                return True
        rows, info = parse_from_text(all_pages)
        if rows:
            _save_rows(word_dir, fname, rows, info)
            return True
        print(f"   [兜底] 未能提取任何信息: {target}")
        return False
    else:
        # 平铺模式: 单个PDF
        grid = _find_form_grid(src)
        if grid:
            _save_docx(word_dir, fname, grid)
            return True
        try:
            import pdfplumber
            with pdfplumber.open(src) as pdf:
                pages = [p.extract_text() or '' for p in pdf.pages]
        except Exception as e:
            print(f"   [兜底] 源PDF无法打开(可能损坏): {target}: {str(e)[:60]}")
            return False
        rows, info = parse_from_text(pages)
        if rows:
            _save_rows(word_dir, fname, rows, info)
            return True
        print(f"   [兜底] 未能提取任何信息(可能为空白模板/壳文件): {target}")
        return False


def run(pdf_folder, word_dir, failed_names):
    """
    批量兜底。failed_names: 失败投标人名列表(不带_基本情况表后缀)。
    返回成功生成的docx数量。
    """
    os.makedirs(word_dir, exist_ok=True)
    print("=" * 60)
    print(f"兜底提取: 共 {len(failed_names)} 个失败投标人")
    print(f"   源PDF文件夹: {pdf_folder}")
    print(f"   Word输出    : {word_dir}")
    print("=" * 60)
    ok = 0
    for name in failed_names:
        print(f" 处理: {name}")
        if fill_one(pdf_folder, word_dir, name):
            ok += 1
    print(f"\n📊 兜底完成: 成功重建 {ok}/{len(failed_names)} 个docx")
    return ok


def main():
    _fix_stdout()
    parser = argparse.ArgumentParser(description="无基本情况表时的兜底提取")
    parser.add_argument("--pdf-folder", required=True, help="源PDF文件夹(步骤1的--input)")
    parser.add_argument("--word-dir", required=True, help="步骤3清洗后的Word文件夹(02_转Word)")
    parser.add_argument("--failed", nargs="+", default=None,
                        help="失败投标人名列表(平铺模式为文件名去.pdf, 子文件夹模式为文件夹名)")
    args = parser.parse_args()
    if not args.failed:
        print("请用 --failed 指定失败投标人名")
        sys.exit(1)
    run(args.pdf_folder, args.word_dir, args.failed)


if __name__ == "__main__":
    main()

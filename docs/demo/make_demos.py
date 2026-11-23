# -*- coding: utf-8 -*-
"""
Generate short demo videos (MP4 + GIF) for each skill.

Design notes:
- Frames are rendered with Pillow, encoded to MP4 with imageio-ffmpeg,
  then converted to an optimized GIF with ffmpeg palettegen/paletteuse.
- All numbers in the demos are synthetic and illustrative.
"""
from __future__ import annotations

import argparse
import math
import os
import shutil
import subprocess
from pathlib import Path

import imageio
import imageio_ffmpeg
import numpy as np
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "docs" / "demos"
FONT_UI = r"C:\Windows\Fonts\msyh.ttc"
FONT_BOLD = r"C:\Windows\Fonts\msyhbd.ttc"
FONT_MONO = r"C:\Windows\Fonts\consola.ttf"

W, H = 960, 540
FPS = 12
DURATION = 12.0
GIF_W = 720
GIF_FPS = 10
GIF_COLORS = 96

# ---------------------------------------------------------------------------
# palette
# ---------------------------------------------------------------------------
BG_TOP = (9, 14, 27)
BG_BOT = (17, 26, 46)
TEXT = (229, 231, 235)
MUTED = (148, 163, 184)
DIM = (100, 116, 139)
CARD = (20, 30, 49)
CARD2 = (25, 38, 61)
BORDER = (51, 65, 85)
GREEN = (34, 197, 94)
BLUE = (56, 189, 248)
AMBER = (245, 158, 11)
RED = (239, 68, 68)
WHITE = (248, 250, 252)

# ---------------------------------------------------------------------------
# content
# ---------------------------------------------------------------------------
SKILLS = [
    {
        "slug": "invoice-bank-reconciliation",
        "num": "01",
        "title": "发票 × 银行流水自动对账",
        "subtitle": "销项、进项、流水，一次配平",
        "pain": "月底对账，VLOOKUP 对到眼花",
        "prompt": "帮我把这个月的发票和银行流水对一下，输出六张表",
        "files": [("销项发票", "18 张 · XML/PDF"), ("进项发票", "32 张 · 台账"), ("银行流水", "126 笔 · Excel")],
        "input_headers": ["日期", "对方", "金额", "类型"],
        "input_rows": [
            ["08-03", "客户 A", "128,400.00", "销项"],
            ["08-11", "供应商 B", "56,200.00", "进项"],
            ["08-19", "客户 C", "32,800.00", "销项"],
            ["08-26", "客户 D", "9,600.00", "销项"],
        ],
        "steps": ["解析发票", "清洗流水", "四层匹配", "红冲抵减", "输出六表"],
        "terminal": [
            ("$ 读取 ./data/invoice_data/*.xml", "cmd"),
            ("  解析销项 18 张 / 进项 32 张", "out"),
            ("$ 清洗 ./data/bank_flow/*.xlsx", "cmd"),
            ("  过滤手续费 / 利息 14 笔", "out"),
            ("$ 四层匹配 + 红冲抵减", "cmd"),
            ("  输出 6 张表 → 对账结果.xlsx", "ok"),
        ],
        "output_title": "对账结果 · 六张表",
        "output_headers": ["对象", "金额", "结果", "标记"],
        "output_rows": [
            ["客户 A", "128,400.00", "已核销", "√"],
            ["供应商 B", "56,200.00", "部分付款", "!"],
            ["客户 C", "0.00", "无票流水", "!"],
            ["客户 D", "32,800.00", "已核销", "√"],
        ],
        "metrics": [("92%", "匹配率", GREEN), ("6", "未核销", AMBER), ("9", "无票流水", RED), ("3", "高风险", RED)],
        "note": "合成数据 · 仅演示流程",
    },
    {
        "slug": "bank-statement-reconcile",
        "num": "02",
        "title": "银行对账单 × 凭证核对",
        "subtitle": "先钉死输入，再谈业务",
        "pain": "回单看错一个数，差异查一下午",
        "prompt": "核对这份对账单和凭证，先做三重自校验",
        "files": [("银行对账单", "9 页 · PDF"), ("记账凭证", "14 张 · Excel"), ("上月末余额", "86,420.15",)],
        "input_headers": ["日期", "摘要", "借方", "贷方", "余额"],
        "input_rows": [
            ["08-01", "承前页", "-", "-", "86,420.15"],
            ["08-05", "收款", "-", "128,400.00", "214,820.15"],
            ["08-12", "付款", "56,200.00", "-", "158,620.15"],
            ["08-28", "付款", "38,020.50", "-", "120,599.65"],
        ],
        "steps": ["探文字层", "三重自校验", "逐笔勾稽", "差异清单"],
        "terminal": [
            ("$ python probe_pdf.py 对账单.pdf", "cmd"),
            ("  字符数 1,842 → 有文字层", "out"),
            ("$ python parse_statement.py 对账单.pdf", "cmd"),
            ("  金额合计 √ 余额链 √ 承前页 √", "ok"),
            ("$ 与 14 张凭证逐笔勾稽", "cmd"),
            ("  差异 0 项", "ok"),
        ],
        "output_title": "勾稽结论",
        "output_headers": ["检查项", "结果", "说明", "标记"],
        "output_rows": [
            ["金额合计", "通过", "借 512,300.00", "√"],
            ["余额链", "通过", "18/18 笔", "√"],
            ["承前页", "通过", "86,420.15", "√"],
            ["差异", "无", "0 项", "√"],
        ],
        "metrics": [("18/18", "余额链", GREEN), ("512,300", "借方", BLUE), ("478,120", "贷方", BLUE), ("0", "差异", GREEN)],
        "note": "合成数据 · 仅演示流程",
    },
    {
        "slug": "audit-working-papers",
        "num": "03",
        "title": "审计底稿自动编制",
        "subtitle": "十年账套，一次填完",
        "pain": "逐年导总账，手工填十年审定表",
        "prompt": "按这个底稿模板，把 2016-2025 年的账套数据填进去",
        "files": [("账套数据", "2016-2025"), ("总账 + 明细账", "312 个科目"), ("底稿模板", "审定表十年列")],
        "input_headers": ["年度", "科目", "期初", "借方", "贷方"],
        "input_rows": [
            ["2016", "货币资金", "1,024,300", "2,841,200", "2,756,400"],
            ["2020", "应付款项", "86,200", "512,300", "498,100"],
            ["2025", "净资产", "3,412,800", "1,208,600", "986,400"],
            ["…", "…", "…", "…", "…"],
        ],
        "steps": ["读取账套", "提取年度", "解析报表", "账表核对", "填充底稿"],
        "terminal": [
            ("$ watch_and_capture.py", "cmd"),
            ("  检测年度切换 → 快照 MyTemp0", "out"),
            ("$ extract_year.py 2016..2025", "cmd"),
            ("  科目 312 / 凭证 8,641 条", "out"),
            ("$ recon_year.py && fill_year.py", "cmd"),
            ("  账表核对 28 项 → 底稿已填", "ok"),
        ],
        "output_title": "底稿 · 审定表十年列",
        "output_headers": ["底稿项目", "期间", "状态", "标记"],
        "output_rows": [
            ["货币资金", "10 年", "已填", "√"],
            ["应付款项", "10 年", "已填", "√"],
            ["净资产", "一 / 二 / 三", "已核对", "√"],
            ["差异说明", "3 处", "已标注", "!"],
        ],
        "metrics": [("10", "年", GREEN), ("28", "BS 项", BLUE), ("3", "差异", AMBER), ("0", "手工重算", GREEN)],
        "note": "合成数据 · 需新纪元/K3 与底稿模板",
    },
    {
        "slug": "public-procurement-audit",
        "num": "04",
        "title": "招采项目合规审核",
        "subtitle": "五类对象，逐条对法规",
        "pain": "招标文件、投标文件、评标报告，翻到眼花",
        "prompt": "审核这个项目的招标文件、投标文件和评标报告",
        "files": [("招标文件", "1 份 · PDF/Word"), ("投标文件", "5 家 · 响应文件"), ("评标报告", "1 份 · 评分表")],
        "input_headers": ["检查项", "标准", "实际", "结果"],
        "input_rows": [
            ["投标保证金", "≤ 80 万", "100 万", "违规"],
            ["发售期", "≥ 5 日", "3 日", "违规"],
            ["主观分", "无畸高", "Z = 2.8", "异常"],
            ["文件元数据", "独立编制", "3 家同创建者", "疑似串标"],
        ],
        "steps": ["异常低价", "时间节点", "评分异常", "元数据", "出报告"],
        "terminal": [
            ("$ abnormal_price_check.py 10000000 housing ...", "cmd"),
            ("  ⚠ 8,900,000 低于阈值", "warn"),
            ("$ timeline_check.py events.json", "cmd"),
            ("  ⚠ 发售期 3 日 < 5 日", "warn"),
            ("$ scoring_check.py scores.json", "cmd"),
            ("  ⚠ 评委 Z = 2.8", "warn"),
        ],
        "output_title": "审核报告",
        "output_headers": ["问题", "证据", "依据", "标记"],
        "output_rows": [
            ["异常低价", "8,900,000", "87 号令", "!"],
            ["发售期不足", "3 日", "实施条例", "!"],
            ["评分畸高", "Z = 2.8", "87 号令 64 条", "!"],
            ["元数据雷同", "3 家", "串标迹象", "!"],
        ],
        "metrics": [("3", "严重", RED), ("5", "一般", AMBER), ("12", "法规", BLUE), ("不合规", "结论", RED)],
        "note": "合成数据 · 仅演示流程",
    },
    {
        "slug": "pdf-data-extraction",
        "num": "05",
        "title": "PDF 信息批量提取",
        "subtitle": "一叠投标文件，变成一张表",
        "pain": "一页页翻 PDF，复制粘贴到 Excel",
        "prompt": "从这堆投标 PDF 里提取投标人名称、地址、联系人和电话",
        "files": [("投标 PDF", "36 份 · 1,200+ 页"), ("字段模板", "7 个字段"), ("输出", "投标人信息汇总.xlsx")],
        "input_headers": ["文件", "页数", "关键词", "状态"],
        "input_rows": [
            ["投标人 A.pdf", "186", "基本情况表", "√"],
            ["投标人 B.pdf", "204", "基本情况表", "√"],
            ["投标人 C.pdf", "98", "未找到", "兜底"],
            ["投标人 D.pdf", "312", "基本情况表", "√"],
        ],
        "steps": ["关键词定位", "转 Word", "清洗表格", "提字段", "兜底提取"],
        "terminal": [
            ("$ run_pipeline.py --pdf-folder 投标文件", "cmd"),
            ("  36 份 PDF → 关键词页 36 页", "out"),
            ("  转 Word → 清洗 → 提字段", "out"),
            ("  兜底：唱标信息 / 法代证明书", "out"),
            ("  √ 34 份成功 / 2 份转人工", "ok"),
        ],
        "output_title": "投标人信息汇总.xlsx",
        "output_headers": ["字段", "提取率", "备注", "标记"],
        "output_rows": [
            ["投标人名称", "36/36", "完整", "√"],
            ["注册地址", "35/36", "1 份缺失", "!"],
            ["联系人电话", "34/36", "2 份转人工", "!"],
            ["法定代表人", "36/36", "完整", "√"],
        ],
        "metrics": [("36", "PDF", BLUE), ("7", "字段", GREEN), ("34", "成功", GREEN), ("2", "转人工", AMBER)],
        "note": "合成数据 · 仅演示流程",
    },
    {
        "slug": "excel-project-extraction",
        "num": "06",
        "title": "Excel 项目批量提取",
        "subtitle": "按颜色筛项目，自动找文件",
        "pain": "黄色标记行，一行行找文件夹",
        "prompt": "把黄色和绿色标记的项目都提取出来，汇总并查重",
        "files": [("抽查清单", "1 份 · 颜色标记"), ("源文件目录", "240 个项目"), ("汇总目录", "42 份汇总表")],
        "input_headers": ["序号", "项目编号", "标记", "状态"],
        "input_rows": [
            ["1", "2026AFXGZ0141", "黄", "已匹配"],
            ["2", "2026AFXGZ0142", "绿", "已匹配"],
            ["7", "2026AFXGZ0147", "黄", "已匹配"],
            ["12", "2026AFXGZ0152", "黄", "待查重"],
        ],
        "steps": ["颜色筛选", "编号匹配", "复制资料", "批量提取", "单列查重"],
        "terminal": [
            ("$ screen_excel.py --colors FFFF00,92D050", "cmd"),
            ("  标记行 42 条 → 匹配清单.json", "out"),
            ("$ batch_run.py --list 匹配清单.json", "cmd"),
            ("  复制 42 个项目 → 提取投标人", "out"),
            ("$ dedup_check.py --folder 汇总表", "cmd"),
            ("  重复 3 组 → 查重报告.xlsx", "ok"),
        ],
        "output_title": "汇总 + 查重报告",
        "output_headers": ["步骤", "结果", "数量", "标记"],
        "output_rows": [
            ["项目筛选", "完成", "42 个", "√"],
            ["资料复制", "完成", "42 份", "√"],
            ["投标人汇总", "完成", "42 张", "√"],
            ["查重", "3 组", "已标色", "!"],
        ],
        "metrics": [("42", "项目", GREEN), ("42", "汇总", BLUE), ("3", "重复", AMBER), ("0", "漏项", GREEN)],
        "note": "合成数据 · 仅演示流程",
    },
]

# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
_font_cache: dict[tuple[str, int], ImageFont.FreeTypeFont] = {}


def font(path: str, size: int) -> ImageFont.FreeTypeFont:
    key = (path, size)
    if key not in _font_cache:
        _font_cache[key] = ImageFont.truetype(path, size)
    return _font_cache[key]


def ui(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    return font(FONT_BOLD if bold else FONT_UI, size)


def mono(size: int) -> ImageFont.FreeTypeFont:
    return font(FONT_MONO, size)


def clamp(x: float, lo: float = 0.0, hi: float = 1.0) -> float:
    return max(lo, min(hi, x))


def ease_out(x: float) -> float:
    x = clamp(x)
    return 1 - (1 - x) ** 3


def alpha_for(t: float, start: float, dur: float = 0.35) -> int:
    return int(255 * ease_out((t - start) / dur))


def tw(draw: ImageDraw.ImageDraw, s: str, f: ImageFont.FreeTypeFont) -> int:
    return draw.textbbox((0, 0), s, font=f)[2]


def draw_text(draw, xy, s, f, fill, alpha=255, anchor=None):
    if alpha <= 0:
        return
    if len(fill) == 3:
        fill = (*fill, alpha)
    else:
        fill = (*fill[:3], int(fill[3] * alpha / 255))
    draw.text(xy, s, font=f, fill=fill, anchor=anchor)


def card(draw, box, fill=CARD, outline=BORDER, radius=12, alpha=255, width=1):
    x0, y0, x1, y1 = box
    if x1 <= x0 or y1 <= y0:
        return
    draw.rounded_rectangle(
        [x0, y0, x1, y1],
        radius=radius,
        fill=(*fill, alpha),
        outline=(*outline, min(alpha, 160)),
        width=width,
    )


def pill(draw, x, y, text, f, color, alpha=255, pad_x=10, h=24):
    w = tw(draw, text, f) + pad_x * 2
    draw.rounded_rectangle(
        [x, y, x + w, y + h],
        radius=h // 2,
        fill=(*color, int(alpha * 0.16)),
        outline=(*color, int(alpha * 0.55)),
        width=1,
    )
    draw_text(draw, (x + pad_x, y + (h - f.size) // 2 - 1), text, f, color, alpha)
    return w


def status_color(s: str):
    if s in ("√", "通过", "已核销", "完成", "已匹配", "已填", "已核对", "完整", "已提取"):
        return GREEN
    if s in ("!", "部分付款", "待查重", "已标注", "3 组", "已标色", "1 份缺失", "2 份转人工", "转人工"):
        return AMBER
    if s in ("违规", "异常", "疑似串标", "无票流水", "不合规", "严重"):
        return RED
    return MUTED


def make_background() -> Image.Image:
    img = Image.new("RGB", (W, H), BG_TOP)
    px = img.load()
    for y in range(H):
        k = y / (H - 1)
        r = int(BG_TOP[0] * (1 - k) + BG_BOT[0] * k)
        g = int(BG_TOP[1] * (1 - k) + BG_BOT[1] * k)
        b = int(BG_TOP[2] * (1 - k) + BG_BOT[2] * k)
        for x in range(W):
            px[x, y] = (r, g, b)
    overlay = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)
    # soft glows
    d.ellipse([W - 260, -180, W + 180, 260], fill=(*BLUE, 14))
    d.ellipse([-220, H - 240, 260, H + 160], fill=(*GREEN, 12))
    # grid
    for x in range(0, W, 48):
        d.line([(x, 0), (x, H)], fill=(255, 255, 255, 5), width=1)
    for y in range(0, H, 48):
        d.line([(0, y), (W, y)], fill=(255, 255, 255, 5), width=1)
    img = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
    return img


BG = make_background()


def draw_chrome(draw, skill, t):
    # top brand
    draw_text(draw, (44, 24), "AI AUDIT SKILLS", ui(14, True), WHITE, 235)
    draw_text(draw, (44, 43), "审计与财务 AI 技能库", ui(11), MUTED, 180)
    label = f"{skill['num']} / 06  ·  {skill['slug']}"
    draw_text(draw, (W - 44 - tw(draw, label, ui(11)), 30), label, ui(11), DIM, 200)
    # header rule
    draw.line([(44, 68), (W - 44, 68)], fill=(*BORDER, 90), width=1)
    # footer
    draw.line([(44, H - 42), (W - 44, H - 42)], fill=(*BORDER, 60), width=1)
    draw_text(draw, (44, H - 30), "github.com/Crazy-cc001/ai-audit-skills", mono(11), DIM, 190)
    note = skill["note"]
    draw_text(draw, (W - 44 - tw(draw, note, ui(11)), H - 30), note, ui(11), DIM, 190)
    # progress
    p = clamp(t / DURATION)
    draw.rounded_rectangle([44, H - 12, W - 44, H - 9], radius=2, fill=(*BORDER, 90))
    draw.rounded_rectangle([44, H - 12, 44 + (W - 88) * p, H - 9], radius=2, fill=(*GREEN, 210))


def scene_label(draw, t, text, start, color=GREEN):
    a = alpha_for(t, start, 0.3)
    if a <= 0:
        return
    x, y = 44, 88
    draw.ellipse([x, y + 4, x + 7, y + 11], fill=(*color, a))
    draw_text(draw, (x + 14, y), text, ui(13, True), color, a)


def draw_number(draw, skill, t):
    a = int(255 * clamp((t - 0.1) / 0.6))
    big = font(FONT_BOLD, 250)
    s = skill["num"]
    w = tw(draw, s, big)
    draw_text(draw, (W - 100 - w, 120), s, big, (255, 255, 255), min(18, a // 8))
    draw_text(draw, (W - 100 - w, 120), s, big, BLUE, min(28, a // 6))


def draw_intro(draw, skill, t):
    scene_label(draw, t, "DEMO", 0.05, GREEN)
    a1 = alpha_for(t, 0.15, 0.5)
    a2 = alpha_for(t, 0.35, 0.5)
    a3 = alpha_for(t, 0.55, 0.5)
    y = 150
    draw_text(draw, (44, y), skill["title"], ui(46, True), WHITE, a1)
    draw_text(draw, (44, y + 62), skill["subtitle"], ui(22), MUTED, a2)
    # pain quote
    draw_text(draw, (44, y + 112), "“", ui(44, True), AMBER, a3)
    draw_text(draw, (78, y + 124), skill["pain"], ui(19), TEXT, a3)
    # prompt pill
    a4 = alpha_for(t, 1.0, 0.45)
    if a4 > 0:
        y4 = 402
        card(draw, (44, y4, W - 44, y4 + 52), CARD, BORDER, radius=12, alpha=a4)
        draw_text(draw, (60, y4 + 16), ">", mono(15), GREEN, a4)
        draw_text(draw, (84, y4 + 15), f"你可以这样说：{skill['prompt']}", ui(15), TEXT, a4)
    draw_number(draw, skill, t)


def draw_input(draw, skill, t):
    scene_label(draw, t, "01 · 输入", 2.0, BLUE)
    # file cards
    for i, item in enumerate(skill["files"]):
        a = alpha_for(t, 2.15 + i * 0.18, 0.35)
        if a <= 0:
            continue
        x = 44 - int(18 * (1 - ease_out(clamp((t - 2.15 - i * 0.18) / 0.35))))
        y = 126 + i * 82
        card(draw, (x, y, x + 330, y + 66), CARD, BORDER, radius=12, alpha=a)
        # file icon
        draw.rounded_rectangle([x + 16, y + 16, x + 42, y + 50], radius=4, fill=(*BLUE, int(a * 0.18)), outline=(*BLUE, int(a * 0.5)))
        draw.polygon([(x + 32, y + 16), (x + 42, y + 26), (x + 32, y + 26)], fill=(*BLUE, int(a * 0.35)))
        draw_text(draw, (x + 56, y + 14), item[0], ui(16, True), TEXT, a)
        if len(item) > 1:
            draw_text(draw, (x + 56, y + 38), item[1], ui(12), MUTED, a)

    # input table
    a = alpha_for(t, 2.5, 0.45)
    if a <= 0:
        return
    x0, y0 = 404, 126
    w, row_h = 512, 34
    card(draw, (x0, y0, x0 + w, y0 + 44 + row_h * len(skill["input_rows"]) + 14), CARD, BORDER, radius=12, alpha=a)
    draw_text(draw, (x0 + 16, y0 + 12), "数据预览", ui(13, True), MUTED, a)
    cols = skill["input_headers"]
    cw = [w / len(cols)] * len(cols)
    tx = x0 + 16
    for j, c in enumerate(cols):
        draw_text(draw, (tx, y0 + 40), c, ui(11), DIM, a)
        tx += cw[j]
    for r, row in enumerate(skill["input_rows"]):
        ry = y0 + 62 + r * row_h
        if r % 2 == 0:
            draw.rounded_rectangle([x0 + 8, ry - 4, x0 + w - 8, ry + row_h - 6], radius=5, fill=(255, 255, 255, int(a * 0.025)))
        tx = x0 + 16
        for j, val in enumerate(row):
            color = status_color(val) if j == len(row) - 1 else TEXT
            draw_text(draw, (tx, ry), val, ui(12), color, a)
            tx += cw[j]


def draw_process(draw, skill, t):
    scene_label(draw, t, "02 · 处理", 4.5, GREEN)
    # steps
    for i, s in enumerate(skill["steps"]):
        start = 4.55 + i * 0.52
        a = alpha_for(t, start, 0.3)
        if a <= 0:
            continue
        x, y = 64, 132 + i * 56
        done = t >= start + 0.35
        c = GREEN if done else BORDER
        draw.ellipse([x, y, x + 24, y + 24], outline=(*c, a), width=2)
        draw_text(draw, (x + 7, y + 3), str(i + 1), ui(12, True), c, a)
        draw_text(draw, (x + 38, y + 2), s, ui(15), TEXT if done else MUTED, a)
        if done:
            # check mark
            draw.line([(x + 7, y + 12), (x + 11, y + 17), (x + 18, y + 7)], fill=(*GREEN, a), width=2)
        if i < len(skill["steps"]) - 1:
            draw.line([(x + 12, y + 26), (x + 12, y + 54)], fill=(*BORDER, int(a * 0.7)), width=1)

    # terminal card
    a = alpha_for(t, 4.7, 0.45)
    if a <= 0:
        return
    x0, y0, x1, y1 = 372, 122, W - 44, 430
    card(draw, (x0, y0, x1, y1), (12, 19, 33), BORDER, radius=12, alpha=a)
    draw.rounded_rectangle([x0, y0, x1, y0 + 30], radius=12, fill=(255, 255, 255, int(a * 0.035)))
    for i, c in enumerate([(239, 68, 68), (245, 158, 11), (34, 197, 94)]):
        draw.ellipse([x0 + 14 + i * 16, y0 + 11, x0 + 22 + i * 16, y0 + 19], fill=(*c, a))
    draw_text(draw, (x0 + 74, y0 + 7), "terminal", mono(11), DIM, a)

    lines = skill["terminal"]
    visible = int(clamp((t - 4.9) / 3.3) * len(lines) * 1.6)
    y = y0 + 48
    for i, (text, kind) in enumerate(lines):
        if i > visible:
            break
        # typewriter for the last visible line
        if i == visible:
            frac = clamp(visible - i)
            text = text[: max(1, int(len(text) * frac))]
        color = GREEN if kind == "ok" else AMBER if kind == "warn" else TEXT if kind == "cmd" else MUTED
        f = mono(13)
        if any("\u4e00" <= ch <= "\u9fff" for ch in text):
            f = ui(13)
        draw_text(draw, (x0 + 18, y), text, f, color, a)
        y += 25


def draw_output(draw, skill, t):
    scene_label(draw, t, "03 · 输出", 8.4, AMBER)
    # table
    a = alpha_for(t, 8.5, 0.4)
    if a <= 0:
        return
    x0, y0 = 44, 126
    w = 610
    row_h = 38
    card(draw, (x0, y0, x0 + w, y0 + 48 + row_h * len(skill["output_rows"]) + 12), CARD, BORDER, radius=12, alpha=a)
    draw_text(draw, (x0 + 18, y0 + 14), skill["output_title"], ui(14, True), WHITE, a)
    cols = skill["output_headers"]
    cw = [w * 0.30, w * 0.30, w * 0.28, w * 0.12]
    tx = x0 + 18
    for j, c in enumerate(cols):
        draw_text(draw, (tx, y0 + 44), c, ui(11), DIM, a)
        tx += cw[j]
    for r, row in enumerate(skill["output_rows"]):
        ra = alpha_for(t, 8.65 + r * 0.28, 0.3)
        if ra <= 0:
            continue
        ry = y0 + 66 + r * row_h
        if r % 2 == 0:
            draw.rounded_rectangle([x0 + 8, ry - 5, x0 + w - 8, ry + row_h - 7], radius=5, fill=(255, 255, 255, int(ra * 0.025)))
        tx = x0 + 18
        for j, val in enumerate(row):
            if j == len(row) - 1:
                color = status_color(val)
                pill(draw, tx - 4, ry - 2, val, ui(11, True), color, ra, pad_x=8, h=22)
            else:
                draw_text(draw, (tx, ry), val, ui(12), TEXT, ra)
            tx += cw[j]

    # metric cards
    mx = 680
    for i, (value, label, color) in enumerate(skill["metrics"]):
        ma = alpha_for(t, 8.9 + i * 0.14, 0.3)
        if ma <= 0:
            continue
        col = i % 2
        row = i // 2
        x = mx + col * 128
        y = 126 + row * 112
        card(draw, (x, y, x + 116, y + 98), CARD, BORDER, radius=10, alpha=ma)
        draw_text(draw, (x + 12, y + 14), value, ui(24, True), color, ma)
        draw_text(draw, (x + 12, y + 54), label, ui(11), MUTED, ma)


def draw_outro(draw, skill, t):
    a = alpha_for(t, 11.0, 0.4)
    if a <= 0:
        return
    # star
    cx, cy, r = W // 2, 150, 34
    pts = []
    for i in range(10):
        ang = -math.pi / 2 + i * math.pi / 5
        rr = r if i % 2 == 0 else r * 0.45
        pts.append((cx + rr * math.cos(ang), cy + rr * math.sin(ang)))
    draw.polygon(pts, fill=(*AMBER, a))
    draw_text(draw, (W // 2, 210), "把重复工作交给 AI", ui(38, True), WHITE, a, anchor="ma")
    draw_text(draw, (W // 2, 266), "6 个技能 · 全部开源 · MIT", ui(16), MUTED, a, anchor="ma")
    url = "github.com/Crazy-cc001/ai-audit-skills"
    draw_text(draw, (W // 2, 318), url, mono(20), BLUE, a, anchor="ma")
    draw_text(draw, (W // 2, 360), "如果这些技能帮你省了时间，点个 Star", ui(17), TEXT, a, anchor="ma")


def render(skill, t):
    img = BG.copy()
    draw = ImageDraw.Draw(img, "RGBA")
    draw_chrome(draw, skill, t)
    if t < 2.0:
        draw_intro(draw, skill, t)
    elif t < 4.5:
        draw_input(draw, skill, t)
    elif t < 8.4:
        draw_process(draw, skill, t)
    elif t < 11.0:
        draw_output(draw, skill, t)
    else:
        draw_outro(draw, skill, t)
    # fade in/out
    k = 1.0
    if t < 0.25:
        k = t / 0.25
    if t > DURATION - 0.25:
        k = (DURATION - t) / 0.25
    if k < 1.0:
        black = Image.new("RGB", (W, H), (0, 0, 0))
        img = Image.blend(black, img, clamp(k))
    return img


def encode_mp4(skill, frames_iter, path: Path):
    writer = imageio.get_writer(
        str(path),
        fps=FPS,
        codec="libx264",
        quality=8,
        macro_block_size=None,
        output_params=["-pix_fmt", "yuv420p", "-movflags", "+faststart"],
    )
    for frame in frames_iter:
        writer.append_data(np.asarray(frame))
    writer.close()


def make_gif(mp4: Path, gif: Path):
    ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()
    vf = (
        f"fps={GIF_FPS},scale={GIF_W}:-1:flags=lanczos,"
        "split[s0][s1];"
        f"[s0]palettegen=max_colors={GIF_COLORS}:stats_mode=diff[p];"
        "[s1][p]paletteuse=dither=bayer:bayer_scale=3"
    )
    subprocess.run(
        [ffmpeg, "-y", "-i", str(mp4), "-vf", vf, "-loop", "0", str(gif)],
        check=True,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def generate(skill, skip_mp4=False, skip_gif=False):
    OUT.mkdir(parents=True, exist_ok=True)
    mp4 = OUT / f"{skill['slug']}.mp4"
    gif = OUT / f"{skill['slug']}.gif"
    cover = OUT / f"{skill['slug']}.png"
    if not skip_mp4:
        n = int(DURATION * FPS)
        frames = (render(skill, i / FPS) for i in range(n))
        encode_mp4(skill, frames, mp4)
    if not skip_gif:
        make_gif(mp4, gif)
    # cover: output scene, where the value is most visible
    render(skill, 9.6).save(cover)
    print(f"{skill['slug']}: mp4={mp4.stat().st_size/1e6:.2f}MB gif={gif.stat().st_size/1e6:.2f}MB")


def preview(selected):
    p = OUT / "preview"
    p.mkdir(parents=True, exist_ok=True)
    for skill in selected:
        for t in (0.8, 3.2, 6.4, 9.8, 11.5):
            render(skill, t).save(p / f"{skill['slug']}-{t:.1f}.png")
    print(f"previews -> {p}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--only", default="", help="comma-separated slugs")
    ap.add_argument("--preview", action="store_true")
    ap.add_argument("--skip-mp4", action="store_true")
    ap.add_argument("--skip-gif", action="store_true")
    args = ap.parse_args()
    selected = SKILLS
    if args.only:
        wanted = {s.strip() for s in args.only.split(",") if s.strip()}
        selected = [s for s in SKILLS if s["slug"] in wanted]
    if args.preview:
        preview(selected)
        return
    for s in selected:
        generate(s, skip_mp4=args.skip_mp4, skip_gif=args.skip_gif)


if __name__ == "__main__":
    main()

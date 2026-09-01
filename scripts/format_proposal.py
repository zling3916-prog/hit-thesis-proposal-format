#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
hit-thesis-proposal-format: 哈尔滨工业大学硕士学位开题报告格式调整脚本

按照《硕士学位论文报告模板（2025）》的格式要求，把一份开题报告草稿 docx
统一为规范格式（页面设置 / 字体 / 层级标题 / 段落与行距 / 参考文献）。

用法:
    python format_proposal.py <输入.docx> <输出.docx> [--cover-json 封面元数据.json] [--dry-run]

- 输入：任意内容完整的开题报告草稿（正文 + 层级标题 + 参考文献，可有封面）。
- 输出：套用模板格式后的新 docx；原文件不会被修改。
- --cover-json：可选。提供 JSON 时在文首重建模板封面，字段：
      {"title":"...","degree":"学位论文|实践成果","college":"...","discipline":"...",
       "advisor":"...","student":"...","student_id":"...","date":"..."}
- 仅改格式，不增删正文文字。

依赖: python-docx  (pip install python-docx)
"""
import argparse
import json
import os
import re
import sys

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

# ---------- 常量：模板格式参数（与 references/format-spec.md 保持一致） ----------
PAGE = dict(width=Cm(21.0), height=Cm(29.7),
            top=Cm(2.5), bottom=Cm(2.2), left=Cm(2.5), right=Cm(2.5),
            header=Cm(1.5), footer=Cm(1.75))

LINE_TWIPS = 420          # 固定行距 21pt（w:line=420, lineRule=exact）

# 层级规则: 中文字体 / 字号pt / 段前行 / 段后行 / 顶格与否
LEVELS = {
    1: dict(east='黑体', size=15, before_lines=0.5, after_lines=0.5, indent='none'),   # 节
    2: dict(east='黑体', size=14, before_lines=0.5, after_lines=0.5, indent='none'),   # 条
    3: dict(east='黑体', size=12, before_lines=0,   after_lines=0,   indent='none'),   # 款
    4: dict(east='黑体', size=12, before_lines=0,   after_lines=0,   indent='item'),   # 项
}
BODY = dict(east='宋体', size=12, before_lines=0, after_lines=0, indent='body')

LATIN_FONT = 'Times New Roman'
ITEM_LEFT_TWIPS = 480     # 项：空4个半角字符 ≈ 0.85cm（12pt 下 4×6pt=24pt=480缇）
BODY_FIRST_TWIPS = 480    # 正文首行缩进2字符 ≈ 0.85cm

# 层级编号识别（同时兼容 1．/1. /1、 与 1.1、1.1.1、（1））
RE_L1 = re.compile(r'^\s*\d+\s*[．.、]\s*\S')
RE_L2 = re.compile(r'^\s*\d+\s*[．.]\s*\d+\s*\S')
RE_L3 = re.compile(r'^\s*\d+\s*[．.]\s*\d+\s*[．.]\s*\d+\s*\S')
RE_L4 = re.compile(r'^\s*[（(]\s*\d+\s*[)）]\s*\S')
# 项判定：短句且不以句读结尾（避免把正文枚举误判为标题）
ITEM_MAX_LEN = 40
ITEM_END_NO = ('。', '，', '；', '、', '：')

# 封面行识别（限文档前 30 段）
COVER_PATTERNS = [
    ('school', re.compile(r'哈尔滨工业大学')),
    ('main_title', re.compile(r'硕士学位开题报告')),
    ('subtitle', re.compile(r'^[（(](学位论文|实践成果)[)）]')),
    ('topic', re.compile(r'^\s*题\s*目\s*[:：]')),
    ('college', re.compile(r'^\s*学\s*院\s*[（(]部[)）]')),
    ('discipline', re.compile(r'学科\s*/?\s*专业学位类别')),
    ('advisor', re.compile(r'^\s*导\s*师')),
    ('student', re.compile(r'^\s*研\s*究\s*生')),
    ('student_id', re.compile(r'^\s*学\s*号')),
    ('date', re.compile(r'开题报告日期')),
    ('foot', re.compile(r'研究生院制')),
]
COVER_MAX_INDEX = 30

REF_HEAD_RE = re.compile(r'^\s*\d+\s*[．.、]?\s*参考文献')


# ---------- 底层工具 ----------
def set_run_font(run, east='宋体', latin=LATIN_FONT, size_pt=None, bold=None, underline=None):
    """设置 run 的中英文字体、字号、加粗、下划线（会清掉旧的 rFonts 相关属性）。"""
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), latin)
    rFonts.set(qn('w:hAnsi'), latin)
    rFonts.set(qn('w:eastAsia'), east)
    rFonts.set(qn('w:cs'), latin)
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.font.bold = bold
    if underline is not None:
        run.font.underline = underline


def set_para_spacing(p, line=LINE_TWIPS, rule='exact', before_lines=None, after_lines=None,
                     before_pt=None, after_pt=None):
    """设置段落的行距与段前/段后。before_lines/after_lines 单位是"行"(0.5=半行)。"""
    pPr = p._element.get_or_add_pPr()
    sp = pPr.find(qn('w:spacing'))
    if sp is None:
        sp = OxmlElement('w:spacing')
        pPr.append(sp)
    if line is not None:
        sp.set(qn('w:line'), str(int(line)))
        sp.set(qn('w:lineRule'), rule)
    # 清除旧的 before/after/行单位，再按需设置
    for attr in ('w:before', 'w:after', 'w:beforeLines', 'w:afterLines'):
        sp.attrib.pop(qn(attr), None)
    if before_lines is not None:
        sp.set(qn('w:beforeLines'), str(int(round(before_lines * 100))))
        sp.set(qn('w:before'), str(int(round(before_lines * 240))))
    if after_lines is not None:
        sp.set(qn('w:afterLines'), str(int(round(after_lines * 100))))
        sp.set(qn('w:after'), str(int(round(after_lines * 240))))
    if before_pt is not None:
        sp.set(qn('w:before'), str(int(before_pt * 20)))
    if after_pt is not None:
        sp.set(qn('w:after'), str(int(after_pt * 20)))


def set_para_indent(p, first_line_twips=None, first_line_chars=None,
                    left_twips=None, left_chars=None):
    """设置首行缩进 / 左缩进。传 None 表示清空对应项。"""
    pPr = p._element.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    for attr in ('w:firstLine', 'w:firstLineChars', 'w:left', 'w:leftChars'):
        ind.attrib.pop(qn(attr), None)
    if first_line_twips is not None:
        ind.set(qn('w:firstLine'), str(int(first_line_twips)))
    if first_line_chars is not None:
        ind.set(qn('w:firstLineChars'), str(int(first_line_chars)))
    if left_twips is not None:
        ind.set(qn('w:left'), str(int(left_twips)))
    if left_chars is not None:
        ind.set(qn('w:leftChars'), str(int(left_chars)))


def set_snap_to_grid(p, val='0'):
    pPr = p._element.get_or_add_pPr()
    el = pPr.find(qn('w:snapToGrid'))
    if el is None:
        el = OxmlElement('w:snapToGrid')
        pPr.append(el)
    el.set(qn('w:val'), val)


# ---------- 分类 ----------
def classify(text, style_name):
    """返回 1=节 2=条 3=款 4=项 0=正文 None=封面行。优先用 Word 内置 Heading 样式，再回退正则。"""
    if style_name:
        m = re.search(r'Heading\s*([1-9])', style_name)
        if m:
            return int(m.group(1)) if int(m.group(1)) <= 4 else 3
    t = text.strip()
    if not t:
        return None
    if RE_L3.match(t):
        return 3
    if RE_L2.match(t):
        return 2
    if RE_L1.match(t):
        return 1
    if RE_L4.match(t) and len(t) <= ITEM_MAX_LEN and not t.endswith(ITEM_END_NO):
        return 4
    return 0


def is_cover_line(text):
    for _, pat in COVER_PATTERNS:
        if pat.search(text):
            return True
    return False


# ---------- 应用格式 ----------
def apply_level(p, level, in_ref=False):
    """对单个段落套用格式。level: 1..4 标题, 0 正文, 'cover' 封面。"""
    text = p.text
    # 清空段落级旧格式：对齐、缩进、间距、网格对齐
    pf = p.paragraph_format
    pf.alignment = None
    set_snap_to_grid(p, '0')

    if level == 'cover':
        apply_cover_line(p, text)
        return

    if level in (1, 2, 3, 4):
        spec = LEVELS[level]
        set_para_spacing(p, line=LINE_TWIPS, rule='exact',
                         before_lines=spec['before_lines'], after_lines=spec['after_lines'])
        if spec['indent'] == 'item':
            set_para_indent(p, left_twips=ITEM_LEFT_TWIPS, left_chars=400)
        else:
            set_para_indent(p)
        for run in p.runs:
            set_run_font(run, east=spec['east'], latin=LATIN_FONT, size_pt=spec['size'],
                         bold=False, underline=False)
    else:  # 正文（含参考文献条目）
        set_para_spacing(p, line=LINE_TWIPS, rule='exact',
                         before_lines=0, after_lines=0)
        if in_ref:
            # 参考文献条目：不首行缩进
            set_para_indent(p)
        else:
            # 正文：首行缩进 2 字符
            set_para_indent(p, first_line_twips=BODY_FIRST_TWIPS, first_line_chars=200)
        for run in p.runs:
            set_run_font(run, east=BODY['east'], latin=LATIN_FONT, size_pt=BODY['size'],
                         bold=False, underline=False)


def apply_cover_line(p, text):
    """按模板封面格式设置单行。"""
    set_snap_to_grid(p, '0')
    if any(pat.search(text) for pat in
           (re.compile(r'哈尔滨工业大学'),)):
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_para_spacing(p, line=None, rule='auto')
        for run in p.runs:
            set_run_font(run, east='楷体_GB2312', latin=LATIN_FONT, size_pt=18, bold=True)
    elif re.search(r'硕士学位开题报告', text):
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_para_spacing(p, line=None, rule='auto')
        for run in p.runs:
            set_run_font(run, east='宋体', latin='Times New Roman', size_pt=24, bold=True)
    elif re.search(r'^[（(](学位论文|实践成果)[)）]', text.strip()):
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_para_spacing(p, before_pt=12, after_pt=0)
        for run in p.runs:
            set_run_font(run, east='宋体', latin=LATIN_FONT, size_pt=16, bold=True)
    elif re.search(r'^\s*题\s*目\s*[:：]', text):
        set_para_spacing(p, before_pt=12, after_pt=0)
        for run in p.runs:
            set_run_font(run, east='宋体', latin=LATIN_FONT, size_pt=18, bold=True)
    elif re.search(r'研究生院制', text):
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_para_spacing(p, line=None, rule='auto')
        for run in p.runs:
            set_run_font(run, east='宋体', latin=LATIN_FONT, size_pt=16, bold=True)
    else:  # 学院/学科/导师/研究生/学号/日期 等填表行
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        set_para_spacing(p, line=420, rule='auto', before_lines=0.5, after_lines=0)
        set_para_indent(p, left_twips=1260)  # 顶格左侧留出 2.25cm
        for run in p.runs:
            set_run_font(run, east='宋体', latin=LATIN_FONT, size_pt=16, bold=True, underline=True)


# ---------- 主流程 ----------
def set_page_setup(doc):
    sec = doc.sections[0]
    sec.page_width, sec.page_height = PAGE['width'], PAGE['height']
    sec.top_margin, sec.bottom_margin = PAGE['top'], PAGE['bottom']
    sec.left_margin, sec.right_margin = PAGE['left'], PAGE['right']
    sec.header_distance, sec.footer_distance = PAGE['header'], PAGE['footer']
    # 清除页眉（报告不设页眉）
    for h in (sec.header, sec.first_page_header):
        for p in h.paragraphs:
            for r in list(p.runs):
                r.text = ''


def build_cover(doc, meta):
    """在文档最前面插入模板封面（含若干空行），正文从分页符后开始。"""
    first = doc.paragraphs[0]

    def add(line_text):
        p = first.insert_paragraph_before('')
        if line_text:
            p.add_run(line_text)
            apply_cover_line(p, line_text)

    add(''); add('')
    add('哈尔滨工业大学')
    add('')
    add('硕士学位开题报告')
    add('')
    add('题  目：' + meta.get('title', ''))
    degree = meta.get('degree', '学位论文')
    add('（' + ('学位论文' if degree == '学位论文' else '实践成果') + '）')
    add('学  院（部）：' + meta.get('college', ''))
    add('学科/专业学位类别：' + meta.get('discipline', ''))
    add('导        师：' + meta.get('advisor', ''))
    add('研  究  生：' + meta.get('student', ''))
    add('学      号：' + meta.get('student_id', ''))
    add('开题报告日期：' + meta.get('date', ''))
    add(''); add(''); add('')
    add('研究生院制')
    add('')
    # 分页：正文另起一页
    pb = first.insert_paragraph_before('')
    pb.add_run().add_break(WD_BREAK.PAGE)


def main():
    ap = argparse.ArgumentParser(description='HIT 硕士开题报告格式调整')
    ap.add_argument('input', help='输入 docx 路径')
    ap.add_argument('output', help='输出 docx 路径')
    ap.add_argument('--cover-json', help='可选：封面元数据 JSON 文件路径')
    ap.add_argument('--dry-run', action='store_true', help='只统计不改写')
    args = ap.parse_args()

    if not os.path.exists(args.input):
        sys.exit('输入文件不存在: %s' % args.input)

    doc = Document(args.input)
    set_page_setup(doc)

    meta = {}
    if args.cover_json:
        with open(args.cover_json, encoding='utf-8') as f:
            meta = json.load(f)

    # 逐段分类并应用格式
    in_ref = False
    stats = {'cover': 0, 'L1': 0, 'L2': 0, 'L3': 0, 'L4': 0, 'body': 0}
    for i, p in enumerate(doc.paragraphs):
        text = p.text
        stripped = text.strip()
        if not stripped:
            # 空行保持空，不套格式
            continue
        if i < COVER_MAX_INDEX and is_cover_line(text) and not in_ref:
            apply_level(p, 'cover', in_ref=False)
            stats['cover'] += 1
            continue
        level = classify(text, p.style.name)
        # 进入参考文献区
        if REF_HEAD_RE.match(stripped):
            in_ref = True
        elif level in (1,) and in_ref:
            # 下一个节标题，退出参考文献区
            in_ref = False
        apply_level(p, level, in_ref=in_ref)
        key = {0: 'body', 1: 'L1', 2: 'L2', 3: 'L3', 4: 'L4'}.get(level, 'body')
        stats[key] += 1

    if args.cover_json and not any(is_cover_line(p.text) for p in doc.paragraphs[:10]):
        build_cover(doc, meta)

    if args.dry_run:
        print('dry-run stats:', stats)
        return

    doc.save(args.output)
    print('已输出: %s' % args.output)
    print('统计: %s' % stats)


if __name__ == '__main__':
    main()

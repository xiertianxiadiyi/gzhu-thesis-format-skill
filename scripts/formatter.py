"""
广州大学毕业论文（设计）格式修正工具

基于学校模板自动修正论文格式：
- 页面设置（封面/正文页边距）
- 字体字号（宋体/黑体/Times New Roman，三号/四号/小四/五号）
- 段落格式（首行缩进2字符、行距固定23pt）
- 标题样式（章标题、节标题、小节标题）
- 摘要/关键词格式
- 英文摘要格式
- 参考文献格式（悬挂缩进）
- 图表题注格式
- 页眉页脚格式
"""

import re
import sys
from pathlib import Path

from docx import Document
from docx.shared import Pt, Cm, Emu, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml, OxmlElement

# ── 单位转换常量 ──
# 1 inch = 914400 EMU
# 1 cm = 360000 EMU
# 1 pt = 12700 EMU

# ── 模板格式常量 ──
COVER_MARGINS = {"top": Cm(2.5), "bottom": Cm(2.5), "left": Cm(3.0), "right": Cm(2.6)}
BODY_MARGINS = {"top": Cm(2.5), "bottom": Cm(2.5), "left": Cm(3.2), "right": Cm(3.2)}

FONT_SONG = "宋体"
FONT_HEI = "黑体"
FONT_TIMES = "Times New Roman"
FONT_KAI = "楷体_GB2312"

SIZE_SANHAO = Pt(16)   # 三号
SIZE_SIHAO = Pt(14)    # 四号
SIZE_XIAOSI = Pt(12)   # 小四
SIZE_WUHAO = Pt(10.5)  # 五号
SIZE_XIAOWU = Pt(9)    # 小五

BODY_LINE_SPACING = 1.5   # 1.5倍行距
BODY_INDENT = Pt(24)      # 首行缩进2字符（约24pt）


def set_run_font(run, cn_font=None, en_font=None, size=None, bold=None, color=None):
    """设置 run 的中英文字体、字号、加粗、颜色。"""
    if cn_font:
        run.font.name = cn_font
        r = run._element
        rPr = r.find(qn('w:rPr'))
        if rPr is None:
            rPr = parse_xml(f'<w:rPr {nsdecls("w")}></w:rPr>')
            r.insert(0, rPr)
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = parse_xml(f'<w:rFonts {nsdecls("w")}></w:rFonts>')
            rPr.insert(0, rFonts)
        rFonts.set(qn('w:eastAsia'), cn_font)
        if en_font:
            rFonts.set(qn('w:ascii'), en_font)
            rFonts.set(qn('w:hAnsi'), en_font)
    if en_font and not cn_font:
        run.font.name = en_font
    if size:
        run.font.size = size
    if bold is not None:
        run.font.bold = bold
    if color:
        run.font.color.rgb = color


def set_paragraph_format(paragraph, alignment=None, line_spacing=None,
                          line_spacing_rule=None,
                          first_line_indent=None, space_before=None, space_after=None,
                          keep_with_next=None, outline_level=None):
    """设置段落格式。"""
    pf = paragraph.paragraph_format
    if alignment is not None:
        pf.alignment = alignment
    if line_spacing is not None:
        pf.line_spacing = line_spacing
        # 整数行距（twips）→固定值；浮点行距（倍数）→多倍行距
        if line_spacing_rule is None and isinstance(line_spacing, (int, float)):
            if isinstance(line_spacing, int):
                pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            else:
                pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    if first_line_indent is not None:
        pf.first_line_indent = first_line_indent
    if space_before is not None:
        pf.space_before = space_before
    if space_after is not None:
        pf.space_after = space_after
    if keep_with_next is not None:
        pf.keep_with_next = keep_with_next
    if outline_level is not None:
        pf.outline_level = outline_level


def is_heading_paragraph(text):
    """判断段落是否为标题。"""
    # 章标题：纯中文如"前言"、"第一章"、"第1章"、"总结与展望"、"1 前言"、"3 系统需求分析"
    chapter_patterns = [
        r'^前言$', r'^绪论$', r'^结论$', r'^致谢$', r'^参考文献$',
        r'^第[一二三四五六七八九十\d]+章',
        r'^[一二三四五六七八九十]+、',  # "一、研究背景"
        r'^总结与展望$', r'^结束语$',
        r'^[1-9]\d*\s+.+',  # "1 前言", "2 相关理论基础"
    ]
    for pat in chapter_patterns:
        if re.match(pat, text):
            return "chapter"
    return None


def remove_empty_paragraphs(doc):
    """移除空白段落。"""
    paragraphs_to_remove = []
    for i, p in enumerate(doc.paragraphs):
        if not p.text.strip():
            paragraphs_to_remove.append(p)
    for p in paragraphs_to_remove:
        p._element.getparent().remove(p._element)


def format_page_setup(doc):
    """修正页面设置。"""
    sections = doc.sections
    if len(sections) >= 1:
        # 封面页
        sec = sections[0]
        sec.page_width = Cm(21.0)
        sec.page_height = Cm(29.7)
        sec.top_margin = COVER_MARGINS["top"]
        sec.bottom_margin = COVER_MARGINS["bottom"]
        sec.left_margin = COVER_MARGINS["left"]
        sec.right_margin = COVER_MARGINS["right"]
    if len(sections) >= 2:
        # 正文页
        sec = sections[1]
        sec.page_width = Cm(21.0)
        sec.page_height = Cm(29.7)
        sec.top_margin = BODY_MARGINS["top"]
        sec.bottom_margin = BODY_MARGINS["bottom"]
        sec.left_margin = BODY_MARGINS["left"]
        sec.right_margin = BODY_MARGINS["right"]


def get_structure_zones(doc):
    """分析文档结构，返回各区域的段落索引范围。"""
    paragraphs = doc.paragraphs
    zones = {
        "cover": [],           # 封面
        "abstract_cn": [],     # 中文摘要
        "abstract_en": [],     # 英文摘要
        "toc": [],             # 目录
        "body": [],            # 正文
        "references": [],      # 参考文献
        "acknowledgement": [], # 致谢
    }

    current_zone = "cover"
    found_abstract = False
    found_abstract_en = False
    found_toc = False
    found_first_chapter = False
    found_references = False
    found_ack = False

    for i, p in enumerate(paragraphs):
        text = p.text.strip()

        # 检测中文摘要
        if text == "摘要" or (text.startswith("摘要") and not found_abstract):
            current_zone = "abstract_cn"
            found_abstract = True
        # 检测英文摘要
        elif re.match(r'^ABSTRACT', text, re.IGNORECASE) and not found_abstract_en:
            current_zone = "abstract_en"
            found_abstract_en = True
        # 检测目录
        elif text == "目  录" or text == "目录" and not found_toc:
            current_zone = "toc"
            found_toc = True
        # 检测正文开始（第一个章标题）
        elif not found_first_chapter and is_heading_paragraph(text) == "chapter":
            current_zone = "body"
            found_first_chapter = True
        # 检测参考文献
        elif text == "参考文献" and current_zone == "body":
            current_zone = "references"
            found_references = True
        # 检测致谢
        elif text == "致谢":
            current_zone = "acknowledgement"
            found_ack = True

        zones.setdefault(current_zone, []).append(i)

    return zones


SIZE_YIHAO = Pt(26)    # 一号

def format_cover(doc, zones):
    """修正封面格式（含表格）。"""
    cover_indices = zones.get("cover", [])

    for idx in cover_indices:
        p = doc.paragraphs[idx]
        text = p.text.strip()
        if not text:
            continue

        # 论文标题（"本 科 毕 业 论 文（设计）"）→ 楷体一号居中
        if ("毕业" in text.replace(" ", "") and "论文" in text.replace(" ", "") and "设计" in text.replace(" ", "")):
            for run in p.runs:
                set_run_font(run, cn_font=FONT_KAI, en_font=FONT_TIMES,
                            size=SIZE_YIHAO, bold=None)
            set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER)
            continue

        # "教 务 处 制" → 楷体三号居中
        if re.match(r'^教\s*务\s*处\s*制$', text):
            for run in p.runs:
                set_run_font(run, cn_font=FONT_KAI, en_font=FONT_TIMES,
                            size=SIZE_SANHAO)
            set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER)
            continue

        # 论文具体标题 → 黑体三号居中
        if len(text) >= 8 and not any(x in text for x in ["务", "处", "制", "班", "指导"]):
            for run in p.runs:
                set_run_font(run, cn_font=FONT_HEI, en_font=FONT_TIMES,
                            size=SIZE_SANHAO)
            set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER)
            continue

    # 修正封面表格（第一张表格）
    if doc.tables:
        cover_table = doc.tables[0]
        for row in cover_table.rows:
            cells = row.cells
            if len(cells) >= 2:
                # 左列：标签列 → 楷体小二号(18pt) 右对齐
                for p in cells[0].paragraphs:
                    for run in p.runs:
                        set_run_font(run, cn_font=FONT_KAI, en_font=FONT_TIMES,
                                    size=Pt(18), bold=None)
                    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.RIGHT,
                                        line_spacing=BODY_LINE_SPACING)
                # 右列：内容列 → 楷体三号(16pt) 居中
                for p in cells[1].paragraphs:
                    for run in p.runs:
                        set_run_font(run, cn_font=FONT_KAI, en_font=FONT_TIMES,
                                    size=SIZE_SANHAO, bold=None)
                    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                                        line_spacing=BODY_LINE_SPACING)


def _is_chapter_style(style_name, text):
    """判断是否为章标题样式。"""
    if style_name in ("章标题", "一级标题"):
        return True
    if style_name == "Normal" and is_heading_paragraph(text) == "chapter":
        return True
    return False


def _is_section_style(style_name):
    """判断是否为节标题（一级条/二级标题）样式。"""
    return style_name in ("一级条标题", "二级标题")


def _is_subsection_style(style_name):
    """判断是否为小节标题（二级条/三级标题）样式。"""
    return style_name in ("二级条标题", "三级标题")


def _is_caption_style(style_name):
    """判断是否为图表题注样式。"""
    return style_name in ("图和表", "Caption", "Normal (Web)")


def format_body_paragraph(p, is_reference=False):
    """修正正文段落格式。"""
    text = p.text.strip()
    if not text:
        return

    style_name = p.style.name

    # ── 章标题（添加段前分页，每章新起一页） ──
    if _is_chapter_style(style_name, text):
        for run in p.runs:
            set_run_font(run, cn_font=FONT_HEI, en_font=FONT_TIMES,
                        size=SIZE_SANHAO, bold=None)
        set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                            line_spacing=BODY_LINE_SPACING,
                            space_before=Pt(3), space_after=Pt(3))
        # 添加段前分页（"参考文献"和"致谢"除外，它们通常紧跟正文不强制分页）
        if text not in ("参考文献", "致谢"):
            pPr = p._element.get_or_add_pPr()
            pb = pPr.find(qn('w:pageBreakBefore'))
            if pb is None:
                pb = parse_xml(f'<w:pageBreakBefore {nsdecls("w")} />')
                pPr.append(pb)
        return

    # ── 节标题（如"1.1 XXX"） ──
    if _is_section_style(style_name):
        for run in p.runs:
            set_run_font(run, cn_font=FONT_HEI, en_font=FONT_TIMES, size=SIZE_XIAOSI)
        set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                            line_spacing=BODY_LINE_SPACING,
                            first_line_indent=None)
        return

    # ── 小节标题（如"1.1.1 XXX"） ──
    if _is_subsection_style(style_name):
        for run in p.runs:
            set_run_font(run, cn_font=FONT_HEI, en_font=FONT_TIMES, size=SIZE_XIAOSI)
        set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                            line_spacing=BODY_LINE_SPACING,
                            first_line_indent=None)
        return

    # ── 图和表（含内容模式匹配：以"图"或"表"开头+数字的段落） ──
    is_caption = _is_caption_style(style_name) or bool(
        re.match(r'^(图|表)\s*[\d.]+[\s\S]', text)
    )
    if is_caption:
        for run in p.runs:
            set_run_font(run, cn_font=FONT_SONG, en_font=FONT_TIMES, size=SIZE_WUHAO)
        set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                            line_spacing=BODY_LINE_SPACING,
                            first_line_indent=None)
        return

    # ── 公式 ──
    if style_name == "公式":
        for run in p.runs:
            set_run_font(run, cn_font=FONT_SONG, en_font=FONT_TIMES, size=SIZE_XIAOSI)
        set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                            line_spacing=BODY_LINE_SPACING)
        return

    # ── 参考文献条目 ──
    if is_reference or style_name == "参考文献":
        for run in p.runs:
            set_run_font(run, cn_font=FONT_SONG, en_font=FONT_TIMES, size=SIZE_XIAOSI)
        set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                            line_spacing=BODY_LINE_SPACING,
                            first_line_indent=None)
        return

    # ── 代码块（Consolas 等宽字体，不修改） ──
    has_mono = any(r.font.name and r.font.name in ("Consolas", "Courier New", "monospace")
                   for r in p.runs)
    if has_mono:
        return

    # ── 普通正文 ──
    for run in p.runs:
        if not run.font.name and run.font.size is None:
            set_run_font(run, cn_font=FONT_SONG, en_font=FONT_TIMES, size=SIZE_XIAOSI)

    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                        line_spacing=BODY_LINE_SPACING,
                        first_line_indent=BODY_INDENT)


def _add_blank_before_headings(doc, zones):
    """确保每个标题前都有一个空行（章/节/小节标题前均需空行）。"""
    body_indices = zones.get("body", [])
    if not body_indices:
        return

    paragraphs = list(doc.paragraphs)  # snapshot
    body_elem = paragraphs[body_indices[0]]._element.getparent()

    for idx in body_indices:
        p = paragraphs[idx]
        text = p.text.strip()
        if not text:
            continue
        s = p.style.name

        # 判断是否为需要前空行的标题类型
        is_heading = _is_chapter_style(s, text) or _is_section_style(s) \
                     or _is_subsection_style(s) or _is_caption_style(s)
        if not is_heading:
            continue
        # 如果"参考文献"或"致谢"本身也是独立章标题，也需要前空行
        # 但参考文献条目（正文111样式）不需要

        # 检查前面段落是否已经是空行
        if idx > 0:
            prev_p = paragraphs[idx - 1]
            if prev_p.text.strip():
                # 在此段前插入一个空行
                empty_p = OxmlElement('w:p')
                body_elem.insert(
                    list(body_elem).index(p._element),
                    empty_p
                )
                # 更新 zones（插入后索引会变，需要重建）
                # 但这里不重建，因为后续的循环会处理


def format_body(doc, zones):
    """修正正文格式。"""
    body_indices = zones.get("body", [])
    ref_indices = zones.get("references", [])

    for idx in body_indices:
        p = doc.paragraphs[idx]
        format_body_paragraph(p)

    for idx in ref_indices:
        p = doc.paragraphs[idx]
        format_body_paragraph(p, is_reference=True)


def format_abstract(doc, zones):
    """修正摘要格式。"""
    cn_indices = zones.get("abstract_cn", [])
    en_indices = zones.get("abstract_en", [])

    for idx in cn_indices:
        p = doc.paragraphs[idx]
        text = p.text.strip()
        if not text:
            continue

        # "摘要"标题行（标题词黑体四号，正文内容宋体小四不加粗）
        if text == "摘要" or text.startswith("摘要"):
            for run in p.runs:
                run_text = run.text
                if re.match(r'^摘要\s*$', run_text) or run_text == "摘要":
                    set_run_font(run, cn_font=FONT_HEI, en_font=FONT_TIMES,
                                size=SIZE_SIHAO)
                else:
                    set_run_font(run, cn_font=FONT_SONG, en_font=FONT_TIMES,
                                size=SIZE_XIAOSI, bold=False)
            set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                                line_spacing=BODY_LINE_SPACING)
            continue

        # "关键词"标题行（标题词黑体四号，关键词内容宋体小四）
        if text.startswith("关键词") or text.startswith("关键词"):
            for run in p.runs:
                run_text = run.text
                if re.match(r'^关键词\s*$', run_text) or run_text == "关键词":
                    set_run_font(run, cn_font=FONT_HEI, en_font=FONT_TIMES,
                                size=SIZE_SIHAO)
                else:
                    set_run_font(run, cn_font=FONT_SONG, en_font=FONT_TIMES,
                                size=SIZE_XIAOSI, bold=False)
            set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                                line_spacing=BODY_LINE_SPACING)
            continue

        # 摘要正文内容
        for run in p.runs:
            set_run_font(run, cn_font=FONT_SONG, en_font=FONT_TIMES, size=SIZE_XIAOSI)
        set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                            line_spacing=BODY_LINE_SPACING,
                            first_line_indent=BODY_INDENT)

    for idx in en_indices:
        p = doc.paragraphs[idx]
        text = p.text.strip()
        if not text:
            continue

        # ABSTRACT / KEY WORDS 标题行（标题词加粗，内容不加粗）
        is_abst_title = re.match(r'^(ABSTRACT|KEY\s*WORDS)', text, re.IGNORECASE)
        if is_abst_title:
            abst_kw = is_abst_title.group(1).upper()
            for run in p.runs:
                run_text = run.text
                # 判断这个 run 是否只是标题词本身
                is_title_word = bool(re.match(r'^(ABSTRACT|KEY|WORDS|KEY\s*WORDS)\s*$',
                                              run_text, re.IGNORECASE))
                if is_title_word:
                    set_run_font(run, en_font=FONT_TIMES,
                                size=SIZE_SIHAO, bold=True)
                else:
                    set_run_font(run, cn_font=FONT_SONG, en_font=FONT_TIMES,
                                size=SIZE_XIAOSI, bold=False)
            set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                                line_spacing=BODY_LINE_SPACING)
            continue

        # 英文摘要正文（不加粗）
        for run in p.runs:
            set_run_font(run, cn_font=FONT_SONG, en_font=FONT_TIMES,
                        size=SIZE_XIAOSI, bold=False)
        set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                            line_spacing=BODY_LINE_SPACING,
                            first_line_indent=BODY_INDENT)


def format_acknowledgement(doc, zones):
    """修正致谢格式。"""
    for idx in zones.get("acknowledgement", []):
        p = doc.paragraphs[idx]
        text = p.text.strip()
        if not text:
            continue

        if text == "致谢":
            for run in p.runs:
                set_run_font(run, cn_font=FONT_HEI, en_font=FONT_TIMES,
                            size=SIZE_SANHAO)
            set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                                line_spacing=BODY_LINE_SPACING)
        else:
            for run in p.runs:
                set_run_font(run, cn_font=FONT_SONG, en_font=FONT_TIMES,
                            size=SIZE_XIAOSI)
            set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                                line_spacing=BODY_LINE_SPACING,
                                first_line_indent=BODY_INDENT)


def fix_figure_table_numbering(doc, zones):
    """修正图序表序：将节序号格式改为章序号-序号格式。

    例如: 图2.1.2 检索增强生成（RAG）系统工作流程图
       -> 图2-2 检索增强生成（RAG）系统工作流程图

    逻辑:
    1. 扫描所有章标题(一级标题/章标题)，记录章序号及其后内容范围
    2. 在每个章的范围内，对图/表题注重新编号
    3. 图序格式: 图{章号}-{该章的第几张图}
    4. 表序格式: 表{章号}-{该章的第几个表}
    """
    body_indices = zones.get("body", [])
    if not body_indices:
        return

    paragraphs = doc.paragraphs

    # Step 1: 找出章标题和其所属段落范围
    chapter_starts = []  # (para_index, chapter_label, chapter_number)
    ch_pattern = re.compile(r'^(\d+)\s+')
    for idx in body_indices:
        p = paragraphs[idx]
        text = p.text.strip()
        if _is_chapter_style(p.style.name, text):
            m = ch_pattern.match(text)
            if m:
                chapter_starts.append((idx, text, int(m.group(1))))
            elif text in ("参考文献", "致谢"):
                continue  # skip backmatter

    if not chapter_starts:
        return

    # Step 2: 在每个章节范围内修正图/表编号
    for ci, (ch_start, ch_text, ch_num) in enumerate(chapter_starts):
        # Determine range end
        if ci + 1 < len(chapter_starts):
            ch_end = chapter_starts[ci + 1][0]
        else:
            ch_end = body_indices[-1] + 1

        fig_count = 0
        tbl_count = 0
        cap_pattern = re.compile(r'^(图|表)\s*[\d.]+\s*(.+)')

        for pi in range(ch_start + 1, ch_end):
            p = paragraphs[pi]
            text = p.text.strip()
            m = cap_pattern.match(text)
            if not m:
                continue

            prefix = m.group(1)  # "图" or "表"
            rest = m.group(2)    # the text after the number

            if prefix == "图":
                fig_count += 1
                new_num = f"{prefix}{ch_num}-{fig_count}"
            else:
                tbl_count += 1
                new_num = f"{prefix}{ch_num}-{tbl_count}"

            # 只修改小节号格式 -> 改为章-序号格式
            new_text = f"{new_num} {rest}"

            # 写回文本到 runs（保留 drawing 元素）
            # 先合并所有文本，找到需要修改的位置
            if p.runs:
                # 找到第一个包含可见文本的 run 并替换其文本
                for run in p.runs:
                    if run.text.strip():
                        # 保留 drawing 等子元素，只替换 w:t 节点的文本
                        t_elems = run._element.findall(
                            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                        if t_elems:
                            t_elems[0].text = new_text
                            # 清空后续 t 节点
                            for te in t_elems[1:]:
                                te.text = ''
                        else:
                            run.text = new_text
                        break
                # 清空其他 run 的文本（保留 drawing 等非文本子元素）
                found_first = False
                for run in p.runs:
                    if run.text.strip() and not found_first:
                        found_first = True
                        continue
                    if run.text.strip():
                        t_elems = run._element.findall(
                            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                        for te in t_elems:
                            te.text = ''


SIZE_XIAOER = Pt(18)   # 小二

def insert_toc(doc):
    """在摘要和正文之间插入自动目录域。

    在关键词段落后、第一个章标题前插入 TOC 域代码。
    用户打开 Word 后右键 -> 更新域 即可生成目录。
    """
    paragraphs = doc.paragraphs

    # 查找插入位置：关键词段落之后、第一个章标题之前
    toc_insert_after = None
    first_chapter_idx = None

    for i, p in enumerate(paragraphs):
        text = p.text.strip()
        if text.startswith("关键词"):
            toc_insert_after = i
        if toc_insert_after is not None and i > toc_insert_after and _is_chapter_style(p.style.name, text):
            first_chapter_idx = i
            break

    if toc_insert_after is None or first_chapter_idx is None:
        return

    # 检查是否已存在 TOC
    for i in range(toc_insert_after + 1, first_chapter_idx):
        instrs = paragraphs[i]._element.findall(
            './/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}instrText')
        for ins in instrs:
            if ins.text and 'TOC' in ins.text.upper():
                return  # TOC already exists

    # 在关键词段落后插入空行 + TOC 域
    body_elem = paragraphs[toc_insert_after]._element.getparent()

    # 插入一个空行
    empty_p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    empty_p.append(pPr)
    body_elem.insert(list(body_elem).index(paragraphs[toc_insert_after]._element) + 1, empty_p)

    # 插入目录标题（小二黑体居中）
    toc_title = OxmlElement('w:p')
    tPPr = OxmlElement('w:pPr')
    tjc = OxmlElement('w:jc')
    tjc.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'center')
    tPPr.append(tjc)
    toc_title.append(tPPr)
    tR = OxmlElement('w:r')
    tRPr = OxmlElement('w:rPr')
    tRF = OxmlElement('w:rFonts')
    tRF.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', FONT_HEI)
    tRF.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii', FONT_TIMES)
    tRPr.append(tRF)
    tSz = OxmlElement('w:sz')
    tSz.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', str(int(SIZE_XIAOER / 6350)))
    tRPr.append(tSz)
    tR.append(tRPr)
    tT = OxmlElement('w:t')
    tT.set('{http://schemas.xmlsoap.org/XML/1998/namespace}space', 'preserve')
    tT.text = '目  录'
    tR.append(tT)
    toc_title.append(tR)
    body_elem.insert(list(body_elem).index(paragraphs[toc_insert_after]._element) + 2, toc_title)

    # 插入 TOC 域
    toc_p = OxmlElement('w:p')
    tocPPr = OxmlElement('w:pPr')
    toc_p.append(tocPPr)

    # begin fldChar
    r1 = OxmlElement('w:r')
    fld1 = OxmlElement('w:fldChar')
    fld1.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldCharType', 'begin')
    r1.append(fld1)
    toc_p.append(r1)

    # instrText
    r2 = OxmlElement('w:r')
    instr = OxmlElement('w:instrText')
    instr.set('{http://schemas.xmlsoap.org/XML/1998/namespace}space', 'preserve')
    instr.text = ' TOC \\o "1-3" \\h \\z \\u '
    r2.append(instr)
    toc_p.append(r2)

    # separate
    r3 = OxmlElement('w:r')
    sep = OxmlElement('w:fldChar')
    sep.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldCharType', 'separate')
    r3.append(sep)
    toc_p.append(r3)

    # placeholder text
    r4 = OxmlElement('w:r')
    pt = OxmlElement('w:t')
    pt.text = '（请在 Word 中右键此处 → 更新域，以生成目录）'
    r4.append(pt)
    toc_p.append(r4)

    # end fldChar
    r5 = OxmlElement('w:r')
    fld2 = OxmlElement('w:fldChar')
    fld2.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldCharType', 'end')
    r5.append(fld2)
    toc_p.append(r5)

    body_elem.insert(list(body_elem).index(paragraphs[toc_insert_after]._element) + 3, toc_p)

    # 插入空行
    empty_p2 = OxmlElement('w:p')
    body_elem.insert(list(body_elem).index(paragraphs[toc_insert_after]._element) + 4, empty_p2)

    # 在目录后、第一章标题前插入分页符
    ch_elem = paragraphs[first_chapter_idx]._element
    ch_pPr = ch_elem.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
    if ch_pPr is None:
        ch_pPr = OxmlElement('w:pPr')
        ch_elem.insert(0, ch_pPr)
    pb = ch_pPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pageBreakBefore')
    if pb is None:
        pb = OxmlElement('w:pageBreakBefore')
        ch_pPr.append(pb)


def format_tables(doc):
    """修正正文表格内文字格式（跳过封面表格）。"""
    for ti, table in enumerate(doc.tables):
        if ti == 0:
            continue  # 封面表格已在 format_cover 中处理
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    text = paragraph.text.strip()
                    if not text:
                        continue
                    for run in paragraph.runs:
                        set_run_font(run, cn_font=FONT_SONG, en_font=FONT_TIMES,
                                    size=SIZE_XIAOSI)
                    set_paragraph_format(paragraph, line_spacing=BODY_LINE_SPACING)


def format_page_numbers(doc):
    """添加页码（页脚）。"""
    sections = doc.sections
    for i, section in enumerate(sections):
        footer = section.footer
        if not footer.paragraphs:
            footer.add_paragraph()

        # 清除现有页脚内容
        for p in footer.paragraphs:
            for run in p.runs:
                run.text = ""

        p = footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        run = p.add_run()
        set_run_font(run, cn_font=FONT_SONG, en_font=FONT_TIMES,
                    size=SIZE_XIAOWU)

        # 插入页码域代码
        fldChar1 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>')
        run._element.append(fldChar1)

        run2 = p.add_run()
        instrText = parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> PAGE </w:instrText>')
        run2._element.append(instrText)

        run3 = p.add_run()
        fldChar2 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>')
        run3._element.append(fldChar2)


def format_headers_footers(doc):
    """修正页眉页脚字体。"""
    for section in doc.sections:
        # 页眉
        if section.header:
            for p in section.header.paragraphs:
                if p.text.strip():
                    for run in p.runs:
                        set_run_font(run, cn_font=FONT_SONG, en_font=FONT_TIMES,
                                    size=SIZE_XIAOWU)
                    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER)

        # 页脚
        if section.footer:
            for p in section.footer.paragraphs:
                if p.text.strip():
                    for run in p.runs:
                        set_run_font(run, cn_font=FONT_SONG, en_font=FONT_TIMES,
                                    size=SIZE_XIAOWU)


def main():
    if len(sys.argv) < 3:
        print("用法: uv run --with python-docx python3 scripts/formatter.py input.docx output.docx")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]

    if not Path(input_path).exists():
        print(f"错误: 找不到输入文件 {input_path}")
        sys.exit(1)

    print(f"正在读取: {input_path}")
    doc = Document(input_path)

    print("正在修正格式...")

    # 1. 页面设置
    print("  [1/10] 页面设置...")
    format_page_setup(doc)

    # 2. 分析文档结构
    print("  [2/10] 分析文档结构...")
    zones = get_structure_zones(doc)

    # 3. 封面格式
    print("  [3/10] 封面格式...")
    format_cover(doc, zones)

    # 4. 图序表序修正（可在 SKILL.md 中查看说明）
    print("  [4/10] 图序表序修正...")
    fix_figure_table_numbering(doc, zones)

    # 5. 摘要格式
    print("  [5/10] 摘要格式...")
    format_abstract(doc, zones)

    # 6. 标题前空行 + 正文格式（含章标题分页）
    print("  [6/10] 标题前空行...")
    _add_blank_before_headings(doc, zones)
    print("  [7/10] 正文格式...")
    format_body(doc, zones)

    # 8. 致谢格式
    print("  [8/10] 致谢格式...")
    format_acknowledgement(doc, zones)

    # 9. 插入目录域
    print("  [9/10] 插入目录域...")
    insert_toc(doc)

    # 10. 表格格式 + 页眉页脚
    print("  [10/10] 表格格式...")
    format_tables(doc)
    format_headers_footers(doc)

    # 11. 页码
    print("  [11/11] 页码...")
    format_page_numbers(doc)

    # 保存
    doc.save(output_path)
    print(f"完成! 已保存到: {output_path}")


if __name__ == "__main__":
    main()

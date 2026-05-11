"""
广州大学毕业论文（设计）格式诊断工具

对比学校模板标准，诊断论文格式问题并输出报告。
"""

import re
import sys
from pathlib import Path
from collections import defaultdict

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ── 模板标准值 ──
STD_COVER_MARGINS = {"top": Cm(2.5), "bottom": Cm(2.5), "left": Cm(3.0), "right": Cm(2.6)}
STD_BODY_MARGINS = {"top": Cm(2.5), "bottom": Cm(2.5), "left": Cm(3.2), "right": Cm(3.2)}

SIZE_SANHAO = Pt(16)   # 三号
SIZE_SIHAO = Pt(14)    # 四号
SIZE_XIAOSI = Pt(12)   # 小四
SIZE_WUHAO = Pt(10.5)  # 五号

STD_LINE_SPACING = 460  # 固定值23pt (twips)


def is_heading_paragraph(text):
    """判断段落是否为章节标题。"""
    chapter_patterns = [
        r'^前言$', r'^绪论$', r'^结论$', r'^致谢$', r'^参考文献$',
        r'^第[一二三四五六七八九十\d]+章',
        r'^[一二三四五六七八九十]+、',
        r'^总结与展望$', r'^结束语$',
    ]
    for pat in chapter_patterns:
        if re.match(pat, text):
            return "chapter"
    return None


def diagnose_page_setup(doc):
    """诊断页面设置问题。"""
    issues = []
    sections = doc.sections

    for i, sec in enumerate(sections):
        label = "封面" if i == 0 else "正文"
        std = STD_COVER_MARGINS if i == 0 else STD_BODY_MARGINS

        for key, std_val in std.items():
            current = getattr(sec, f"{key}_margin", None)
            if current is not None:
                diff = abs(current - std_val)
                if diff > 5000:  # 0.5mm 容差
                    cur_cm = current / 360000
                    std_cm = std_val / 360000
                    issues.append(f"  第{i+1}节{label}页{key}边距: 当前{cur_cm:.1f}cm → 应为{std_cm:.1f}cm")

    return issues


def main():
    if len(sys.argv) < 2:
        print("用法: uv run --with python-docx python3 scripts/diagnose.py input.docx")
        sys.exit(1)

    input_path = sys.argv[1]
    if not Path(input_path).exists():
        print(f"错误: 找不到文件 {input_path}")
        sys.exit(1)

    doc = Document(input_path)
    paragraphs = doc.paragraphs

    issues = defaultdict(list)

    # ── 1. 页面设置 ──
    issues["页面设置"] = diagnose_page_setup(doc)

    # ── 2. 段落格式 ──
    for i, p in enumerate(paragraphs):
        text = p.text.strip()
        if not text:
            continue

        pf = p.paragraph_format
        style = p.style.name

        # 章标题
        if style == "章标题":
            # 检查字号
            for run in p.runs:
                if run.font.size is not None and abs(run.font.size - SIZE_SANHAO) > 100:
                    issues["章标题字号"].append(
                        f"  段落[{i}] \"{text[:30]}\": 字号{run.font.size/12700:.1f}pt → 应为16pt(三号)")
            # 检查行距
            if pf.line_spacing is None:
                issues["章标题行距"].append(f"  段落[{i}] \"{text[:30]}\": 行距未设置")

        # 节标题
        elif style == "一级条标题":
            for run in p.runs:
                if run.font.size is not None and abs(run.font.size - SIZE_XIAOSI) > 100:
                    issues["节标题字号"].append(
                        f"  段落[{i}] \"{text[:30]}\": 字号{run.font.size/12700:.1f}pt → 应为12pt(小四)")

        # 正文
        elif style == "Normal" and pf.alignment != WD_ALIGN_PARAGRAPH.CENTER:
            # 检查首行缩进
            if pf.first_line_indent is None or Emu_to_cm(pf.first_line_indent) and \
               pf.first_line_indent < Pt(10):
                if len(text) > 15 and not text.startswith(('图', '表')) and \
                   not re.match(r'^[（(]', text) and not re.match(r'^[a-zA-Z]', text):
                    issues["正文缩进"].append(f"  段落[{i}] \"{text[:40]}...\"")

            # 检查行距
            if pf.line_spacing is None:
                issues["正文行距"].append(f"  段落[{i}] 行距未设置")

        # 图和表
        elif style in ("图和表", "Caption"):
            for run in p.runs:
                if run.font.size is not None and abs(run.font.size - SIZE_WUHAO) > 100:
                    issues["图题表题字号"].append(
                        f"  段落[{i}] \"{text[:30]}\": 字号{run.font.size/12700:.1f}pt → 应为10.5pt(五号)")

    # ── 3. 字体字号总览 ──
    font_set = set()
    size_set = set()
    for p in paragraphs:
        for r in p.runs:
            if r.font.name:
                font_set.add(r.font.name)
            if r.font.size:
                size_set.add(r.font.size)

    if len(font_set) > 4:
        issues["字体检测"].append(
            f"  检测到{len(font_set)}种字体: {', '.join(sorted(font_set)[:8])}...")

    # 输出报告
    print("=" * 60)
    print("  广州大学毕业论文（设计）格式诊断报告")
    print("=" * 60)

    total = sum(len(v) for v in issues.values())
    if total == 0:
        print("\n✅ 未检测到格式问题，论文格式符合模板要求。")
        return

    print(f"\n共检测到 {total} 个格式问题:\n")
    for category, items in sorted(issues.items()):
        print(f"【{category}】共 {len(items)} 处")
        for item in items:
            print(item)
        print()

    print("=" * 60)
    print("提示: 可使用 formatter.py 自动修正格式")
    print(f"  python3 scripts/formatter.py {input_path} output.docx")


def Emu_to_cm(val):
    """EMU 转厘米。"""
    if val is None:
        return None
    return val / 360000


if __name__ == "__main__":
    main()

---
name: gzhu-thesis-format-skill
description: 广州大学毕业论文（设计）格式修正工具。基于学校模板自动修正论文页面设置、字体字号、段落缩进、行距、标题样式、标点符号、目录、页眉页脚、页码等格式。
---

# 广州大学毕业论文（设计）格式修正工具

根据「广州大学学生毕业论文（设计）模板(中文）」自动修正论文格式。

## 功能概览

| 功能 | 说明 | 脚本 |
|------|------|------|
| 格式诊断 | 对比模板诊断论文格式问题 | `diagnose.py` |
| 格式修正 | 自动按模板修正论文格式 | `formatter.py` |
| 图序表序修正 | 将"图X.X.X"改为"图X-N"格式 | `formatter.py` |
| 章标题分页 | 每章标题前自动添加分页符 | `formatter.py` |
| 目录生成 | 自动插入目录域代码 | `formatter.py` |
| 标点修复 | 修复中英文标点混用 | `punctuation.py` |

## 格式规范总览

```
页面：A4（21.0×29.7cm）
封面页边距：上2.5cm，下2.5cm，左3.0cm，右2.6cm
正文页边距：上2.5cm，下2.5cm，左3.2cm，右3.2cm

论文标题：三号（16pt）黑体，居中
章标题（1. 前言/第一章）：三号（16pt）黑体，居中
节标题（1.1）：小四号（12pt）黑体，顶格
小节标题（1.1.1）：小四号（12pt）黑体，顶格
正文：小四号（12pt）宋体/Times New Roman，首行缩进2字符（约24pt），行距1.5倍
摘要内容：小四号（12pt）宋体
摘要标题/关键词标题：四号（14pt）黑体
英文摘要（ABSTRACT）：四号（14pt）Times New Roman 加粗
关键词内容：小四号（12pt）宋体
参考文献：小四号（12pt）宋体
图题表题：五号（10.5pt）宋体
页眉页脚：小五号（9pt）宋体
```

## 使用方法

### 格式诊断

对比学校模板，输出论文格式差异报告：

```bash
uv run --with python-docx python3 scripts/diagnose.py input.docx
```

输出示例：
```
=== 格式诊断报告 ===

【页面设置】共 2 处差异
  - 封面页左边距：当前2.7cm → 应为3.0cm
  - 封面页右边距：当前2.5cm → 应为2.6cm

【段落格式】共 15 处差异
  - 第30-45段：缺少首行缩进2字符
  - 第26段（章标题"前言"）：行距当前1.5倍 → 应固定值23pt

【字体字号】共 8 处差异
  - 多处段落字号未显式设置（依赖默认值）→ 应小四号（12pt）宋体
  - 英文摘要（第20段）缺Times New Roman字体
```

### 格式修正

一键按模板修正：

```bash
uv run --with python-docx python3 scripts/formatter.py input.docx output.docx
```

处理流程：
1. 页面设置 → 按模板设置封面/正文页边距
2. 图序表序修正 → 将"图X.X.X"自动改为"图X-N"（章序号-序号）格式
3. 封面格式
4. 摘要/关键词 → 四号黑体（标题）、小四宋体（内容），内容不加粗
5. 英文摘要 → Times New Roman 四号标題加粗，内容小四不加粗
6. 正文格式 → 小四宋体、1.5倍行距、首行缩进2字符
7. 章标题 → 三号黑体居中，每章自动添加分页符
8. 节/小节标题 → 小四黑体左对齐
9. 致谢 → 三号黑体居中
10. 目录 → 自动插入目录域（Word中右键更新即可生成）
11. 图题表题 → 五号宋体居中
12. 参考文献 → 小四宋体、1.5倍行距
13. 页眉页脚 → 页码小五宋体居中
14. 表格内容格式

### 组合使用

```bash
# 先诊断
uv run --with python-docx python3 scripts/diagnose.py my_thesis.docx

# 修正格式
uv run --with python-docx python3 scripts/formatter.py my_thesis.docx my_thesis_formatted.docx

# 修正标点（可选）
uv run --with python-docx python3 scripts/punctuation.py my_thesis_formatted.docx my_thesis_clean.docx
```

## 文件结构

```
gzhu-thesis-format-skill/
├── SKILL.md
├── scripts/
│   ├── diagnose.py       # 格式诊断
│   ├── formatter.py      # 格式修正
│   └── punctuation.py    # 标点修复
```

## 依赖

- python-docx

使用 `uv run --with python-docx` 自动安装。

## 注意事项

1. **仅支持 .docx 格式**
2. **建议先备份原文件**
3. **输出文件需要系统安装对应字体（宋体、黑体、Times New Roman）才能正确显示**
4. **表格内文字也会被处理**

## 图序表序修正说明

formatter 会自动将错误的节序号格式图序改为正确的章序号格式：

- `图2.1.2 检索增强生成系统` → `图2-1 检索增强生成系统`
- `表4.1.3 系统架构解析` → `表4-1 系统架构解析`

修正逻辑：扫描每个章内的图/表标题，按「章序号-该章内第几张图/表」重新编号。图和表独立计数。

## 目录使用说明

formatter 自动在关键词后插入目录域代码。打开输出文件后：
1. 在 Word 中找到标记文字「（请在 Word 中右键此处 → 更新域，以生成目录）」
2. 右键点击 → 选择「更新域」→ 选择「更新整个目录」
3. 目录会自动根据正文中的章、节、小节标题生成

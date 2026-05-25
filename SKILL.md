---
name: zuel-formatter
description: >-
  中南财经政法大学 (ZUEL) 学术论文格式编译与校验工具。支持将 Markdown 渲染为 ZUEL 规范的 A4 学术 PDF，或校验现有 Word (.docx) 文件的格式合规性（校名、标题大小、段落缩进、页边距等）。
---

# ZUEL-Formatter: 学术论文格式编译与校验工具

## Overview
本技能专为**中南财经政法大学 (ZUEL)** 本科生论文及结课报告设计，提供以下两大核心功能：
1. **Markdown 编译 PDF/DOCX 双格式**：将标准 Markdown 渲染为符合 ZUEL 规范的 A4 学术报告 PDF（Edge 渲染）与可二次编辑微调的 Word `.docx` 格式文件（保留大字号封面、得分栏、首行缩进等样式）。
2. **Docx 格式校验**：自动读取并诊断学生编写的 Word `.docx` 论文文件，对其封面元素、打分栏、页边距、摘要关键词、多级标题字号和字体、正文缩进等进行合规性校验，并输出详细的 Markdown 和 JSON 诊断报告。

## Dependencies
- `markdown` (Python 库)
- `python-docx` (Python 库)
- `pypdf` 与 `reportlab` (Python 库)
- Microsoft Edge 浏览器 (默认安装路径：`C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe`)

## Quick Start
AI 助手可以直接通过 Python 调用此技能的脚本：

```powershell
# 1. 校验学生的 Docx 论文格式是否符合 ZUEL 规范
python scripts/zuel_formatter.py validate-docx --input "path/to/your_paper.docx"

# 2. 将 Markdown 格式的论文编译为合规的 A4 学术 PDF 和可微调的 Word DOCX
python scripts/zuel_formatter.py convert --input "path/to/your_paper.md" --header "中南财经政法大学 - 结课论文报告"
```

## Utility Scripts

### 1. `validate-docx` 子命令
读取 Docx 文件并执行完整的格式诊断。
*   `--input` / `-i` (必填): 待校验的 Word `.docx` 路径。
*   `--output-report` / `-o` (可选): 输出的 Markdown 报告路径（默认生成在同目录下 `*_validation_report.md`）。
*   `--output-json` / `-j` (可选): 输出的 JSON 数据路径。

### 2. `convert` 子命令
将 Markdown 文件编译为符合 ZUEL 标准的 PDF 和/或 DOCX 文件。默认同时生成两种格式，以实现交付与用户自主微调的完美结合。
*   `--input` / `-i` (必填): 源 Markdown 文件的路径。
*   `--output` / `-o` (可选): 输出的文件路径。系统将自动处理并输出同名 `.pdf` 和 `.docx` 文件。
*   `--docx` (可选): 仅输出 ZUEL 格式的 `.docx` 文件（可用于手动微调）。
*   `--pdf` (可选): 仅输出 ZUEL 格式的 `.pdf` 文件。
*   `--header` (可选): 页眉文字。
*   `--margin-top` / `--margin-bottom` / `--margin-left` / `--margin-right` (可选): 页边距（厘米）。

### 3. `preview` 子命令
将 Markdown 渲染为带排版样式的预览 HTML 文件。

## Common Mistakes
1. **Edge 浏览器路径缺失**：如未在标准 Windows 路径下找到 Edge 浏览器，PDF 渲染将报错。
2. **段落格式错乱**：在 DOCX 中若使用 ad-hoc 的硬空格进行段落对齐或首行缩进（而非标准样式），诊断工具会警告“首行未缩进 2 字符”。
3. **参考文献缺失**：论文末尾必须包含 “主要参考文献” 章节，否则诊断工具将报结构缺失错误。

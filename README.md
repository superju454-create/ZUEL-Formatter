# ZUEL-Formatter

[![Python Version](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Format](https://img.shields.io/badge/format-Markdown%20%2F%20PDF%20%2F%20Word-orange.svg)]()

**ZUEL-Formatter** 是一个专为**中南财经政法大学 (ZUEL)** 本科生学术论文与结课报告设计的自动化排版规范与合规性校验系统。项目基于 Markdown 标记语言，支持一键编译输出高保真 A4 PDF 和可用于人工微调的 Word DOCX 文档，并内置基于状态机的格式校验引擎，实现 0 虚警、全方位的合规性扫描与报告输出。

---

## ✨ 核心特性

- **📄 双通道高保真排版生成**：
  - **PDF 通道**：利用 Microsoft Edge 无头浏览器打印技术，将 Markdown 精准转换为符合 ZUEL 规范的 A4 学术 PDF。自动在正文页插入小五号宋体居中页码，并插入动态学术页眉。
  - **Word DOCX 通道**：基于 ZUEL 官方标准样式模板进行二次渲染，自动绘制 5行2列 的带下划线封面个人信息表格，强制适配中西文字体（宋体/黑体/Times New Roman），将正文行距重构为固定值 20 磅。
- **🔍 智能格式校验器 (`validate-docx`)**：
  - 自动读取 DOCX 文件，全面诊断页边距、封面打分栏、正文首行缩进、多级标题字号（一号/二号/三号/小三/四号/小四/五号）等指标。
  - **去点化图表题注校验**：依据国内规范校验表格上方 `表1 XXX`、图片下方 `图1 XXX` 题注的空格分隔、无英文点号、居中及五号宋体加粗格式。
  - **基于状态机的目录跳过诊断算法**：精准识别并跳过目录，避免对同名标题产生格式误报，并能动态校验引言前的正文大标题样式。
  - 输出美观的 Markdown 校验报告和结构化 JSON 报告。
- **📦 完全自包含与零硬编码**：
  - 模板及校徽文件已内置在资源目录中，脚本内部采用相对路径设计，解耦个人隐私路径，开箱即用。

---

## 🛠️ 安装要求与依赖

运行本工具需要系统预装以下环境：

1. **Python 3.8+**
2. **Pandoc**：用于 Markdown 到 DOCX 的基础转换。请确保 `pandoc` 已加入系统环境变量 `PATH`。
3. **Microsoft Edge 浏览器**：Windows 系统默认自带（若在其他系统运行，可在脚本内指定无头浏览器可执行路径 `EDGE_PATH`）。

### 安装 Python 依赖

```bash
pip install markdown python-docx pypdf reportlab
```

---

## 🚀 快速上手

在项目根目录下，直接通过 Python 运行 `scripts/zuel_formatter.py`：

### 1. 将 Markdown 格式的论文编译为 PDF 和 Word (双格式完美交付)

```bash
python scripts/zuel_formatter.py convert --input "path/to/your_paper.md" --output "output/my_paper" --header "基于Pandoc与无头浏览器的中南财经政法大学学术论文排版规范系统设计与实现"
```

- **说明**：系统将自动在 `output` 目录下输出 `my_paper.pdf` 和 `my_paper.docx`。封面学生填空表格预留空数据（仅保留下划线线框），方便打印或后续手动填空。

### 2. 校验 Word (DOCX) 论文格式合规性

```bash
python scripts/zuel_formatter.py validate-docx --input "path/to/your_paper.docx"
```

- **说明**：工具会执行 40 余项规范扫描，并在同级目录下输出 `*_validation_report.md` (Markdown 报告) 和 `*_validation_report.json` (JSON 数据)，清晰呈现每个校验项是否通过。

---

## 📂 项目结构

```text
ZUEL-Formatter/
├── README.md               # 项目使用说明
├── SKILL.md                # 技能详细定义与使用说明
├── .gitignore              # Git 忽略文件配置
├── resources/
│   ├── zuel_logo.png       # 中南财经政法大学官方校徽 (Base64内置)
│   └── zuel_template.docx  # 本科生论文报告官方 Word 样式参考模板
└── scripts/
    └── zuel_formatter.py   # 排版与格式校验核心脚本
```

---

## 📜 论文格式规范参考标准

系统严格对照中南财经政法大学本科生结课报告排版要求进行实现与校验：

| 排版要素 | 规范格式标准 | 系统校验状态 |
| :--- | :--- | :---: |
| **页边距** | 上 3.0cm，下 2.5cm，左 3.0cm，右 2.5cm | 🟢 自动校验 |
| **封面校名** | 初号 48pt 黑体加粗居中，校徽图片居中 | 🟢 自动校验 |
| **填空表格** | 5行2列信息表，20pt 楷体，左对齐带底边框下划线 | 🟢 自动校验 |
| **独立页面** | 摘要、目录、结语、参考文献必须独立分页 | 🟢 自动校验 |
| **独立标题** | `摘  要`、`目  录` 等为三号 16pt 黑体加粗，保留双字中间空格 | 🟢 自动校验 |
| **论文大标题** | 三号 16pt 黑体加粗居中，正文第一页（Page 4）顶格 | 🟢 自动校验 |
| **正文一级标题** | 三号 16pt 黑体加粗，段前 24 磅，段后 18 磅 | 🟢 自动校验 |
| **正文二级标题** | 四号 14pt 宋体加粗，段前 18 磅，段后 12 磅 | 🟢 自动校验 |
| **正文三级标题** | 小四 12pt 宋体加粗，段前 12 磅，段后 6 磅 | 🟢 自动校验 |
| **正文段落** | 小四 12pt 宋体/新罗马，**行距固定值 20 磅**，首行缩进 2 字符 | 🟢 自动校验 |
| **图表题注** | 五号 10.5pt 宋体加粗居中，**空格分隔（如：表1 XXX）无点号** | 🟢 自动校验 |

---

## 📄 开源协议

本项目基于 [MIT License](LICENSE) 协议开源。

# APA 7 Word Formatter｜APA 7 Word 格式助手

> AI understands the manuscript. Python applies the formatting. The author keeps control.

[中文介绍](#中文介绍) · [English Introduction](#english-introduction) · [使用方法](#运行) · [APA 官方依据](#apa-官网依据)

## 中文介绍

APA 格式一直是论文写作中非常消耗时间和精力的一部分。即使论文内容已经完成，学生仍然可能需要花费几个小时反复调整页边距、字体、双倍行距、页码、标题层级、参考文献悬挂缩进、表格边框以及图片和图表标题。学生论文与专业／投稿论文之间还有不同要求，学校、导师和期刊也可能规定额外格式。

普通的 Word 格式脚本只能机械地统一字体和行距，却无法理解一个段落究竟是论文标题、作者信息、一级标题、普通正文、块引用、参考文献，还是图表注释。完全让 AI 直接修改 Word，同样可能出现格式不稳定、误改原文、破坏公式或丢失 Zotero／EndNote 引用域等问题。

**APA 7 Word Formatter** 因此采用 **AI Skill + Python** 的组合：AI 负责阅读上下文、判断论文结构和发现视觉问题；Python 负责按照已确认的结构执行可测试、可重复的格式修改，并检查文字、公式、域代码、图片和嵌入资源是否得到保留。它不是简单地套用一个模板，而是把“理解文档”和“稳定改格式”分成两个相互检查的步骤。

开始前，用户必须选择 `student`（学生论文）或 `professional`（专业／投稿论文）。完成后默认只生成 **1 个新的 Word 格式副本**，不会覆盖原稿，也不会在论文旁边留下复杂的审计文件。反馈只包括：改了什么、APA 来源、改了原稿哪里，以及还要审核什么。

Python 引擎本身不会上传论文，也不需要单独的 AI API Key；使用 Skill 时，承载 Skill 的 AI 仍需要读取论文内容，因此数据处理方式取决于实际使用的平台，不能把“本地执行 Python”理解为“整个 AI 流程完全离线”。

## English Introduction

APA formatting is one of the most repetitive and time-consuming parts of academic writing. Even after the research and writing are complete, authors may still spend hours correcting margins, fonts, double spacing, page numbers, heading levels, reference indentation, table borders, figure sizing, and caption placement. Student papers and professional manuscripts also follow different requirements, while universities and journals may add their own exceptions.

Conventional Word-formatting scripts can apply styles, but they cannot reliably determine whether a paragraph is a paper title, author information, a section heading, body text, a block quotation, a reference entry, or a figure note. Letting an AI rewrite the document directly introduces a different risk: inconsistent formatting or accidental changes to text, equations, citation fields, links, and embedded objects.

**APA 7 Word Formatter** combines an **AI Skill with a deterministic Python engine**. The AI reads the manuscript in context, interprets its document hierarchy, classifies paragraphs and tables, and reviews rendered pages. Python then applies the approved formatting rules, verifies preservation of document content and embedded resources, and produces a new Word file without overwriting the source.

The project supports both `student` and `professional` paper profiles. It can format page layout, paragraph spacing, headers, title pages, APA heading levels, abstracts, block quotations, references, appendices, editable Word tables, figures, and captions. It can also inspect native Word charts and optionally export supported vector content or reconstruct simple charts as SVG without pretending that raster images are true vectors.

By default, one run produces only **one formatted DOCX copy**. The user receives concise feedback explaining what changed, which official APA sources were used, where formatting was applied in the original manuscript, and which items still require human review. The tool is a formatting and review assistant, not an official APA compliance certificate.

## 为什么使用 Skill + Python？

根据 [OpenAI 的 Skill 文档](https://learn.chatgpt.com/docs/build-skills)，Skill 可以把操作说明、参考资料和可执行脚本组合成可重复使用的工作流程。本项目将任务拆成两部分：

- **AI Skill**：理解语义、判断文档层级、识别表格和图片的作用、处理不确定情况，并逐页检查输出。
- **Python 引擎**：修改 Word XML 和样式、执行固定规则、保护原文与嵌入资源，并拒绝覆盖原稿或使用过期结构配置。

AI 不确定的对象会被标为 `preserve` 并保留原样，而不是为了追求“全自动”强行修改。

## 文档层级解读

这个项目不会把 Word 中所有加粗文字都当成标题，也不会把所有 Word 表格都当成数据表。Skill 会同时阅读段落内容、相邻上下文、Word 样式、分页位置、正文顺序、表格单元格和视觉对象，再建立类似下面的文档层级：

```text
Paper
├── Title page
│   ├── Paper title
│   └── Author, affiliation, course or author-note information
├── Abstract
│   └── Keywords
├── Main text
│   ├── Level 1 heading
│   │   ├── Level 2 heading
│   │   │   ├── Level 3 heading
│   │   │   └── Level 4/5 run-in heading + body text
│   │   ├── Body paragraphs
│   │   ├── Block quotations
│   │   ├── Tables
│   │   │   ├── Table number
│   │   │   ├── Table title
│   │   │   ├── Header rows and data cells
│   │   │   └── Table note
│   │   └── Figures
│   │       ├── Figure number
│   │       ├── Figure title
│   │       ├── Image or native chart
│   │       └── Figure note
├── References
│   └── Individual reference entries
└── Appendices
    ├── Appendix label and title
    └── Appendix text, tables and figures
```

层级判断会直接影响格式。例如，普通正文使用首行缩进，而摘要首段不缩进；参考文献使用悬挂缩进；一级标题与三级标题的对齐和斜体规则不同；四、五级标题必须和后续正文处于同一段；表格编号和标题位于表格上方，表格注释位于下方。

如果上下文不足以确定标题级别、标题页范围、表头行数或对象类型，Skill 会保留该对象，并在最终反馈的“还要审核”中说明，而不会静默猜测。

## 使用 Skill

完整技能在 `skills/apa7-word-formatter/`，含 `SKILL.md`、执行脚本和操作参考。可以把这个文件夹复制到个人 Skill 目录：

```text
$HOME/.agents/skills/apa7-word-formatter/
```

也可以只在当前项目中使用，把它放到项目根目录的：

```text
.agents/skills/apa7-word-formatter/
```

Codex 通常会自动发现新 Skill；如果没有显示，重新启动 Codex。也可以使用 `$skill-installer`，让它从这个 GitHub 仓库安装 `skills/apa7-word-formatter`。安装完成后，把 Word 文档放进工作区，然后输入：

> 使用 $apa7-word-formatter，把这份 Word 按 APA 7 学生论文格式处理，并检查所有页面。

投稿论文请明确选择 professional，并提供 running head；要矢量导出时同时说明。Skill 会先让 Python 读取结构，由 AI 给每个段落及表格填写角色和依据，再通过 `apa7_workflow.py apply` 执行。未完成判断、配置与原稿不匹配、或不确定对象被要求强制改写时，执行入口会拒绝。

渲染使用当前环境的 documents 技能及其 `render_docx.py`。渲染结果和逐页复核记录会绑定输出 Word 的哈希；没有渲染条件时只能交付待复核状态，不能宣称已经完成视觉验收。

`apa7_workflow.py` 不会自行调用大模型；运行 `prepare` 后，负责使用 Skill 的 AI 会完成分类并填写配置。单独运行 Python 并不等于自动启用了 AI。

它可以重复执行已确定的排版规则；**目前不能把任意 Word 文档无条件转换成完全符合 APA 7 的成稿**。尤其是图片内文字、未标注的标题层级、文献条目的语义、图表首次提及顺序、作者信息以及最终分页，需要复核。运行成功也只表示支持范围内的格式处理成功。

## 运行

Python 3.10 或以上，先安装一次依赖：

```sh
python3 -m pip install -r requirements.txt
```

在本文件夹运行以下命令，会打开文件选择窗口。使用包含 Tk 的 Python 发行版才能显示窗口；没有 Tk 时使用下面的命令行方式。

```sh
python3 apa7_format.py
```

论文类型是必选项。学生论文：

```sh
python3 apa7_format.py "/完整路径/论文.docx" --profile student
```

专业／投稿论文：

```sh
python3 apa7_format.py "/完整路径/论文.docx" --profile professional --running-head "A SHORT PAPER TITLE"
```

默认输出就在原文档旁边，使用带时间戳的新文件名：

- `论文_APA7_时间.docx`：格式副本。

命令行会直接显示一份简单反馈，只包含“改了什么、APA 来源、改了原稿哪里、还要审核”。如果确实需要另存这份反馈，可加 `--save-feedback`，生成一个 `.feedback.html`；默认不生成。

运行时会检查原稿哈希、正文、域代码、原生公式、图片和嵌入资源的保留情况；这些检查不能替代逐页视觉检查。

## 识别图片和导出矢量图

文件选择窗口中可勾选“同时提取图片／原始矢量，并将支持的原生图表导出 SVG”。命令行对应：

```sh
python3 apa7_format.py "/完整路径/论文.docx" --profile student --export-visuals
```

只识别内容类型、不修改文档：

```sh
python3 apa7_format.py "/完整路径/论文.docx" --inspect-visuals
```

识别依赖 Word 文件内部的对象结构，不只是页面外观：原生表格具有行列结构；原生图表具有图表类型和数据缓存；插图具有图片关系和媒体文件。照片里拍到的表格、截图里的折线图仍会被识别为位图，本版不执行 OCR 或曲线取点。

导出的文件放在同名 `.visuals` 文件夹内，含逐项说明的 JSON 清单。Word 中原有的可编辑图表仍保留。

| 来源 | 导出方式 | 准确性边界 |
|---|---|---|
| 原始 SVG／EMF／WMF | 按原始字节提取 | SVG 可能混有位图；EMF／WMF 也不因扩展名而被认定为纯矢量 |
| 简单原生柱形／折线／散点图 | 根据完整的数据缓存重绘为独立 SVG | 是数据重绘，不是视觉效果逐像素复制；缓存可能与外链工作簿不同步，应核对数据及图例 |
| 不支持的复杂图表 | 清单列明原因 | 双轴、误差线、趋势线、平滑曲线、缓存不完整等不会被简化成不准确的图 |
| PNG／JPEG 等位图 | 原样提取 | 不会包装成带位图的 SVG 并宣称已经矢量化 |
| Word 原生表格 | 保持为可编辑 Word 表格 | 表格排版与图片矢量化是不同处理过程 |

SVG 格式允许混合矢量与位图，[W3C 的 SVG 规范](https://www.w3.org/TR/SVG/embedded.html)明确规定 `image` 元素可引用 PNG／JPEG。因此必须检查实际内容，而不能只看文件后缀。

导出模块是同目录的 `apa7_visuals.py`，使用 ReportLab 绘制真正的图形与文本。仅使用核心格式处理时可以只复制 `apa7_format.py` 并安装 python-docx；使用图形识别和矢量导出时，应保留两个 Python 文件并安装完整 requirements。

## 当前自动处理范围

| 内容 | 已实现 | 需要注意 |
|---|---|---|
| 页面 | 四边 1 英寸页边距；保留纸张大小和方向 | 学校可能要求 A4 或装订边距；多栏文档需复核 |
| 字体与正文 | 支持官网列出的六种字体组合；正文双倍行距、左齐、首行 0.5 英寸 | 保留普通正文中的粗斜体、上下标和特定符号字体；中文字体不宣称为 APA 指定字体 |
| 页眉 | 自动页码；投稿模式增加全大写短标题 | 复杂原页眉默认保留并报告；页脚需检查是否有重复页码 |
| 标题 | 已有 Heading 1–3 按对应层级排版 | 不根据粗体外观猜层级；四／五级需明确同行标题范围 |
| 标题页 | 已确认的 Title 和元信息居中，标题加粗，设置标题起始留白与作者间隔 | 不编造作者、课程、单位和作者注；复杂标题页垂直位置需复核 |
| 摘要／参考文献／附录 | 根据英文区段标签和明确角色设置对齐、缩进、区段分页 | 引文及参考文献内容、排序、标点和对应关系未自动校正；中文区段可用配置明确标记 |
| 表格 | 简单原生 Word 表去竖线和网格，保留顶底／表头横线，表头重复；单元格单倍行距 | 默认第一行为表头；复杂合并／嵌套表保留并报告；不把问卷和数据表一概认定为同类 |
| 图及图表标题 | 识别单独编号段和紧邻视觉对象的标题段，编号粗体、标题斜体；超宽内嵌图片等比例缩小 | 编号与标题在同一段时暂不拆分；浮动图、图内字体、坐标、清晰度、版权及编号顺序需复核 |
| 块引用 | 对明确标为 Quote 的段落整体缩进、双倍行距 | 40 词规则、引用范围、后续段缩进与出处需确认 |

旧 `.doc` 可先在 Word 中另存为 `.docx`。也可通过 `--soffice "/可执行文件路径/soffice"` 调用已安装的 LibreOffice 转换器；转换后保留检查以转换所得 DOCX 为基准，旧格式转换保真需另查。本版不支持加密 Word、`.docm` 宏文件及 RTF。带未处理修订的文档会停止处理，要求先由作者审阅修订。

## 复杂文档如何明确结构

推荐先生成结构配置草稿（只读 Word，不生成格式副本）：

```sh
python3 apa7_format.py "/完整路径/论文.docx" --profile student --prepare-config "/完整路径/结构确认.json"
```

草稿的 `_review` 包含原段落文本、推定角色、建议标题页范围与分模式检查清单；顶层 `roles` 初始为空，不把推定自动当成已确认。将明确需要覆盖的角色填入顶层 `roles` 后使用 `--config`。草稿带有 `source_sha256`，只有同一版本原稿才允许应用；如果在 Word 中编辑或重新保存过原稿，应重新生成并确认配置。

先只读列出段落及原有样式：

```sh
python3 apa7_format.py "/完整路径/论文.docx" --inspect
```

程序的段落编号为 1 起始，包含正文空段，但不包含表格单元格、文本框或内容控件内段落。按当前文档的实际编号建立 JSON 配置；不要把下面示例编号直接套到自己的论文。

```json
{
  "roles": {
    "1": "title",
    "12": "heading2",
    "16": "heading4",
    "22": "reference",
    "23": "reference",
    "30": "preserve"
  },
  "title_page": [1, 7],
  "run_in_headings": {
    "16": "Response Accuracy."
  },
  "table_header_rows": {
    "1": 2
  },
  "table_roles": {
    "1": "data",
    "2": "preserve"
  },
  "replace_headers": false
}
```

```sh
python3 apa7_format.py "/完整路径/论文.docx" --profile student --config "/完整路径/配置.json"
```

`run_in_headings` 中的文本必须与该段开头完全匹配，包含结尾英文句号，并且后面已经接有同段正文。`title_page` 是现有标题页的起止段号。`replace_headers=true` 会在格式副本中重建页眉；原文件仍保留。自动结果中的 `assumption` 表示基于结构推定，需要确认。

简单反馈会列出仍要人工确认的标题页信息、结构、引用、表格、图像与逐页视觉检查。这些项目不会因为 Python 成功运行就自动标成通过。

## APA 官网依据

规范只取自 APA 官方网站；核查日期为 2026-09-12。每个规则有官网标题、可点击链接、简短英文原文、中文实现摘要。原文保存在 `apa7_format.py` 的 `SOURCES`，运行下列命令即可查看：

```sh
python3 apa7_format.py --sources
```

主要入口：[Paper Format](https://apastyle.apa.org/style-grammar-guidelines/paper-format)、[Table Setup](https://apastyle.apa.org/style-grammar-guidelines/tables-figures/tables)、[Figure Setup](https://apastyle.apa.org/style-grammar-guidelines/tables-figures/figures)、[Reference List Setup](https://apastyle.apa.org/style-grammar-guidelines/paper-format/reference-list)。

官网免费页面并不等于完整出版手册。官方第七版手册仍是完整规范参考；本工具没有声称覆盖手册中所有文献类型、统计表达、伦理及版权要求。

## Python 与 Skill 的分工

Python 负责可测试的文档读写、资源保留、配置校验和排版；专用 Skill 指导 AI 结合论文内容补全角色配置、解读审计结果并渲染复核。[OpenAI 官方 Skill 文档](https://learn.chatgpt.com/docs/build-skills)说明 Skill 可包含说明、资源和可执行脚本。

使用通用 documents 技能时，让用户指定的 APA 规则优先于通用设计默认值，尤其是图注位置、表格边框和段落间距。不为了视觉美观重写论文或改动统计结论。

## 开发验证

```sh
python3 -m unittest discover -s . -p "test_*.py" -v
```

`qa_demo.py` 是内部渲染验收用的人工样例生成器，需要 Pillow；日常格式化不需要此依赖。它生成的研究数据与参考文献均是测试占位材料。

维护引擎后，运行 `python3 build_skill.py` 将三个引擎脚本及依赖清单同步到技能包。使用 `--archive /一个尚不存在的路径/apa7-word-formatter.zip` 可生成独立技能压缩包。不要手工修改技能目录中这些自动同步的脚本，以免与主程序分叉。

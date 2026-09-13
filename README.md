# APA 7 Word Formatter

> AI understands the paper. Python formats it reliably. You receive one new Word document.

Version 0.7 · Student papers · Professional manuscripts · Tables · Figures · References · Optional Word styles

[中文](#中文介绍) · [English](#english-introduction)

---

## 中文介绍

APA 格式一直是论文写作中非常消耗时间和精力的一部分。即使论文内容已经完成，学生仍然可能需要花费几个小时反复调整页边距、字体、双倍行距、页码、标题层级、参考文献悬挂缩进、表格边框以及图片和图表标题。学生论文与专业／投稿论文之间还有不同要求，学校、导师和期刊也可能规定额外格式。

普通的 Word 格式脚本只能机械地统一字体和行距，却无法理解一个段落究竟是论文标题、作者信息、一级标题、普通正文、块引用、参考文献，还是图表注释。完全让 AI 直接修改 Word，同样可能出现格式不稳定、误改原文、破坏公式或丢失 Zotero／EndNote 引用域等问题。

APA 7 Word Formatter 因此采用 AI Skill + Python 的组合：AI 负责阅读上下文、判断论文结构和发现视觉问题；Python 负责按照已确认的结构执行可测试、可重复的格式修改，并检查文字、公式、域代码、图片和嵌入资源是否得到保留。它不是简单地套用一个模板，而是把“理解文档”和“稳定改格式”分成两个相互检查的步骤。

开始前，用户必须选择 student（学生论文）或 professional（专业／投稿论文）。完成后默认只生成 1 个新的 Word 格式副本，不会覆盖原稿，也不会在论文旁边留下复杂的审计文件。反馈只包括：改了什么、APA 来源、改了原稿哪里，以及还要审核什么。

### Skill 文件结构

```text
skills/apa7-word-formatter/
├── SKILL.md                     AI 的主要工作说明
├── agents/
│   └── openai.yaml              Skill 的显示名称和简介
├── references/
│   └── workflow.md              结构判断、执行和逐页检查流程
├── scripts/
│   ├── apa7_format.py           Word 排版与保存后核验
│   ├── apa7_workflow.py         AI 审核记录和完整流程
│   ├── apa7_visuals.py          图片、图表识别与可选矢量导出
│   └── requirements.txt         Python 依赖
└── build-manifest.json          版本号和脚本完整性信息
```

### 这个 Skill 能做什么

使用时只需先确认 `student`（学生论文）或 `professional`（专业／投稿论文）。默认流程是：**AI 识别文档结构 → Python 执行排版 → 保存后重新核验 → 逐页检查 → 交付一个新 DOCX 和简短反馈**。

#### 默认一键完成

| 功能 | 自动处理的内容 |
| --- | --- |
| 原稿保护 | 不覆盖源文件；不改写论文内容；检查文字、超链接、域代码、Zotero／EndNote 引用、公式、图片、图表和嵌入资源是否保留 |
| 页面设置 | 四边 1 英寸页边距、APA 允许的字体、双倍行距、段前段后间距和页码 |
| 学生／专业模式 | 按所选模式处理标题页和页眉；专业模式使用作者提供的 running head |
| 文档层级 | 根据上下文识别标题页、摘要、关键词、正文、一级至五级标题、块引用、参考文献、图表说明和附录 |
| 正文与参考文献 | 设置对齐、首行缩进、块引用缩进、参考文献悬挂缩进和标题层级格式 |
| 表格 | 对简单 Word 数据表应用 APA 风格横线并去除竖线；复杂、合并或嵌套表格会保留并提示复核 |
| 图片与图表 | 识别位图、SVG／EMF／WMF、原生 Word 图表和 Word 表格；过宽图片等比例缩小，不拉伸、不伪造矢量图 |
| 引文检查 | 提示可能缺少对应参考文献的作者—年份引文，以及可能未在正文出现的文献条目；不会自动改写文献内容 |
| 成品核验 | 重新打开 DOCX，检查核心格式是否真正保存，并逐页查看分页、遮挡、溢出和图表位置 |
| 简短反馈 | 只说明改了什么、APA 来源、改了原稿哪里，以及还需要作者确认什么 |

#### AI 会判断，但不会擅自猜测

| 判断内容 | 处理方式 |
| --- | --- |
| 段落是什么 | AI 结合文字、前后文、Word 样式、对象顺序和页面外观判断，而不是看到粗体就当成标题 |
| 表格是否适合自动处理 | 简单数据表自动排版；复杂表格保留原状并提醒人工检查 |
| 图表属于哪种对象 | 先识别格式和数据是否完整，再决定保留、缩放、提取或重绘 |
| 不确定的内容 | 保留原文和对象，不强行修改，并写入“还需审核” |

#### 可选的额外功能

| 功能 | 什么时候使用 | 默认状态 |
| --- | --- | --- |
| 矢量图导出 | 需要单独取得原始 SVG／EMF／WMF，或把数据完整的简单柱形图、折线图、散点图重绘为 SVG | 关闭 |
| 继续写作样式 | 排版后还要在成品里新增正文、标题、参考文献、图表说明或块引用时使用 | 关闭 |
| 保存 HTML 反馈 | 需要把简短反馈另外保存成可打开的网页文件时使用 | 关闭 |
| 只检查不修改 | 想先看文档结构、风险和可处理对象，不立即生成排版副本时使用 | 按需 |
| 更换 APA 允许字体 | 学校、导师或期刊指定其他 APA 可接受字体时使用 | 按需 |

“继续写作样式”不是最终排版必需功能。它只是给 Word 增加 `APA7 Body`、`APA7 Heading 1`、`APA7 Reference` 等可重复选择的样式，方便用户在已经排好的副本里继续新增内容。如果论文已经写完，这项功能没有必要开启，因此从 v0.7 起默认关闭，也不会让 Word 的样式列表变得杂乱。

格式规则来自 APA 官方网站，包括 [Paper Format](https://apastyle.apa.org/style-grammar-guidelines/paper-format)、[Headings](https://apastyle.apa.org/style-grammar-guidelines/paper-format/headings)、[Table Setup](https://apastyle.apa.org/style-grammar-guidelines/tables-figures/tables)、[Figure Setup](https://apastyle.apa.org/style-grammar-guidelines/tables-figures/figures)、[Reference List Setup](https://apastyle.apa.org/style-grammar-guidelines/paper-format/reference-list) 和 [Author-Date Citation System](https://apastyle.apa.org/style-grammar-guidelines/citations/basic-principles/author-date)。运行 `python3 apa7_format.py --sources` 可以查看项目保存的官方原文摘录和直接链接。

### 如何使用

#### 1. 安装项目

需要 Python 3.10 或以上版本：

```sh
git clone https://github.com/weirdfishes_2120/apa7-word-formatter.git
cd apa7-word-formatter
python3 -m pip install -r requirements.txt
```

要在 Codex 中作为个人 Skill 使用，把 `skills/apa7-word-formatter` 文件夹复制到：

```text
$HOME/.agents/skills/apa7-word-formatter/
```

如果只想让当前项目使用，也可以放在项目根目录的 `.agents/skills/`。关于 Skill 的发现和调用方式，可参考 [OpenAI 官方 Build skills 文档](https://learn.chatgpt.com/docs/build-skills)。

#### 2. 把 Word 文档放进工作区

推荐使用 `.docx`。旧 `.doc` 文件应先在 Word 中另存为 `.docx`，并确认转换后没有丢失内容。

#### 3. 直接告诉 AI 你的要求

学生论文：

> 使用 $apa7-word-formatter，把这份 Word 按 APA 7 学生论文格式处理，并检查所有页面。

专业／投稿论文：

> 使用 $apa7-word-formatter，把这份 Word 按 APA 7 专业论文格式处理。Running head 是 “SHORT PAPER TITLE”。

需要矢量导出：

> 使用 $apa7-word-formatter 按 APA 7 学生论文格式处理，同时导出能够安全提取或重绘的矢量图。

排版后还要继续写作：

> 使用 $apa7-word-formatter 按 APA 7 学生论文格式处理，并加入方便继续写作的 APA7 Word 样式。

如果没有说明论文类型，Skill 会先询问。之后它会分析结构、执行排版、重新检查成品并逐页查看。默认只交付一个新的 Word 副本。

#### 4. 只运行 Python

打开本地文件选择窗口：

```sh
python3 apa7_format.py
```

命令行处理学生论文：

```sh
python3 apa7_format.py "/path/to/paper.docx" --profile student
```

命令行处理专业论文：

```sh
python3 apa7_format.py "/path/to/paper.docx" --profile professional --running-head "SHORT PAPER TITLE"
```

如需在成品中继续写作，可选加上：

```sh
python3 apa7_format.py "/path/to/paper.docx" --profile student --add-styles
```

单独运行 Python 可以执行固定排版和保护检查，但不会代替 AI 对标题层级、对象用途和页面外观的判断。

### 使用边界

- 学校、导师或期刊的明确要求优先于 APA 通用格式。
- 工具不会编造作者信息、参考文献、研究数据或缺失内容。
- 引文匹配、标题大小写、复杂表格、图片清晰度、版权和最终分页仍可能需要作者确认。
- 这是 APA 7 格式与审核助手，不是 APA 官方认证工具。

---

## English Introduction

APA 7 formatting often takes far more time than expected. A paper may contain a title page, several heading levels, an abstract, block quotations, references, tables, figures, native Word charts, and appendices. A conventional script can standardize fonts and spacing, but it cannot reliably understand what each paragraph or object means. Allowing an AI to edit a Word file freely creates a different risk: accidental changes to wording, equations, citation fields, or embedded content.

**APA 7 Word Formatter** combines both approaches:

- The **AI Skill** reads the manuscript in context, interprets its structure, classifies paragraphs and objects, and reviews the rendered pages.
- The **Python engine** applies the approved formatting consistently, reopens the saved file to verify the result, and protects the source manuscript and its contents.

Before formatting, the user chooses `student` or `professional`. By default, the formatter creates one new `.docx` copy and never overwrites the source. Its purpose is to remove repetitive formatting work while clearly identifying anything that still requires human judgment.

### Skill Structure

```text
skills/apa7-word-formatter/
├── SKILL.md                     Main instructions for the AI
├── agents/
│   └── openai.yaml              Display name and Skill metadata
├── references/
│   └── workflow.md              Classification, execution, and page-review workflow
├── scripts/
│   ├── apa7_format.py           Word formatting and post-save verification
│   ├── apa7_workflow.py         AI review record and controlled workflow
│   ├── apa7_visuals.py          Visual inspection and optional vector export
│   └── requirements.txt         Python dependencies
└── build-manifest.json          Version and script integrity information
```

`SKILL.md` explains how the AI should analyze a paper. `workflow.md` defines when formatting is safe and when an object must be preserved. The Python files in `scripts/` perform the actual Word edits. This separation prevents the AI from freely rewriting the document and prevents Python from blindly guessing its structure.

### What the Skill Can Do

The user first confirms `student` or `professional`. The default workflow is then: **AI reads the structure → Python applies the formatting → the saved file is reopened and verified → every page is reviewed → one new DOCX and concise feedback are returned**.

#### Automatic by Default

| Feature | What it does automatically |
| --- | --- |
| Source protection | Never overwrites the source; does not rewrite the paper; checks that text, links, fields, Zotero/EndNote citations, equations, visuals, charts, and embedded resources remain present |
| Page setup | Applies 1-inch margins, an APA-supported font, double spacing, paragraph spacing, and page numbers |
| Student/professional mode | Handles the title page and headers for the selected mode; professional mode uses the author's running head |
| Document hierarchy | Uses context to identify the title page, abstract, keywords, body, Levels 1–5 headings, block quotations, references, captions, notes, and appendices |
| Body and references | Applies alignment, first-line indents, block-quotation indents, hanging reference indents, and heading formatting |
| Tables | Formats simple editable Word data tables with APA-style horizontal rules and no vertical rules; preserves complex, merged, or nested tables for review |
| Figures and charts | Distinguishes raster images, SVG/EMF/WMF, native Word charts, and Word tables; proportionally scales oversized images without stretching or fake vector conversion |
| Citation checks | Flags likely author–year citations without a matching reference and references without an obvious in-text citation; never edits bibliography content automatically |
| Saved-result verification | Reopens the DOCX to confirm that core formatting was saved, then reviews each page for overflow, obstruction, pagination, and visual placement |
| Concise feedback | Reports only what changed, the APA sources, the affected locations, and what still needs author review |

#### AI-Assisted Decisions

| Decision | Behavior |
| --- | --- |
| What a paragraph represents | Uses wording, neighboring content, Word styles, object order, and page appearance; bold text alone is not treated as proof of a heading |
| Whether a table is safe to format | Formats straightforward data tables and preserves complicated ones for human review |
| What kind of visual an object is | Identifies the format and available data before preserving, scaling, extracting, or reconstructing it |
| Ambiguous content | Preserves the original instead of silently guessing and lists it under remaining review |

#### Optional Extras

| Feature | When it is useful | Default |
| --- | --- | --- |
| Vector export | Extract original SVG/EMF/WMF or reconstruct supported simple bar, line, and scatter charts as SVG when complete data are available | Off |
| Continued-writing styles | Add new body text, headings, references, captions, notes, or block quotations inside the formatted copy | Off |
| Saved HTML feedback | Keep the concise feedback as a separate browser-readable file | Off |
| Inspect without editing | Review structure, risks, and supported objects before creating a formatted copy | On request |
| Alternate APA-supported font | Meet a university, instructor, or journal requirement for another permitted font | On request |

Continued-writing styles are not required for final formatting. They only add reusable styles such as `APA7 Body`, `APA7 Heading 1`, and `APA7 Reference` to Word, so new content added later can reuse the correct formatting. If the paper is already complete, leave this option off. It is disabled by default from v0.7 to keep the Word style gallery uncluttered.

Formatting rules are linked to official APA pages, including [Paper Format](https://apastyle.apa.org/style-grammar-guidelines/paper-format), [Headings](https://apastyle.apa.org/style-grammar-guidelines/paper-format/headings), [Table Setup](https://apastyle.apa.org/style-grammar-guidelines/tables-figures/tables), [Figure Setup](https://apastyle.apa.org/style-grammar-guidelines/tables-figures/figures), [Reference List Setup](https://apastyle.apa.org/style-grammar-guidelines/paper-format/reference-list), and the [Author-Date Citation System](https://apastyle.apa.org/style-grammar-guidelines/citations/basic-principles/author-date). Run `python3 apa7_format.py --sources` to view the stored short quotations and direct links.

### How to Use It

#### 1. Install the project

Python 3.10 or later is required:

```sh
git clone https://github.com/weirdfishes_2120/apa7-word-formatter.git
cd apa7-word-formatter
python3 -m pip install -r requirements.txt
```

To use it as a personal Codex Skill, copy `skills/apa7-word-formatter` to:

```text
$HOME/.agents/skills/apa7-word-formatter/
```

For project-only use, place it in `.agents/skills/` at the repository root. See OpenAI's official [Build skills documentation](https://learn.chatgpt.com/docs/build-skills) for Skill discovery and invocation.

#### 2. Add the Word file to the workspace

`.docx` is recommended. Save legacy `.doc` files as `.docx` in Word first and confirm that the conversion preserved the document.

#### 3. Ask the AI to format it

Student paper:

> Use $apa7-word-formatter to format this Word document as an APA 7 student paper and review every page.

Professional manuscript:

> Use $apa7-word-formatter to format this Word document as an APA 7 professional paper. The running head is “SHORT PAPER TITLE.”

Optional vector export:

> Use $apa7-word-formatter to format this document as an APA 7 student paper and export any visuals that can be safely extracted or reconstructed as vectors.

Continued-writing styles:

> Use $apa7-word-formatter to format this document as an APA 7 student paper and add reusable APA7 Word styles for continued writing.

If the paper type is missing, the Skill asks for it first. It then analyzes the structure, applies the formatting, verifies the saved result, and reviews the pages. The default deliverable is one new Word copy.

#### 4. Run Python directly

Open the local file picker:

```sh
python3 apa7_format.py
```

Format a student paper from the command line:

```sh
python3 apa7_format.py "/path/to/paper.docx" --profile student
```

Format a professional manuscript:

```sh
python3 apa7_format.py "/path/to/paper.docx" --profile professional --running-head "SHORT PAPER TITLE"
```

Optionally add reusable Word styles when the formatted copy will still be edited:

```sh
python3 apa7_format.py "/path/to/paper.docx" --profile student --add-styles
```

Running Python by itself applies deterministic formatting and preservation checks. It does not replace the AI's judgment about heading hierarchy, object purpose, or page appearance.

### Important Boundaries

- Specific university, instructor, or journal requirements take priority over general APA formatting.
- The tool never invents author information, references, research data, or missing content.
- Citation matching, title case, complex tables, figure clarity, copyright, and final pagination may still require author review.
- This is an APA 7 formatting and review assistant, not an official APA certification tool.

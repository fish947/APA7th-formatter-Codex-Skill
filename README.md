# APA 7 Word Formatter

> AI understands the paper. Python formats it reliably. You receive one new Word document.

Version 0.11.0 · Student papers · Professional manuscripts · Statistics · Equations · Tables · Figures · References

[中文](#中文介绍) · [English](#english-introduction)

---

## 中文介绍

APA 格式一直是论文写作中非常消耗时间和精力的一部分。即使论文内容已经完成，学生仍然可能需要花费几个小时反复调整页边距、字体、双倍行距、页码、标题层级、参考文献、统计数据、公式、表格以及图片和图表。学生论文与专业／投稿论文之间还有不同要求，学校、导师和期刊也可能规定额外格式。

普通的 Word 格式脚本只能机械地统一字体和行距，却无法理解一个段落究竟是论文标题、作者信息、一级标题、普通正文、块引用、参考文献，还是图表注释。完全让 AI 直接修改 Word，同样可能出现格式不稳定、误改原文、破坏公式或丢失 Zotero／EndNote 引用域等问题。

APA 7 Word Formatter 因此采用 AI Skill + Python 的组合：AI 负责阅读上下文、判断论文结构和发现视觉问题；Python 负责按照已确认的结构执行可测试、可重复的格式修改，并检查文字、公式、域代码、图片和嵌入资源是否得到保留。它不是简单地套用一个模板，而是把“理解文档”和“稳定改格式”分成两个相互检查的步骤。

开始前，用户必须选择 student（学生论文）或 professional（专业／投稿论文）。完成后默认只生成 1 个新的 Word 格式副本，不会覆盖原稿，也不会在论文旁边留下复杂的审计文件。反馈只包括：改了什么、APA 来源、改了原稿哪里，以及还要审核什么。

### 项目文件结构

```text
apa7-word-formatter/
├── apa7_format.py               主程序，也是版本号的唯一来源
├── apa7_statistics.py           统计表达和公式检查
├── apa7_visuals.py              图片、图表和矢量导出
├── apa7_workflow.py             AI 审核与逐页检查流程
├── build_skill.py               自动同步可安装 Skill
├── requirements.txt             Python 依赖
├── tests/                       全部自动测试
├── tools/
│   └── qa_demo.py               生成合成文档做视觉测试
├── skills/apa7-word-formatter/  可直接安装的 Skill
└── .github/workflows/tests.yml  GitHub 自动测试
```

根目录的四个 `apa7_*.py` 是唯一维护来源。`skills/.../scripts/` 是构建生成的安装副本，不需要手动修改；运行 `python3 build_skill.py` 会同步脚本并更新版本与完整性哈希。

### 这个 Skill 能做什么

使用时只需先确认 `student`（学生论文）或 `professional`（专业／投稿论文）。默认流程是：**AI 识别文档结构 → Python 执行排版 → 保存后重新核验 → 逐页检查 → 交付一个新 DOCX 和简短反馈**。

#### 默认一键完成

| 功能 | 自动处理的内容 |
| --- | --- |
| 原稿保护 | 不覆盖源文件；不改变研究内容或统计数值；检查文字、超链接、域代码、Zotero／EndNote 引用、公式、图片、图表和嵌入资源是否保留 |
| 页面设置 | 四边 1 英寸页边距、APA 允许的字体、双倍行距、段前段后间距和页码 |
| 学生／专业模式 | 按所选模式处理标题页和页眉；专业模式使用作者提供的 running head |
| 标题页与前置页检查 | 按学生／专业模式核对标题页候选信息、正文首页重复标题、Abstract 和 Keywords 的基本关系；不会补写姓名、课程或作者注 |
| 文档层级 | 根据上下文识别标题页、摘要、关键词、正文、一级至五级标题、块引用、参考文献、图表说明和附录 |
| 正文与参考文献 | 设置对齐、首行缩进、块引用缩进、参考文献悬挂缩进和标题层级格式 |
| DOI／URL 超链接 | 将已识别参考文献中已有的完整 `http://`、`https://` 和 `https://doi.org/` 文字设为可点击 Word 链接；显示文字不变，不搜索或编造缺失 DOI |
| 参考文献质量 | 离线检查重复条目、重复 DOI、明显的字母排序异常，以及同作者同年文献的 `2024a/2024b` 后缀；不自动改写书目信息 |
| 统计数据汇报 | 识别 *p*、*t*、*F*、*M*、*SD*、效应量等表达；在数值不变的前提下规范统计符号、运算符空格和前导零 |
| 统计完整性提醒 | 提示 `p = .000`、异常精度、可能缺少自由度、精确 *p* 值、效应量或置信区间；不会自动计算或补写数据 |
| 公式 | 识别 Word 原生公式和纯文本公式候选；规范已确认独立公式的段落缩进和间距，保留公式内容并检查编号、标点和页面位置 |
| 表格 | 对简单 Word 数据表应用 APA 风格横线并去除竖线；复杂、合并或嵌套表格会保留并提示复核 |
| 图片与图表 | 识别位图、SVG／EMF／WMF、原生 Word 图表和 Word 表格；过宽图片等比例缩小，不拉伸、不伪造矢量图 |
| 图表／公式编号 | 检查重复编号、跳号、顺序异常、正文提及但对象不存在，以及有编号但正文未明确提及；不自动重新编号 |
| 图表与说明配对 | 检查表格／图形附近是否有匹配的编号和独立标题，识别孤立说明和多个图形共用说明的情况；不会移动对象 |
| 引文检查 | 提示可能缺少对应参考文献的作者—年份引文，以及可能未在正文出现的文献条目；不会自动改写文献内容 |
| 一次完整预检 | 把标题页、引用、参考文献、统计、公式、图表和编号检查汇总成一份简洁结果，避免分散运行多个检查 |
| 成品核验 | 重新打开 DOCX，检查核心格式是否真正保存，并逐页查看分页、遮挡、溢出和图表位置 |
| 简短反馈 | 只说明改了什么、APA 来源、改了原稿哪里，以及还需要作者确认什么 |

#### AI 会判断，但不会擅自猜测

| 判断内容 | 处理方式 |
| --- | --- |
| 段落是什么 | AI 结合文字、前后文、Word 样式、对象顺序和页面外观判断，而不是看到粗体就当成标题 |
| 表格是否适合自动处理 | 简单数据表自动排版；复杂表格保留原状并提醒人工检查 |
| 图表属于哪种对象 | 先识别格式和数据是否完整，再决定保留、缩放、提取或重绘 |
| 统计报告是否完整 | AI 结合检验类型判断自由度、*p* 值、效应量和置信区间是否需要补充；Python 不会猜测缺失结果 |
| 公式如何处理 | AI 区分行内公式、独立公式和普通文字；不确定时保留，不自动转换或重编号 |
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

格式规则来自 APA 官方网站，包括 [Paper Format](https://apastyle.apa.org/style-grammar-guidelines/paper-format)、[学生论文设置指南](https://www.apa.org/ed/precollege/psn/2020/09/apa-style-student-papers)、[DOIs and URLs](https://apastyle.apa.org/style-grammar-guidelines/references/dois-urls)、[Numbers and Statistics Guide](https://apastyle.apa.org/instructional-aids/numbers-statistics-guide.pdf)、[APA Research Transparency Standards](https://www.apa.org/pubs/journals/resources/standards-disclosures)、[Headings](https://apastyle.apa.org/style-grammar-guidelines/paper-format/headings)、[Table Setup](https://apastyle.apa.org/style-grammar-guidelines/tables-figures/tables)、[Figure Setup](https://apastyle.apa.org/style-grammar-guidelines/tables-figures/figures)、[Reference List Setup](https://apastyle.apa.org/style-grammar-guidelines/paper-format/reference-list)、[Citing Works With the Same Author and Date](https://apastyle.apa.org/style-grammar-guidelines/citations/basic-principles/same-year-author) 和 [Author-Date Citation System](https://apastyle.apa.org/style-grammar-guidelines/citations/basic-principles/author-date)。运行 `python3 apa7_format.py --sources` 可以查看项目保存的官方原文摘录和直接链接。

参考文献链接遵循 APA 7 的电子文档规则：完整 DOI／URL 保持可点击，但链接不必显示成蓝色下划线。工具只给原稿已经写出的完整地址增加链接，不会联网搜索 DOI，也不会修改文献管理器生成的域。

参考文献质量审核也不会自动移动或改写条目。它只对能清晰识别作者和年份的条目检查字母顺序，并提示重复文字、重复 DOI 和同作者同年后缀问题。非拉丁字母排序、文献类型、标题大小写、斜体位置和书目事实仍由 AI 与作者核对。

统计和公式处理遵守一个简单原则：工具可以把 `P=0.032` 规范为斜体 *p* `= .032`，但不会把 `p = .000` 擅自改成 `p < .001`，也不会舍入、重算、补写效应量或改变显著性。涉及研究结论的内容只会列入“还要审核”。

图、表和公式的一致性检查同样只做安全提示。它能识别 `Tables 1–3`、`Figure A1` 和 `Equation (2)` 等明确写法，但不会自动改编号或破坏 Word 交叉引用域；图片是论文图还是校徽等装饰对象，仍由 AI 结合页面判断。

v0.11 会把图表说明与实际对象按正文顺序配对，并把标题页、引用、参考文献、统计、公式和图表检查合并为一次预检。配对只用于发现风险，不会移动图表，也不会把校徽等装饰图片强行当作论文图。

APA 官方原文包括：“Alphabetize references according to the first word of the reference”；“When multiple references have an identical author (or authors) and publication year, include a lowercase letter after the year”；“Number figures in the order in which they are mentioned in your paper”；“Report exact p values to two or three decimals”；“Number all displayed equations consecutively”。完整上下文和例外见上面的官方链接。

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

标题页、统计数据、公式、图表／公式编号、图表说明配对、参考文献质量和 DOI／URL 链接检查默认包含在一次完整预检中，不需要额外开关。

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

单独运行 Python 可以执行固定排版、统计表达规范、公式识别和保护检查，但不会代替 AI 对标题层级、统计方法、公式含义、对象用途和页面外观的判断。

维护或发布前运行：

```sh
python3 -m unittest discover -s tests -p 'test_*.py' -v
python3 build_skill.py
```

### 使用边界

- 学校、导师或期刊的明确要求优先于 APA 通用格式。
- 工具不会编造作者信息、参考文献、研究数据、统计结果或公式内容。
- 工具不会搜索、猜测或补写缺失的 DOI／URL；文献管理器域中的链接保持原状并提示复核。
- 参考文献检查不会自动移动条目，也不会认证作者、年份、标题、期刊或 DOI 的事实正确性。
- 工具不会自动舍入或重算统计值，不会改变显著性，也不会自动重编号图、表或公式。
- 引文匹配、统计方法、标题大小写、复杂表格、图片清晰度、版权和最终分页仍可能需要作者确认。
- 这是 APA 7 格式与审核助手，不是 APA 官方认证工具。

---

## English Introduction

APA 7 formatting often takes far more time than expected. A paper may contain a title page, headings, an abstract, block quotations, references, statistical results, equations, tables, figures, Word charts, and appendices. A conventional script can standardize fonts and spacing, but it cannot reliably understand what each paragraph or object means. Allowing an AI to edit a Word file freely creates a different risk: accidental changes to wording, numeric results, equations, citation fields, or embedded content.

**APA 7 Word Formatter** combines both approaches:

- The **AI Skill** reads the manuscript in context, interprets its structure, classifies paragraphs and objects, and reviews the rendered pages.
- The **Python engine** applies the approved formatting consistently, reopens the saved file to verify the result, and protects the source manuscript and its contents.

Before formatting, the user chooses `student` or `professional`. By default, the formatter creates one new `.docx` copy and never overwrites the source. Its purpose is to remove repetitive formatting work while clearly identifying anything that still requires human judgment.

### Project Structure

```text
apa7-word-formatter/
├── apa7_format.py               Main program and single version source
├── apa7_statistics.py           Statistical and equation checks
├── apa7_visuals.py              Visual inventory and vector export
├── apa7_workflow.py             AI review and page-QA workflow
├── build_skill.py               Synchronizes the installable Skill
├── requirements.txt             Python dependencies
├── tests/                       Automated regression tests
├── tools/
│   └── qa_demo.py               Synthetic visual-QA fixture generator
├── skills/apa7-word-formatter/  Installable Skill package
└── .github/workflows/tests.yml  GitHub test workflow
```

The four root `apa7_*.py` files are the only maintained code source. Files under `skills/.../scripts/` are generated installation copies. Run `python3 build_skill.py` to synchronize them and refresh the version and integrity hashes.

### What the Skill Can Do

The user first confirms `student` or `professional`. The default workflow is then: **AI reads the structure → Python applies the formatting → the saved file is reopened and verified → every page is reviewed → one new DOCX and concise feedback are returned**.

#### Automatic by Default

| Feature | What it does automatically |
| --- | --- |
| Source protection | Never overwrites the source or changes research meaning or numeric results; checks that text, links, fields, Zotero/EndNote citations, equations, visuals, charts, and embedded resources remain present |
| Page setup | Applies 1-inch margins, an APA-supported font, double spacing, paragraph spacing, and page numbers |
| Student/professional mode | Handles the title page and headers for the selected mode; professional mode uses the author's running head |
| Front-matter checks | Reviews title-page candidates, the repeated title at the start of the text, and the basic relationship between Abstract and Keywords; never invents names, course details, or author notes |
| Document hierarchy | Uses context to identify the title page, abstract, keywords, body, Levels 1–5 headings, block quotations, references, captions, notes, and appendices |
| Body and references | Applies alignment, first-line indents, block-quotation indents, hanging reference indents, and heading formatting |
| DOI/URL hyperlinks | Makes complete `http://`, `https://`, and `https://doi.org/` text in identified references clickable in Word while preserving displayed text; never searches for or invents a missing DOI |
| Reference quality | Checks for duplicate entries, duplicate DOIs, clear alphabetical-order problems, and missing or inconsistent `2024a/2024b` suffixes for the same authors and year; never rewrites bibliographic facts |
| Statistical reporting | Recognizes expressions such as *p*, *t*, *F*, *M*, *SD*, and effect sizes; normalizes symbols, operator spacing, and leading zeros only when the numeric value remains unchanged |
| Reporting reminders | Flags `p = .000`, questionable precision, and possibly missing degrees of freedom, exact *p* values, effect sizes, or confidence intervals; never calculates or invents results |
| Equations | Identifies native Word math and plain-text formula candidates; formats the paragraph layout of confirmed display equations while preserving equation content and reviewing numbering, punctuation, and placement |
| Tables | Formats simple editable Word data tables with APA-style horizontal rules and no vertical rules; preserves complex, merged, or nested tables for review |
| Figures and charts | Distinguishes raster images, SVG/EMF/WMF, native Word charts, and Word tables; proportionally scales oversized images without stretching or fake vector conversion |
| Number and callout checks | Flags duplicate, skipped, or out-of-order table, figure, and equation numbers; also checks explicit body callouts against identified labels without renumbering anything |
| Caption-object pairing | Checks whether tables and drawing objects have nearby matching numbers and separate titles; flags orphan caption parts and possible multi-panel figures without moving objects |
| Citation checks | Flags likely author–year citations without a matching reference and references without an obvious in-text citation; never edits bibliography content automatically |
| Unified preflight | Summarizes front matter, citations, references, statistics, equations, visuals, and numbering in one compact review instead of separate runs |
| Saved-result verification | Reopens the DOCX to confirm that core formatting was saved, then reviews each page for overflow, obstruction, pagination, and visual placement |
| Concise feedback | Reports only what changed, the APA sources, the affected locations, and what still needs author review |

#### AI-Assisted Decisions

| Decision | Behavior |
| --- | --- |
| What a paragraph represents | Uses wording, neighboring content, Word styles, object order, and page appearance; bold text alone is not treated as proof of a heading |
| Whether a table is safe to format | Formats straightforward data tables and preserves complicated ones for human review |
| What kind of visual an object is | Identifies the format and available data before preserving, scaling, extracting, or reconstructing it |
| Whether a statistical report is complete | Uses the analysis type to assess degrees of freedom, *p* values, effect sizes, and confidence intervals; Python never guesses missing results |
| How an equation should be treated | Distinguishes inline math, display equations, and ordinary text; preserves uncertain content and never converts or renumbers it automatically |
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

Formatting rules are linked to official APA pages, including [Paper Format](https://apastyle.apa.org/style-grammar-guidelines/paper-format), the [student-paper setup guide](https://www.apa.org/ed/precollege/psn/2020/09/apa-style-student-papers), [DOIs and URLs](https://apastyle.apa.org/style-grammar-guidelines/references/dois-urls), the [Numbers and Statistics Guide](https://apastyle.apa.org/instructional-aids/numbers-statistics-guide.pdf), [APA Research Transparency Standards](https://www.apa.org/pubs/journals/resources/standards-disclosures), [Headings](https://apastyle.apa.org/style-grammar-guidelines/paper-format/headings), [Table Setup](https://apastyle.apa.org/style-grammar-guidelines/tables-figures/tables), [Figure Setup](https://apastyle.apa.org/style-grammar-guidelines/tables-figures/figures), [Reference List Setup](https://apastyle.apa.org/style-grammar-guidelines/paper-format/reference-list), [Citing Works With the Same Author and Date](https://apastyle.apa.org/style-grammar-guidelines/citations/basic-principles/same-year-author), and the [Author-Date Citation System](https://apastyle.apa.org/style-grammar-guidelines/citations/basic-principles/author-date). Run `python3 apa7_format.py --sources` to view the stored short quotations and direct links.

For electronic documents, complete DOI/URL text remains clickable, although APA permits either ordinary black text or the word processor's blue-underlined appearance. The formatter only links addresses already written in the manuscript; it does not search for DOIs or modify reference-manager fields.

The reference-quality audit is deliberately conservative. It checks alphabetical order only when author and year patterns are clear, and it reports duplicate text, duplicate DOIs, and same-author/same-year suffix problems without moving or rewriting entries. Non-Latin collation, source type, title capitalization, italics, and bibliographic facts still require AI and author review.

The statistics and equation formatter follows a strict boundary: it may normalize `P=0.032` to italic *p* `= .032`, but it will not silently change `p = .000` into `p < .001`, round or recompute a value, invent an effect size, or change significance. Anything that could affect the scientific result remains an author-review item.

The numbering check is also review-only. It recognizes explicit forms such as `Tables 1–3`, `Figure A1`, and `Equation (2)`, but never renumbers objects or rewrites Word cross-reference fields. The AI still decides whether a drawing is a research figure or a decorative object such as a logo.

Version 0.11 pairs captions with nearby objects in document order and combines front-matter, citation, reference, statistical, equation, visual, and numbering checks into one preflight. Pairing is a review aid only: it does not move objects or force decorative images to become research figures.

Key APA wording includes “Alphabetize references according to the first word of the reference,” “When multiple references have an identical author (or authors) and publication year, include a lowercase letter after the year,” “Number figures in the order in which they are mentioned in your paper,” “Report exact p values to two or three decimals,” and “Number all displayed equations consecutively.” Follow the official links above for the full context and exceptions.

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

Front-matter, statistical-reporting, equation, table/figure/equation numbering, caption-object pairing, reference-quality, and DOI/URL checks are included in one default preflight; no extra switch is required.

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

Running Python by itself applies deterministic formatting, statistical-presentation normalization, equation detection, and preservation checks. It does not replace the AI's judgment about hierarchy, analysis methods, equation meaning, object purpose, or page appearance.

Before maintenance or release, run:

```sh
python3 -m unittest discover -s tests -p 'test_*.py' -v
python3 build_skill.py
```

### Important Boundaries

- Specific university, instructor, or journal requirements take priority over general APA formatting.
- The tool never invents author information, references, research data, statistical results, or equation content.
- It never searches for, guesses, or inserts a missing DOI/URL; links inside reference-manager fields are preserved for review.
- Reference checks never move entries or certify that authors, dates, titles, journals, or DOIs are factually correct.
- It never rounds or recomputes statistical values, changes significance, or automatically renumbers tables, figures, or equations.
- Citation matching, analysis methods, title case, complex tables, figure clarity, copyright, and final pagination may still require author review.
- This is an APA 7 formatting and review assistant, not an official APA certification tool.

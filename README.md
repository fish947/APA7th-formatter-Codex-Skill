# APA 7 Word Formatter

> AI understands the paper. Python formats it reliably. You receive one new Word document.

Version 0.14.0 · Student papers · Professional manuscripts · Statistics · Equations · Tables · Figures · References

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
├── apa7_results.py              从结构化统计结果生成 APA 表格和图
├── apa7_statistics.py           统计表达和公式检查
├── apa7_visuals.py              图片、图表和矢量导出
├── apa7_workflow.py             AI 审核与逐页检查流程
├── build_skill.py               自动同步可安装 Skill
├── requirements.txt             Python 依赖
├── tests/                       全部自动测试
├── benchmark/public-sources.json 公开测试来源、许可与哈希登记（不含论文原文）
├── tools/
│   ├── apa7_benchmark.py         可重复的论文基准与质量门槛
│   └── qa_demo.py               生成合成文档做视觉测试
├── skills/apa7-word-formatter/  可直接安装的 Skill
└── .github/workflows/tests.yml  GitHub 自动测试
```

根目录的五个 `apa7_*.py` 和 `tools/apa7_benchmark.py` 是唯一维护来源。`skills/.../scripts/` 是构建生成的安装副本，不需要手动修改；运行 `python3 build_skill.py` 会同步脚本并更新版本与完整性哈希。

### 这个 Skill 能做什么

一句话来说：**把一篇现有的 Word 论文整理成 APA 7 格式，并把不能安全自动修改的问题告诉你。**

开始前只需选择 `student`（学生论文）或 `professional`（专业／投稿论文）。AI 先读懂论文各部分，Python 再稳定地修改格式。完成后会得到一个新的 Word 文件，原稿不会被覆盖。

#### 默认一键完成

| 功能 | 用简单的话说 |
| --- | --- |
| 保护原稿 | 永远另存一个新文件，不覆盖原稿；论文文字、数据、公式、图片和文献管理器内容会受到保护 |
| 学生版与专业版 | 按你选择的论文类型处理标题页、页眉和页码；不会凭空编写姓名、学校、课程或作者信息 |
| 页面基本格式 | 调整页边距、字体、双倍行距、段落间距和页码等常见 APA 7 格式 |
| 识别论文结构 | 分清标题页、摘要、关键词、正文标题、普通正文、长引用、参考文献、附录和图表说明 |
| 正文与标题 | 调整正文缩进、对齐方式和五级标题格式，让整篇论文的层级更统一 |
| 参考文献 | 设置悬挂缩进，并提醒重复文献、明显的顺序问题和同作者同年份标记问题；不会擅自改写文献信息 |
| DOI 和网页链接 | 把参考文献中已经写完整的网址变成可点击链接；不会上网猜测或补写缺失的 DOI |
| 统计数据 | 识别 *p*、*t*、*F*、*M*、*SD*、效应量等常见写法，在不改变数字的情况下统一符号和空格 |
| 统计问题提醒 | 提醒可能错误的 `p = .000`、小数位数，以及可能缺少的自由度、效应量或置信区间；不会编造或重新计算结果 |
| 公式 | 保留 Word 公式，识别可能的纯文字公式，并检查独立公式的排版、编号和位置；不会改写公式内容 |
| 表格 | 自动整理普通数据表的横线、竖线和文字格式；复杂表格会保持原样并提醒检查 |
| 图片和图表 | 识别图片、Word 图表和矢量图；过宽图片会按比例缩小，不会拉伸，也不会把普通图片假装成矢量图 |
| 图、表和公式编号 | 提醒重复编号、跳号、顺序错误，以及正文提到但文档里找不到的图表或公式；不会擅自重新编号 |
| 图表标题与正文引用 | 检查图表旁边是否有编号和标题，也会提醒正文引用与参考文献之间可能不对应的地方 |
| 成品检查 | 保存后重新打开 Word 文件，并逐页检查分页、遮挡、文字溢出和图表位置 |
| 简短结果说明 | 最后只告诉你：改了什么、依据哪些 APA 来源、改了哪些位置、还有什么需要自己确认 |

#### 需要 AI 帮忙判断的内容

| 内容 | AI 会怎么做 |
| --- | --- |
| 这段话是什么 | 根据内容和上下文判断它是标题、正文、引用、参考文献还是说明，而不是只看字体是否加粗 |
| 表格能不能自动改 | 普通数据表可以自动处理；合并单元格很多、结构特殊的表格会保留并提醒你查看 |
| 图片是什么用途 | 判断它是论文图、图表、截图还是装饰图片，再决定是否缩放、提取或保留 |
| 统计报告是否完整 | 根据统计方法提醒可能缺少的内容，但不会猜测不存在的数据 |
| 公式属于哪一类 | 区分行内公式、单独一行的公式和普通文字；不确定时保持原样 |
| 无法确定的内容 | 不强行修改，保留原文，并在最后的“还要审核”中直接告诉你 |

#### 可选的额外功能

| 功能 | 适合什么时候用 | 默认状态 |
| --- | --- | --- |
| 导出矢量图 | 想单独取出原有矢量图，或把数据完整的简单柱形图、折线图和散点图另存为 SVG | 关闭 |
| 继续写作样式 | 排版后还要继续写论文，想在 Word 中直接选择已经设好的正文、标题或参考文献样式 | 关闭 |
| 保存网页版说明 | 想把修改说明另外保存成一个可打开的网页文件 | 关闭 |
| 从统计结果生成表格和图 | 把 CSV／JSON 中已经确认的数据做成 APA 表格，或从原始绘图数据生成柱形图、折线图和散点图；不会重算统计结果 | 按需 |
| 只检查不修改 | 想先知道论文有什么格式问题，不马上生成排版成品 | 按需 |
| 更换字体 | 学校、导师或期刊指定了另一种 APA 允许的字体 | 按需 |

“继续写作样式”不是最终排版必需功能。它的作用很简单：如果你还要继续写论文，可以在 Word 里直接选择已经设置好的正文、标题和参考文献格式。如果论文已经写完，就不用开启。

“从统计结果生成表格和图”使用作者提供的结构化数据：它会制作表格编号、斜体标题、简洁横线、表注、重复表头和数字列对齐，也能从数据生成清晰的柱形图、折线图或散点图。Word 中使用高分辨率预览图；需要时可另外导出由原始数据绘制的 SVG。工具不会运行统计检验、补写缺失结果、改变小数精度或把截图假装成矢量图。

格式规则来自 APA 官方网站，包括 [Paper Format](https://apastyle.apa.org/style-grammar-guidelines/paper-format)、[学生论文设置指南](https://www.apa.org/ed/precollege/psn/2020/09/apa-style-student-papers)、[DOIs and URLs](https://apastyle.apa.org/style-grammar-guidelines/references/dois-urls)、[Numbers and Statistics Guide](https://apastyle.apa.org/instructional-aids/numbers-statistics-guide.pdf)、[APA Research Transparency Standards](https://www.apa.org/pubs/journals/resources/standards-disclosures)、[Headings](https://apastyle.apa.org/style-grammar-guidelines/paper-format/headings)、[Table Setup](https://apastyle.apa.org/style-grammar-guidelines/tables-figures/tables)、[Figure Setup](https://apastyle.apa.org/style-grammar-guidelines/tables-figures/figures)、[Reference List Setup](https://apastyle.apa.org/style-grammar-guidelines/paper-format/reference-list)、[Citing Works With the Same Author and Date](https://apastyle.apa.org/style-grammar-guidelines/citations/basic-principles/same-year-author) 和 [Author-Date Citation System](https://apastyle.apa.org/style-grammar-guidelines/citations/basic-principles/author-date)。运行 `python3 apa7_format.py --sources` 可以查看项目保存的官方原文摘录和直接链接。

参考文献链接遵循 APA 7 的电子文档规则：完整 DOI／URL 保持可点击，但链接不必显示成蓝色下划线。工具只给原稿已经写出的完整地址增加链接，不会联网搜索 DOI，也不会修改文献管理器生成的域。

参考文献质量审核也不会自动移动或改写条目。它只对能清晰识别作者和年份的条目检查字母顺序，并提示重复文字、重复 DOI 和同作者同年后缀问题。非拉丁字母排序、文献类型、标题大小写、斜体位置和书目事实仍由 AI 与作者核对。

统计和公式处理遵守一个简单原则：工具可以把 `P=0.032` 规范为斜体 *p* `= .032`，但不会把 `p = .000` 擅自改成 `p < .001`，也不会舍入、重算、补写效应量或改变显著性。涉及研究结论的内容只会列入“还要审核”。

图、表和公式的一致性检查同样只做安全提示。它能识别 `Tables 1–3`、`Figure A1` 和 `Equation (2)` 等明确写法，但不会自动改编号或破坏 Word 交叉引用域；图片是论文图还是校徽等装饰对象，仍由 AI 结合页面判断。

v0.13 增加公开真实文档来源登记和哈希核验，并加强第三方 Word 文件兼容性。项目只保存来源、许可和校验值，不把测试论文原文提交到 GitHub。真实论文发现的问题会被缩小成不含原文的合成测试，成为永久回归测试。

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
python3 tools/apa7_benchmark.py generate benchmark/generated/smoke
python3 tools/apa7_benchmark.py run benchmark/generated/smoke/cases.json --output-dir benchmark/runs/smoke
python3 build_skill.py
```

#### 5. 大量论文测试与迭代

基准不会因为“成功生成了 Word”就判定通过。原稿、文字、数值、公式、引用域或媒体意外变化会直接失败；要求视觉检查的案例，在全部页面完成复核之前只会显示“待审核”。

真实论文使用 `prepare-case` 建立私有案例，原文不会被复制到仓库：

```sh
python3 tools/apa7_benchmark.py prepare-case "/private/paper.docx" --id case_001 --profile student --output-dir benchmark/private/case_001
```

公开测试样本的下载地址、开放许可、文件大小和校验值登记在 `benchmark/public-sources.json`。论文文件仍放在被 Git 忽略的 `benchmark/private/` 中。下载后可核对文件是否与登记版本完全一致：

```sh
python3 tools/apa7_benchmark.py verify-sources benchmark/public-sources.json --corpus-dir benchmark/private/open-corpus
```

一条命令扫描整个测试库并生成汇总报告：

```sh
python3 tools/apa7_benchmark.py inspect-sources benchmark/public-sources.json --corpus-dir benchmark/private/open-corpus --profile professional --output-dir benchmark/private/corpus-inspection
```

当前基准库已登记并校验 **20 份 CC BY 4.0 Word 学术文档**。20 份文档都已完成只读结构扫描，共覆盖 5,532 个段落、44 个表格、173 个图形对象、65 个含原生公式的段落和 162 个统计表达。这里的“结构扫描通过”表示工具能安全读取并建立审核草稿，不表示每篇论文已经完成 APA 7 排版和逐页审核。目前已有 6 份文档完成完整排版、渲染和逐页复核：其中 4 份通过全部门槛，另外 2 份发现的版面问题已用于建立回归测试。20 份适合作为第一阶段稳定性基线，后续仍应继续补充不同学校模板、语言、学科和复杂版式。

发现问题后，先制作不含私人内容的最小合成文档和失败测试，再修改代码。这样同一问题以后不会悄悄回来。详细流程见 Skill 中的 `references/benchmarking.md`。

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
├── apa7_results.py              APA tables and figures from structured results
├── apa7_statistics.py           Statistical and equation checks
├── apa7_visuals.py              Visual inventory and vector export
├── apa7_workflow.py             AI review and page-QA workflow
├── build_skill.py               Synchronizes the installable Skill
├── requirements.txt             Python dependencies
├── tests/                       Automated regression tests
├── benchmark/public-sources.json Public source, license, and checksum registry; no papers
├── tools/
│   ├── apa7_benchmark.py         Repeatable corpus benchmark and quality gate
│   └── qa_demo.py               Synthetic visual-QA fixture generator
├── skills/apa7-word-formatter/  Installable Skill package
└── .github/workflows/tests.yml  GitHub test workflow
```

The five root `apa7_*.py` files and `tools/apa7_benchmark.py` are the maintained code sources. Files under `skills/.../scripts/` are generated installation copies. Run `python3 build_skill.py` to synchronize them and refresh the version and integrity hashes.

### What the Skill Can Do

In simple terms, it **formats an existing Word paper for APA 7 and tells you which issues still need human review**.

First choose `student` or `professional`. The AI identifies the parts of the paper, and Python applies the formatting consistently. You receive one new Word file; the original is never overwritten.

#### Automatic by Default

| Feature | In plain language |
| --- | --- |
| Protects the original | Always saves a new file instead of overwriting the original. It protects the paper's wording, numbers, equations, images, and reference-manager content |
| Student and professional modes | Formats the title page, header, and page numbers for the selected paper type. It never invents names, institutions, course details, or author information |
| Basic page formatting | Applies common APA 7 settings such as margins, an allowed font, double spacing, paragraph spacing, and page numbers |
| Understands paper sections | Distinguishes the title page, abstract, keywords, headings, body text, long quotations, references, appendices, and figure or table notes |
| Formats body text and headings | Adjusts paragraph indents, alignment, and the five APA heading levels so the paper has a consistent structure |
| Checks references | Applies hanging indents and flags duplicate entries, obvious ordering problems, and same-author/same-year labels without rewriting source details |
| Makes DOI and web links clickable | Turns complete URLs already written in the reference list into Word hyperlinks. It does not search for or guess a missing DOI |
| Formats statistical results | Recognizes common symbols such as *p*, *t*, *F*, *M*, and *SD*, then fixes safe spacing and symbol formatting without changing any number |
| Flags missing statistical details | Warns about issues such as `p = .000`, unusual decimal precision, or possibly missing degrees of freedom, effect sizes, and confidence intervals. It never invents or recalculates results |
| Preserves equations | Keeps native Word equations, identifies likely plain-text equations, and checks the layout, numbering, and placement of display equations without rewriting their content |
| Formats tables | Cleans up ordinary data tables with APA-style rules and text formatting. Complex tables are preserved and marked for review |
| Handles figures and charts | Identifies images, Word charts, and vector graphics. Oversized images are scaled proportionally and are never stretched or falsely labeled as vector graphics |
| Checks numbering | Flags duplicate, skipped, or out-of-order table, figure, and equation numbers, plus items mentioned in the text but not found in the paper. It never renumbers content on its own |
| Checks captions and citations | Looks for matching figure/table numbers and titles, and flags likely mismatches between in-text citations and the reference list |
| Reviews the finished file | Reopens the saved Word document and checks every page for bad page breaks, overlap, clipped text, and misplaced figures or tables |
| Gives a short report | Tells you only what changed, which APA sources were used, where the changes were made, and what still needs your review |

#### What the AI Helps Decide

| Content | What the AI does |
| --- | --- |
| What a paragraph means | Uses its wording and surrounding content to decide whether it is a heading, body paragraph, quotation, reference, or note instead of relying only on bold text |
| Whether a table is safe to change | Formats ordinary data tables, but leaves heavily merged or unusual tables unchanged for review |
| What an image is used for | Decides whether an object is a research figure, chart, screenshot, or decorative image before changing its size or exporting it |
| Whether statistical reporting is complete | Uses the type of analysis to point out possibly missing information without guessing nonexistent results |
| What kind of equation it is | Distinguishes inline math, display equations, and ordinary text; uncertain content remains unchanged |
| Anything uncertain | Keeps the original content and lists the question clearly under items that still need review |

#### Optional Extras

| Feature | When to use it | Default |
| --- | --- | --- |
| Vector export | Extract an existing vector image or save a supported simple bar, line, or scatter chart as SVG when complete chart data are available | Off |
| Continued-writing styles | Keep writing in the formatted document and select ready-made Word styles for body text, headings, references, or captions | Off |
| Web-page report | Save the short change report as a separate HTML file | Off |
| Results to APA tables and figures | Turn confirmed CSV/JSON values into APA tables, or create bar, line, and scatter figures from supplied plotting data without recalculating results | On request |
| Inspect without editing | See the paper's structure and likely problems before creating a formatted copy | On request |
| Change the font | Use another APA-supported font required by a university, instructor, or journal | On request |

Continued-writing styles are not required for final formatting. They only add reusable styles such as `APA7 Body`, `APA7 Heading 1`, and `APA7 Reference` to Word, so new content added later can reuse the correct formatting. If the paper is already complete, leave this option off.

The results generator uses structured data supplied by the author. It creates table/figure numbers, italic titles, minimal table rules, notes, repeating headers, and aligned numeric columns. It can also draw bar, line, and scatter figures, inserting a high-resolution preview into Word and optionally exporting a genuine data-derived SVG. It does not perform statistical tests, fill missing results, change precision, or claim that a screenshot has become a vector graphic.

Formatting rules are linked to official APA pages, including [Paper Format](https://apastyle.apa.org/style-grammar-guidelines/paper-format), the [student-paper setup guide](https://www.apa.org/ed/precollege/psn/2020/09/apa-style-student-papers), [DOIs and URLs](https://apastyle.apa.org/style-grammar-guidelines/references/dois-urls), the [Numbers and Statistics Guide](https://apastyle.apa.org/instructional-aids/numbers-statistics-guide.pdf), [APA Research Transparency Standards](https://www.apa.org/pubs/journals/resources/standards-disclosures), [Headings](https://apastyle.apa.org/style-grammar-guidelines/paper-format/headings), [Table Setup](https://apastyle.apa.org/style-grammar-guidelines/tables-figures/tables), [Figure Setup](https://apastyle.apa.org/style-grammar-guidelines/tables-figures/figures), [Reference List Setup](https://apastyle.apa.org/style-grammar-guidelines/paper-format/reference-list), [Citing Works With the Same Author and Date](https://apastyle.apa.org/style-grammar-guidelines/citations/basic-principles/same-year-author), and the [Author-Date Citation System](https://apastyle.apa.org/style-grammar-guidelines/citations/basic-principles/author-date). Run `python3 apa7_format.py --sources` to view the stored short quotations and direct links.

For electronic documents, complete DOI/URL text remains clickable, although APA permits either ordinary black text or the word processor's blue-underlined appearance. The formatter only links addresses already written in the manuscript; it does not search for DOIs or modify reference-manager fields.

The reference-quality audit is deliberately conservative. It checks alphabetical order only when author and year patterns are clear, and it reports duplicate text, duplicate DOIs, and same-author/same-year suffix problems without moving or rewriting entries. Non-Latin collation, source type, title capitalization, italics, and bibliographic facts still require AI and author review.

The statistics and equation formatter follows a strict boundary: it may normalize `P=0.032` to italic *p* `= .032`, but it will not silently change `p = .000` into `p < .001`, round or recompute a value, invent an effect size, or change significance. Anything that could affect the scientific result remains an author-review item.

The numbering check is also review-only. It recognizes explicit forms such as `Tables 1–3`, `Figure A1`, and `Equation (2)`, but never renumbers objects or rewrites Word cross-reference fields. The AI still decides whether a drawing is a research figure or a decorative object such as a logo.

Version 0.13 adds a public real-document source registry with checksum verification and improves compatibility with third-party Word packages. Only source, license, and integrity metadata are committed; the test papers remain outside Git. Each confirmed real-document defect is reduced to a fictional regression fixture.

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

Results to APA tables and figures:

> Use $apa7-word-formatter to turn this CSV or JSON results file into APA 7 tables and figures. Preserve every reported value and export data-derived SVG figures.

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

Create an APA table directly from CSV:

```sh
python3 apa7_results.py results.csv --output results_APA7.docx --table-title "Descriptive Statistics by Condition"
```

Create several tables and data-derived figures from JSON, with optional SVG export:

```sh
python3 apa7_results.py results.json --output results_APA7.docx --export-svg results_svg
```

The JSON schema and a small example are documented in `skills/apa7-word-formatter/references/results-generation.md`.

Running Python by itself applies deterministic formatting, statistical-presentation normalization, equation detection, and preservation checks. It does not replace the AI's judgment about hierarchy, analysis methods, equation meaning, object purpose, or page appearance.

Before maintenance or release, run:

```sh
python3 -m unittest discover -s tests -p 'test_*.py' -v
python3 tools/apa7_benchmark.py generate benchmark/generated/smoke
python3 tools/apa7_benchmark.py run benchmark/generated/smoke/cases.json --output-dir benchmark/runs/smoke
python3 build_skill.py
```

#### 5. Test a corpus and iterate safely

The benchmark does not pass a case merely because a Word file was created. Any unexpected change to the source, wording, numeric values, equations, citation fields, or media is a failure. A case that requires page review remains pending until every rendered page has a fresh review record.

Create a private real-paper case without copying the manuscript into the repository:

```sh
python3 tools/apa7_benchmark.py prepare-case "/private/paper.docx" --id case_001 --profile student --output-dir benchmark/private/case_001
```

`benchmark/public-sources.json` records the download URL, open license, file size, and checksums for public held-out cases. Keep the downloaded DOCX files under the ignored `benchmark/private/` folder, then verify them with:

```sh
python3 tools/apa7_benchmark.py verify-sources benchmark/public-sources.json --corpus-dir benchmark/private/open-corpus
```

Inspect the complete corpus and create a compact aggregate report with one command:

```sh
python3 tools/apa7_benchmark.py inspect-sources benchmark/public-sources.json --corpus-dir benchmark/private/open-corpus --profile professional --output-dir benchmark/private/corpus-inspection
```

The current corpus registers and verifies **20 CC BY 4.0 scholarly Word documents**. All 20 have completed read-only structure inspection, covering 5,532 paragraphs, 44 tables, 173 drawing objects, 65 paragraphs containing native equations, and 162 statistical expressions. “Structure inspection passed” means the tool could safely read the document and prepare a review draft; it does not mean every paper has completed APA 7 formatting and page-by-page review. Six documents have now completed the full format, render, and page-review cycle: four passed every gate, while layout issues found in two earlier documents were converted into regression tests. Twenty documents are a useful first stability baseline, while future releases should continue adding university templates, languages, disciplines, and difficult layouts.

When a real paper reveals a defect, reduce it to a fictional minimal DOCX, add a failing regression test, and only then change the formatter. See `references/benchmarking.md` in the Skill package for the complete workflow.

### Important Boundaries

- Specific university, instructor, or journal requirements take priority over general APA formatting.
- The tool never invents author information, references, research data, statistical results, or equation content.
- It never searches for, guesses, or inserts a missing DOI/URL; links inside reference-manager fields are preserved for review.
- Reference checks never move entries or certify that authors, dates, titles, journals, or DOIs are factually correct.
- It never rounds or recomputes statistical values, changes significance, or automatically renumbers tables, figures, or equations.
- Citation matching, analysis methods, title case, complex tables, figure clarity, copyright, and final pagination may still require author review.
- This is an APA 7 formatting and review assistant, not an official APA certification tool.

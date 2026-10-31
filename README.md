# 纸译 zhiyi

把一篇英文论文 PDF 变成一份可交付的中文译稿 `.docx`：图表按原文编号截图插入、公式是 Word 原生公式（可编辑）并连续编号、术语全局统一并在首次出现处注明英文、文末自动生成中英文术语对照表。

## 安装

```powershell
cd zhiyi
pip install -e .
```

不装也能用：在工作区根目录运行 `python -m zhiyi ...`。

外部依赖：

- **pandoc**：LaTeX → OMML。优先用环境变量 `ZHIYI_PANDOC`，其次在 `PATH` 中查找，最后回退到 Windows 默认路径。
- **Microsoft Word**：仅 `verify` 子命令需要（docx → PDF → 预览图）。

## 五步流水线

```powershell
# 1 抽正文（按页分节）+ 布局探测（图/表标题坐标）
zhiyi extract paper.pdf -o work --layout

# 2 定位图表：打印每页的 Figure/Table 标题坐标与图像块
zhiyi probe paper.pdf

# 3 裁图：写 crops.toml，一次裁完并生成对照图
zhiyi crop crops.toml -o work/figs

# 4 排版：写 content_xx.py（见下），生成 docx
zhiyi build content_xx.py -o out/论文中文译本.docx

# 5 校验：导出 PDF 并生成整册预览图，一眼看版式
zhiyi verify out/论文中文译本.docx
```

## crops.toml

```toml
[[figure]]
name  = "pp_fig2"          # 输出 work/figs/pp_fig2.png
pdf   = "paper.pdf"
page  = 3
rect  = [46, 58, 568, 229] # PDF 坐标（点，左上角原点）
cut   = "Figure 2."        # 可选：裁掉底部以该文字开头的图注行
# auto_top = true          # 可选：上边界自动避开上一段正文
# dpi  = 450               # 默认 450
```

坐标怎么来：`zhiyi probe` 给出标题块的 `[x0,y0,x1,y1]`，图通常在标题上方、表在标题下方；先按区间粗裁，再看 `sheet_*.png` 对照图微调。

## content_xx.py

一个普通 Python 文件，暴露 `OPS`（元组列表），`zhiyi build` 按顺序渲染。

| op | 用法 |
|---|---|
| `("title", 中文, 英文)` | 标题 + 英文原题 |
| `("meta", [行1, 行2])` | 作者 / 单位 / 版本行 |
| `("note", 文本)` | 灰底小字的译者说明、原文脚注 |
| `("h1"/"h2"/"h3", 文本)` | 三级标题（黑体） |
| `("p", 文本)` | 正文段，首行缩进 2 字符 |
| `("bullet", 文本)` | 项目符号行 |
| `("eq", latex, "1.1")` | 独立公式，居中，右端编号 (1.1) |
| `("fig", png, 题注, 宽cm)` | 插图 + 中文题注 |
| `("tab", png, 题注, 宽cm)` | 表格截图 + 中文题注 |
| `("eqtable", 题注, [(步骤, latex), ...])` | 算法表（真表格 + 原生公式） |
| `("figterms", "English 中文；…")` | 图/表下方「图中英文标注对照」 |
| `("glossary", 标题, 说明)` | 文末术语对照表（自动收集本篇用到的术语） |
| `("pagebreak",)` / `("space", pt)` | 分页 / 留白 |

文本里支持三种标记：

- `$x_k = A x_{k-1}$` → 行内原生公式
- `**加粗**`
- `«pillar»` → 首次出现渲染成 `柱（pillar）`，之后只留 `柱`，并自动进入文末术语表

## 术语表

`zhiyi/glossary.json` 是唯一事实来源，三份译稿共用：

```json
"pillar": { "zh": "柱", "en": "pillar", "note": "…" },
"mota":   { "zh": "多目标跟踪准确度", "en": "multi-object tracking accuracy (MOTA)", "abbr": "MOTA" }
```

`abbr` 存在时，首次出现只注缩写：`多目标跟踪准确度（MOTA）`。

```powershell
zhiyi glossary list                       # 列出全部词条
zhiyi glossary check a.py b.py c.py       # 检查未知标记 & 一译多名冲突
```

项目专有术语放单独 JSON，用 `zhiyi build content.py --glossary extra.json` 合并。

## 为什么这么做

- 公式走 pandoc 的 LaTeX→OMML，产物是 Word 原生公式：可双击编辑、可搜索、缩放不失真，比截图清晰。同一公式被多处引用时必须 `deepcopy` 后再插入，否则 lxml 会把元素从上一处搬走、留下空位。
- 图表一律用原文 450 dpi 截图 + 中文题注，避免重画走形；题注统一「图 1　中文（原图 Figure 1）」以便与原文对照。
- 术语用标记而非手写括注，保证「同一英文词全文同一译名」是机制而非自觉。

## 目录

```
zhiyi/
  cli.py        子命令入口
  docx.py       排版引擎（DocBuilder / MathBank）
  pdf.py        抽取、布局探测、裁图、对照图
  terms.py      术语表读写
  glossary.json 共用术语库
  verify.py     Word 导出 PDF + 预览图
```

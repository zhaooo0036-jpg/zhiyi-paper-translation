---
name: zhiyi-paper-translation
description: 把英文论文 PDF 译成规范中文 docx（图表按原文编号截图、公式为原生 OMML 并编号、术语全局统一、文末附中英术语对照表）。当任务是把 PDF 文献译成中文并交付 Word 文档时使用。
---

# 纸译 zhiyi · 论文中译流水线

代码在 `zhiyi/`，命令 `python -m zhiyi`（或 `pip install -e zhiyi` 后用 `zhiyi`）。

## 何时用

用户要求「把 PDF/文献译成中文」「做成 Word 译稿」，并且在意图表、公式、术语一致性时。

## 执行顺序

1. `zhiyi extract paper.pdf -o work --layout` —— 先读 `work/*.txt` 判断文献结构（页数、章节、公式分布）。
2. `zhiyi probe paper.pdf` —— 打印 Figure/Table 标题坐标，据此写 `crops.toml`。
3. `zhiyi crop crops.toml -o work/figs` —— 生成 PNG 与 `sheet_*.png`；**必须看对照图**，确认没有把标题/正文裁进去。
4. 写 `content_xx.py`（op 列表见 README），术语一律用 `«key»` 标记；`«»` 里的 key 必须存在于 `glossary.json`，缺词先补词条。
5. `zhiyi build content_xx.py -o out/xxx.docx`
6. `zhiyi verify out/xxx.docx` —— 看整册预览图，检查页数、图表落位、公式是否有空位。

## 硬性要求

- 同一英文术语在全文、跨文档译名一致；首次出现注英文（机制自动完成，不要手写括注）。
- 图表保留原文编号，题注写「图 1　中文题注（原图 Figure 1）」。
- 结构图、曲线图、表格后必须有 `figterms`，逐词给出英文 → 汉语。
- 公式用 `(节.序)` 编号；原文无编号的按出现顺序补编，并在 `note` 里交代。
- 文末必须有 `("glossary", ...)`。
- 文风近直译，不添加原文没有的解释；确需补注时标「译者注」或「原文脚注」。

## 已知坑

- pandoc 缺失时 `build` 会在 `MathBank.build` 报错；用 `ZHIYI_PANDOC` 指路径。
- 同一 latex 字符串重复插入必须深拷贝，`DocBuilder` 已处理，勿绕过 `math.get()`。
- `crop` 的 `rect` 是 PDF 点坐标；`cut` 只裁掉「以该文字开头」的那一行，图注在下方时不要用。

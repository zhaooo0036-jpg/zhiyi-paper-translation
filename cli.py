# -*- coding: utf-8 -*-
"""zhiyi command line: PDF -> Chinese translation .docx, end to end."""
import argparse, importlib.util, io, json, os, re, shutil, sys, tempfile

from . import __version__, pdf as pdfmod, terms as termsmod


def _load_ops(path):
    """Import a content module (a plain .py exposing OPS) by file path."""
    path = os.path.abspath(path)
    spec = importlib.util.spec_from_file_location("zhiyi_content_%d" % abs(hash(path)), path)
    mod = importlib.util.module_from_spec(spec)
    sys.path.insert(0, os.path.dirname(path))
    spec.loader.exec_module(mod)
    if not hasattr(mod, "OPS"):
        raise SystemExit("%s does not define OPS" % path)
    return mod.OPS


def cmd_extract(a):
    for p in a.pdf:
        dst, n, chars = pdfmod.extract_text(p, a.out)
        print("%-46s pages=%-4d chars=%-7d -> %s" % (os.path.basename(p), n, chars, dst))
        if a.layout:
            lp, _ = pdfmod.probe(p, a.out)
            print("%46s layout -> %s" % ("", lp))


def cmd_probe(a):
    for p in a.pdf:
        lp, pages = pdfmod.probe(p, a.out)
        print("== %s  (%d pages)  ->  %s" % (os.path.basename(p), len(pages), lp))
        cap = re.compile(r"^(Figure|Fig\.|Table)\s*\d")
        for pg in pages:
            hits = [b for b in pg["blocks"] if cap.search(b["text"])]
            if not hits and not pg["images"]:
                continue
            print("  p%-3d images=%-3d" % (pg["page"], len(pg["images"])))
            for b in hits:
                print("     %-28s %s" % (str(b["bbox"]), b["text"][:70].replace("\n", " ")))


def cmd_crop(a):
    out = a.out or os.path.join(os.getcwd(), "figs")
    made, sheets = pdfmod.crop_spec(a.spec, out, dpi=a.dpi, sheet=not a.no_sheet)
    for m in made:
        print("%-14s %s  %.1f x %.1f cm" % (m["name"], m["file"], m["cm"][0], m["cm"][1]))
    for s in sheets:
        print("sheet ->", s)


def cmd_build(a):
    from .docx import DocBuilder
    if a.glossary:
        termsmod.override(a.glossary)
    ops = _load_ops(a.content)
    out = a.out or os.path.join(os.path.dirname(os.path.abspath(a.content)),
                                os.path.splitext(os.path.basename(a.content))[0] + ".docx")
    b = DocBuilder(os.path.dirname(os.path.abspath(out)), body_size=a.body_size,
                   line_spacing=a.line_spacing, margins=a.margins)
    b.render(ops)
    b.save(out)
    used = len(b.used)
    print("built %s  (%d ops, %d glossary terms)" % (out, len(ops), used))
    return out


def cmd_verify(a):
    from .verify import docx_to_pdf, contact_sheet
    pdf_path = a.out or os.path.splitext(a.docx)[0] + ".pdf"
    docx_to_pdf(a.docx, pdf_path)
    png, n = contact_sheet(pdf_path, a.png or os.path.splitext(a.docx)[0] + ".png",
                           pages=a.pages, tile_w=a.tile)
    print("pages=%d  pdf=%s  preview=%s" % (n, pdf_path, png))


def cmd_glossary(a):
    if a.action == "list":
        for k, v in sorted(termsmod.TERMS.items(), key=lambda kv: kv[1][1].lower()):
            note = v[2] if len(v) > 2 else ""
            print("%-34s %-16s %s" % (v[1], v[0], note))
        print("\n%d entries" % len(termsmod.TERMS))
        return
    bad = 0
    for path in a.content:
        src = io.open(path, encoding="utf-8").read()
        used = sorted(set(re.findall(r"«([a-z_0-9]+)»", src)))
        unknown = [k for k in used if k not in termsmod.TERMS]
        dup = {}
        for k in used:
            zh = termsmod.TERMS[k][0]
            dup.setdefault(zh, []).append(k)
        clash = {z: ks for z, ks in dup.items() if len(ks) > 1}
        print("%-32s terms=%-3d unknown=%s" % (os.path.basename(path), len(used), unknown or "-"))
        for z, ks in clash.items():
            print("   ! two keys share one Chinese name: %s <- %s" % (z, ks))
            bad += 1
        bad += len(unknown)
    raise SystemExit(1 if bad else 0)


def cmd_init(a):
    d = a.dir
    os.makedirs(d, exist_ok=True)
    content = os.path.join(d, "content_%s.py" % a.slug)
    if not os.path.exists(content) or a.force:
        io.open(content, "w", encoding="utf-8", newline="\n").write(_CONTENT_TEMPLATE.format(slug=a.slug))
    spec = os.path.join(d, "crops.toml")
    if not os.path.exists(spec) or a.force:
        io.open(spec, "w", encoding="utf-8", newline="\n").write(_CROP_TEMPLATE)
    print("scaffolded:\n  %s\n  %s" % (content, spec))


_CONTENT_TEMPLATE = '''# -*- coding: utf-8 -*-
"""Content ops for one paper.  See `zhiyi init` / README for the op list."""
import os

F = os.path.join(os.path.dirname(os.path.abspath(__file__)), "figs")
def fig(n):
    return os.path.join(F, n + ".png")

OPS = [
    ("title", "中文标题", "Original English Title"),
    ("meta", ["Author A，Author B", "单位"]),
    ("note", "**译者说明**：本文为原文的中文译本……首次出现术语以括号注明英文，文末附术语对照表。"),
    ("h1", "摘要"),
    ("p", "第一段。术语用 «key» 标记，首次出现渲染为「中文（English）」。公式用 $x_k$ 行内、(\"eq\", r\"...\", \"1\") 独立编号。"),
    ("tab", fig("{slug}_tab1"), "表 1　题注（原表 Table 1）", 15.8),
    ("figterms", "Method 方法；Speed 速度"),
    ("glossary", "附录　中英文术语对照表", "按英文字母顺序排列。"),
]
'''

_CROP_TEMPLATE = '''# crop one figure/table per [[figure]] block; rect = [x0, y0, x1, y1] in PDF points
[[figure]]
name  = "paper_fig1"
pdf   = "paper.pdf"
page  = 1
rect  = [46, 58, 568, 229]
# cut   = "Figure 1."      # optional: trim bottom above this caption line
# auto_top = true          # optional: start below the previous prose block
# dpi  = 450
'''


def main(argv=None):
    p = argparse.ArgumentParser(prog="zhiyi", description="纸译 · PDF → 中文译稿 (.docx) 流水线")
    p.add_argument("--version", action="version", version="zhiyi " + __version__)
    sub = p.add_subparsers(dest="cmd", required=True)

    e = sub.add_parser("extract", help="抽取 PDF 文本（按页分节）")
    e.add_argument("pdf", nargs="+")
    e.add_argument("-o", "--out")
    e.add_argument("--layout", action="store_true", help="同时输出布局探测 layout.json")
    e.set_defaults(func=cmd_extract)

    q = sub.add_parser("probe", help="定位图/表标题与图像块坐标")
    q.add_argument("pdf", nargs="+")
    q.add_argument("-o", "--out")
    q.set_defaults(func=cmd_probe)

    c = sub.add_parser("crop", help="按 crops.toml 裁图，并生成对照图")
    c.add_argument("spec")
    c.add_argument("-o", "--out")
    c.add_argument("--dpi", type=int, default=450)
    c.add_argument("--no-sheet", action="store_true")
    c.set_defaults(func=cmd_crop)

    b = sub.add_parser("build", help="由 content .py 生成 docx")
    b.add_argument("content")
    b.add_argument("-o", "--out")
    b.add_argument("--glossary", help="额外的术语表 JSON（与内置表合并）")
    b.add_argument("--body-size", type=float, default=12.0)
    b.add_argument("--line-spacing", type=float, default=1.5)
    b.add_argument("--margins", type=float, default=2.6)
    b.set_defaults(func=cmd_build)

    v = sub.add_parser("verify", help="docx → pdf → 预览图（需本机 Word）")
    v.add_argument("docx")
    v.add_argument("-o", "--out")
    v.add_argument("--png")
    v.add_argument("--pages", type=int, nargs="*")
    v.add_argument("--tile", type=int, default=560)
    v.set_defaults(func=cmd_verify)

    g = sub.add_parser("glossary", help="术语表：列出 / 检查一致性")
    g.add_argument("action", choices=["list", "check"])
    g.add_argument("content", nargs="*")
    g.set_defaults(func=cmd_glossary)

    i = sub.add_parser("init", help="生成 content 模块与 crops.toml 骨架")
    i.add_argument("dir")
    i.add_argument("--slug", default="paper")
    i.add_argument("--force", action="store_true")
    i.set_defaults(func=cmd_init)

    a = p.parse_args(argv)
    if a.cmd == "glossary" and a.action == "check" and not a.content:
        raise SystemExit("glossary check needs at least one content .py")
    a.func(a)
    return 0

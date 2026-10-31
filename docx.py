# -*- coding: utf-8 -*-
"""Shared docx builder: Chinese academic layout + OMML equations via pandoc + captions.

Content is described by a list of ops and rendered in one pass:
    ("title", cn, en)          ("meta", [line, ...])      ("h1"|"h2"|"h3", text)
    ("p", text)                ("note", text)             ("bullet", text)
    ("eq", tex, "1.1")         ("fig"|"tab", png, caption, width_cm)
    ("pagebreak",)             ("space", points)
Text may contain $inline math$ and **bold**.
"""
import copy, io, os, re, shutil, subprocess
import docx
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from PIL import Image

from .terms import TERMS

PANDOC = (os.environ.get("ZHIYI_PANDOC")
          or shutil.which("pandoc")
          or r"F:\Anaconda\Scripts\pandoc.exe")

BODY_EA, BODY_LATIN = "宋体", "Times New Roman"
HEAD_EA = "黑体"
CAP_EA = "宋体"


def norm(tex):
    return " ".join(tex.split())


class MathBank:
    def __init__(self, workdir):
        self.workdir = workdir
        self.tex, self.index, self.omml = [], {}, {}

    def add(self, tex):
        tex = norm(tex)
        if tex not in self.index:
            self.index[tex] = len(self.tex)
            self.tex.append(tex)
        return tex

    def build(self):
        if not self.tex:
            return
        md = os.path.join(self.workdir, "_eq.md")
        dx = os.path.join(self.workdir, "_eq.docx")
        with io.open(md, "w", encoding="utf-8") as fh:
            for t in self.tex:
                fh.write("$$" + t + "$$\n\n")
        subprocess.run([PANDOC, md, "-o", dx, "--from", "markdown", "--to", "docx"],
                       check=True, cwd=self.workdir,
                       stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        d = docx.Document(dx)
        found = []
        for p in d.paragraphs:
            els = list(p._p.iter(qn("m:oMath")))
            if els:
                found.append(els[0])
        if len(found) != len(self.tex):
            raise RuntimeError("equation count mismatch: %d vs %d" % (len(found), len(self.tex)))
        for t, el in zip(self.tex, found):
            self.omml[t] = copy.deepcopy(el)

    def get(self, tex):
        # always hand out a fresh copy: an lxml element can only live in one place,
        # so re-using the banked element would silently empty the earlier insertion.
        return copy.deepcopy(self.omml[norm(tex)])


def _set_run_font(run, size=12, bold=False, ea=BODY_EA, latin=BODY_LATIN,
                  italic=False, color=None, spacing=None):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = latin
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    rf.set(qn("w:ascii"), latin)
    rf.set(qn("w:hAnsi"), latin)
    rf.set(qn("w:eastAsia"), ea)
    if color:
        run.font.color.rgb = RGBColor(*color)
    if spacing:
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:val"), str(int(spacing * 20)))
        rpr.append(sp)


class DocBuilder:
    def __init__(self, workdir, body_size=12.0, line_spacing=1.5, margins=2.6):
        self.workdir = workdir
        self.body_size = body_size
        self.line_spacing = line_spacing
        self.math = MathBank(workdir)
        self.seen = set()
        self.used = []
        self.doc = Document()
        sec = self.doc.sections[0]
        sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
        sec.left_margin = sec.right_margin = Cm(margins)
        sec.top_margin = Cm(2.5)
        sec.bottom_margin = Cm(2.4)
        self.text_width_cm = 21.0 - 2 * margins
        st = self.doc.styles["Normal"]
        st.font.name = BODY_LATIN
        st.font.size = Pt(body_size)
        rpr = st.element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rpr.insert(0, rf)
        rf.set(qn("w:eastAsia"), BODY_EA)
        self._footer()

    def _footer(self):
        p = self.doc.sections[0].footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run("— ")
        _set_run_font(r, size=9)
        fld = OxmlElement("w:fldSimple")
        fld.set(qn("w:instr"), "PAGE")
        rr = OxmlElement("w:r")
        rpr = OxmlElement("w:rPr")
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), "18")
        rpr.append(sz)
        rr.append(rpr)
        t = OxmlElement("w:t")
        t.text = "1"
        rr.append(t)
        fld.append(rr)
        p._p.append(fld)
        r2 = p.add_run(" —")
        _set_run_font(r2, size=9)

    # ---------------------------------------------------------- text
    def expand(self, text):
        """«key» -> 中文（English） on first use in this document, 中文 afterwards."""
        def repl(m):
            key = m.group(1)
            if key not in TERMS:
                raise KeyError("unknown term marker: %s" % key)
            entry = TERMS[key]
            zh, en = entry[0], entry[1]
            abbr = entry[3] if len(entry) > 3 else None
            if key not in self.seen:
                self.seen.add(key)
                self.used.append(key)
                if zh == en:
                    return zh
                if abbr:
                    return "%s（%s）" % (zh, abbr)
                return "%s（%s）" % (zh, en)
            return zh
        return re.sub(r"«([a-z_0-9]+)»", repl, text)

    def _emit(self, p, text, size, color=None, bold_all=False):
        for chunk in re.split(r"(\$[^$]+\$|\*\*[^*]+\*\*)", text):
            if not chunk:
                continue
            if chunk.startswith("$") and chunk.endswith("$") and len(chunk) > 2:
                p._p.append(self.math.get(chunk[1:-1]))
            elif chunk.startswith("**") and chunk.endswith("**"):
                r = p.add_run(self.expand(chunk[2:-2]))
                _set_run_font(r, size=size, bold=True, ea=BODY_EA, color=color)
            else:
                r = p.add_run(self.expand(chunk))
                _set_run_font(r, size=size, bold=bold_all, ea=BODY_EA, color=color)

    def figterms(self, text):
        p = self._p(WD_ALIGN_PARAGRAPH.JUSTIFY, 0, 9, indent=None, ls=1.2)
        p.paragraph_format.left_indent = Cm(0.8)
        p.paragraph_format.right_indent = Cm(0.8)
        r = p.add_run("图中英文标注对照：")
        _set_run_font(r, size=9.5, bold=True, ea=BODY_EA, color=(0x44, 0x44, 0x44))
        self._emit(p, text, 9.5, color=(0x44, 0x44, 0x44))

    def glossary(self, title="附录　中英文术语对照表", note=None):
        self.pagebreak()
        self.h1(title)
        if note:
            self.para(note, indent=0, size=self.body_size - 1, sa=8)
        keys = sorted(self.used, key=lambda k: TERMS[k][1].lower())
        rows = [("英文术语", "中文译名", "说明")]
        for k in keys:
            entry = TERMS[k]
            zh, en = entry[0], entry[1]
            nt = entry[2] if len(entry) > 2 else ""
            rows.append((en, zh, nt))
        tbl = self.doc.add_table(rows=len(rows), cols=3)
        tbl.style = "Table Grid"
        tbl.autofit = False
        widths = [Cm(6.6), Cm(4.8), Cm(4.4)]
        for i, row in enumerate(rows):
            for j, val in enumerate(row):
                cell = tbl.cell(i, j)
                cell.width = widths[j]
                p = cell.paragraphs[0]
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.0
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j == 2 else WD_ALIGN_PARAGRAPH.LEFT
                r = p.add_run(val)
                _set_run_font(r, size=9.5, bold=(i == 0), ea=BODY_EA)

    def eqtable(self, caption, rows, widths=(5.4, 10.2), size=10.0):
        """Two-column table of (label, latex) rows, used for algorithm tables."""
        self.caption(caption, sb=8, sa=4)
        tbl = self.doc.add_table(rows=len(rows), cols=2)
        tbl.style = "Table Grid"
        tbl.autofit = False
        for i, (label, tex) in enumerate(rows):
            c0 = tbl.cell(i, 0)
            c0.width = Cm(widths[0])
            p0 = c0.paragraphs[0]
            p0.paragraph_format.space_before = Pt(3)
            p0.paragraph_format.space_after = Pt(3)
            p0.paragraph_format.line_spacing = 1.1
            self._emit(p0, label, size)
            c1 = tbl.cell(i, 1)
            c1.width = Cm(widths[1])
            p1 = c1.paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p1.paragraph_format.space_before = Pt(3)
            p1.paragraph_format.space_after = Pt(3)
            p1.paragraph_format.line_spacing = 1.1
            p1._p.append(self.math.get(tex))
        self.space(8)

    def _p(self, align=WD_ALIGN_PARAGRAPH.JUSTIFY, sb=0, sa=6, indent=None, ls=None):
        p = self.doc.add_paragraph()
        pf = p.paragraph_format
        pf.alignment = align
        pf.space_before = Pt(sb)
        pf.space_after = Pt(sa)
        pf.line_spacing = self.line_spacing if ls is None else ls
        if indent:
            pf.first_line_indent = Pt(self.body_size * indent)
        return p

    def title(self, cn, en=None):
        p = self._p(WD_ALIGN_PARAGRAPH.CENTER, 6, 8, ls=1.35)
        r = p.add_run(cn)
        _set_run_font(r, size=19, bold=True, ea=HEAD_EA, spacing=0.5)
        if en:
            p2 = self._p(WD_ALIGN_PARAGRAPH.CENTER, 0, 12, ls=1.2)
            r2 = p2.add_run(en)
            _set_run_font(r2, size=11.5, color=(0x33, 0x33, 0x33))

    def meta(self, lines, size=10.5):
        for ln in lines:
            p = self._p(WD_ALIGN_PARAGRAPH.CENTER, 0, 2, ls=1.2)
            self._emit(p, ln, size)

    def h1(self, text):
        p = self._p(WD_ALIGN_PARAGRAPH.LEFT, 15, 6, ls=1.2)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        _set_run_font(r, size=14, bold=True, ea=HEAD_EA)

    def h2(self, text):
        p = self._p(WD_ALIGN_PARAGRAPH.LEFT, 11, 4, ls=1.2)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        _set_run_font(r, size=12.5, bold=True, ea=HEAD_EA)

    def h3(self, text):
        p = self._p(WD_ALIGN_PARAGRAPH.LEFT, 8, 3, ls=1.2)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        _set_run_font(r, size=12, bold=True, ea=HEAD_EA)

    def para(self, text, indent=2, size=None, align=WD_ALIGN_PARAGRAPH.JUSTIFY, sa=6):
        p = self._p(align, 0, sa, indent=indent)
        self._emit(p, text, size or self.body_size)
        return p

    def note(self, text, size=None):
        p = self._p(WD_ALIGN_PARAGRAPH.JUSTIFY, 0, 8, indent=None, ls=1.25)
        p.paragraph_format.left_indent = Cm(0.4)
        p.paragraph_format.right_indent = Cm(0.4)
        self._emit(p, text, size or (self.body_size - 2), color=(0x55, 0x55, 0x55))

    def bullet(self, text, size=None):
        p = self._p(WD_ALIGN_PARAGRAPH.JUSTIFY, 0, 3, indent=None, ls=1.45)
        p.paragraph_format.left_indent = Cm(0.9)
        p.paragraph_format.first_line_indent = Cm(-0.45)
        size = size or self.body_size
        r = p.add_run("• ")
        _set_run_font(r, size=size, ea=BODY_EA)
        self._emit(p, text, size)

    def caption(self, text, size=10.5, sb=3, sa=10, keep=False):
        p = self._p(WD_ALIGN_PARAGRAPH.CENTER, sb, sa, ls=1.15)
        if keep:
            p.paragraph_format.keep_with_next = False
        self._emit(p, text, size)

    def equation(self, tex, number=None):
        p = self._p(WD_ALIGN_PARAGRAPH.LEFT, 5, 5, ls=1.15)
        pf = p.paragraph_format
        pf.tab_stops.add_tab_stop(Cm(self.text_width_cm / 2.0), WD_TAB_ALIGNMENT.CENTER)
        pf.tab_stops.add_tab_stop(Cm(self.text_width_cm), WD_TAB_ALIGNMENT.RIGHT)
        r = p.add_run("\t")
        _set_run_font(r, size=self.body_size)
        p._p.append(self.math.get(tex))
        if number:
            r2 = p.add_run("\t(%s)" % number)
            _set_run_font(r2, size=self.body_size)

    def figure(self, png, caption, width_cm=None, max_fraction=1.0, sa=10):
        if not os.path.exists(png):
            raise FileNotFoundError(png)
        with Image.open(png) as im:
            w, h = im.size
        natural_cm = w / 300.0 * 2.54
        target = width_cm or natural_cm
        target = min(target, self.text_width_cm * max_fraction)
        p = self._p(WD_ALIGN_PARAGRAPH.CENTER, 8, 2, ls=1.0)
        p.paragraph_format.keep_with_next = True
        p.add_run().add_picture(png, width=Cm(target))
        self.caption(caption, sb=3, sa=sa)

    def pagebreak(self):
        self.doc.add_page_break()

    def space(self, pts=10):
        p = self._p(WD_ALIGN_PARAGRAPH.LEFT, 0, 0, ls=1.0)
        p.paragraph_format.space_after = Pt(pts)

    # ---------------------------------------------------------- render
    def scan(self, ops):
        def walk(arg):
            if isinstance(arg, str):
                for m in re.findall(r"\$([^$]+)\$", arg):
                    self.math.add(m)
            elif isinstance(arg, (list, tuple)):
                for a in arg:
                    walk(a)
        for op in ops:
            if op[0] == "eq":
                self.math.add(op[1])
            elif op[0] == "eqtable":
                for _label, tex in op[2]:
                    self.math.add(tex)
            elif op[0] != "title":
                walk(op[1:])

    def render(self, ops):
        self.scan(ops)
        self.math.build()
        for op in ops:
            kind, args = op[0], op[1:]
            if kind == "title":
                self.title(*args)
            elif kind == "meta":
                self.meta(args[0])
            elif kind in ("h1", "h2", "h3"):
                getattr(self, kind)(*args)
            elif kind == "p":
                self.para(*args)
            elif kind == "note":
                self.note(*args)
            elif kind == "bullet":
                self.bullet(*args)
            elif kind == "eq":
                self.equation(*args)
            elif kind in ("fig", "tab"):
                self.figure(*args)
            elif kind == "caption":
                self.caption(*args)
            elif kind == "figterms":
                self.figterms(*args)
            elif kind == "eqtable":
                self.eqtable(*args)
            elif kind == "glossary":
                self.glossary(*args)
            elif kind == "pagebreak":
                self.pagebreak()
            elif kind == "space":
                self.space(*args)
            else:
                raise ValueError("unknown op: %r" % (op,))
        return self.doc

    def save(self, path):
        self.doc.save(path)
        return path

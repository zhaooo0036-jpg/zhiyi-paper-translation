# -*- coding: utf-8 -*-
"""PDF side: text extraction, layout probing, figure/table cropping."""
import io, json, os, re
import pymupdf
from PIL import Image, ImageDraw


# ------------------------------------------------------------------ extract
def extract_text(pdf, out_dir=None, dpi=None):
    """Write <stem>.txt (page-marked) and return (path, page_count, chars)."""
    out_dir = out_dir or os.path.dirname(os.path.abspath(pdf))
    os.makedirs(out_dir, exist_ok=True)
    doc = pymupdf.open(pdf)
    pages = []
    for i, page in enumerate(doc):
        pages.append(page.get_text("text"))
    body = "\n".join("\n===== PAGE %d =====\n%s" % (i + 1, t) for i, t in enumerate(pages))
    stem = os.path.splitext(os.path.basename(pdf))[0]
    dst = os.path.join(out_dir, _safe(stem) + ".txt")
    io.open(dst, "w", encoding="utf-8", errors="replace").write(body)
    return dst, len(doc), len(body)


def probe(pdf, out_dir=None):
    """Layout probe: per-page text blocks + raster image boxes (for locating figures)."""
    out_dir = out_dir or os.path.dirname(os.path.abspath(pdf))
    os.makedirs(out_dir, exist_ok=True)
    doc = pymupdf.open(pdf)
    stem = _safe(os.path.splitext(os.path.basename(pdf))[0])
    pages = []
    for i, page in enumerate(doc):
        blocks = []
        for b in page.get_text("dict")["blocks"]:
            if b["type"] != 0:
                continue
            txt = " ".join(s["text"] for l in b["lines"] for s in l["spans"]).strip()
            if txt:
                blocks.append({"bbox": [round(v, 1) for v in b["bbox"]], "text": txt})
        pages.append({"page": i + 1,
                      "rect": [round(v, 1) for v in page.rect],
                      "blocks": blocks,
                      "images": [[round(v, 1) for v in x["bbox"]] for x in page.get_image_info()]})
    dst = os.path.join(out_dir, stem + ".layout.json")
    io.open(dst, "w", encoding="utf-8").write(json.dumps(pages, ensure_ascii=False, indent=1))
    return dst, pages


# ------------------------------------------------------------------ crop
def _line_top(page, prefix, below):
    best = None
    for b in page.get_text("dict")["blocks"]:
        if b["type"] != 0:
            continue
        for l in b["lines"]:
            txt = "".join(s["text"] for s in l["spans"]).strip()
            if txt.startswith(prefix) and l["bbox"][1] < below:
                if best is None or l["bbox"][1] > best:
                    best = l["bbox"][1]
    return best


def _prose_bottom(page, below, min_len=120):
    best = 0.0
    for b in page.get_text("dict")["blocks"]:
        if b["type"] != 0:
            continue
        txt = " ".join(s["text"] for l in b["lines"] for s in l["spans"]).strip()
        if len(txt) >= min_len and b["bbox"][3] < below - 1:
            best = max(best, b["bbox"][3])
    return best


def crop_one(pdf, page_no, rect, out_png, dpi=450, cut=None, auto_top=False):
    """Crop `rect` (points, top-left origin). `cut` trims the bottom above a caption line."""
    doc = pymupdf.open(pdf)
    page = doc[page_no - 1]
    x0, y0, x1, y1 = [float(v) for v in rect]
    if cut:
        ly = _line_top(page, cut, y1 + 1)
        if ly:
            y1 = min(y1, ly - 2)
    if auto_top:
        y0 = max(y0, _prose_bottom(page, y1) + 5)
    pm = page.get_pixmap(dpi=dpi, clip=pymupdf.Rect(x0, y0, x1, y1))
    os.makedirs(os.path.dirname(os.path.abspath(out_png)), exist_ok=True)
    pm.save(out_png)
    return {"file": out_png, "rect": [x0, y0, x1, y1], "px": [pm.width, pm.height],
            "cm": [round(pm.width / dpi * 2.54, 1), round(pm.height / dpi * 2.54, 1)]}


def load_spec(path):
    """Accepts .toml (py3.11+) or .json crop specification."""
    if path.lower().endswith(".json"):
        return json.load(io.open(path, encoding="utf-8"))
    try:
        import tomllib
        with open(path, "rb") as fh:
            return tomllib.load(fh)
    except Exception:
        return json.load(io.open(path, encoding="utf-8"))


def crop_spec(spec_path, out_dir, dpi=450, sheet=True, tile=620):
    spec = load_spec(spec_path)
    items = spec.get("figure") or spec.get("figures") or []
    made, groups = [], {}
    for it in items:
        res = crop_one(it["pdf"], int(it["page"]), it["rect"],
                       os.path.join(out_dir, it["name"] + ".png"),
                       dpi=int(it.get("dpi", dpi)), cut=it.get("cut"),
                       auto_top=bool(it.get("auto_top", False)))
        res["name"] = it["name"]
        res["group"] = it.get("group", "sheet")
        made.append(res)
        groups.setdefault(res["group"], []).append(res["name"])
    sheets = []
    if sheet:
        for g, names in groups.items():
            sheets.append(make_sheet([os.path.join(out_dir, n + ".png") for n in names],
                                     os.path.join(out_dir, "sheet_%s.png" % g), tile=tile))
    return made, sheets


def make_sheet(pngs, out_png, tile=620, cols=2):
    """One contact sheet, so a whole batch of crops can be checked in a single look."""
    tiles = []
    for p in pngs:
        im = Image.open(p).convert("RGB")
        im = im.resize((tile, max(1, int(im.height * tile / im.width))), Image.LANCZOS)
        tiles.append((os.path.basename(p)[:-4], im))
    rows = (len(tiles) + cols - 1) // cols
    ch = max(t[1].height for t in tiles) + 26
    sheet = Image.new("RGB", (cols * (tile + 20) + 20, rows * ch + 20), "white")
    d = ImageDraw.Draw(sheet)
    for i, (name, im) in enumerate(tiles):
        cx, cy = 20 + (i % cols) * (tile + 20), 20 + (i // cols) * ch
        d.text((cx, cy), name, fill="black")
        sheet.paste(im, (cx, cy + 18))
    sheet.save(out_png)
    return out_png


# ------------------------------------------------------------------ pages
def render_pages(pdf, out_dir, dpi=110, pages=None):
    doc = pymupdf.open(pdf)
    stem = _safe(os.path.splitext(os.path.basename(pdf))[0])
    os.makedirs(out_dir, exist_ok=True)
    out = []
    for i in (pages if pages is not None else range(len(doc))):
        if i >= len(doc):
            continue
        p = os.path.join(out_dir, "%s_p%02d.png" % (stem, i + 1))
        doc[i].get_pixmap(dpi=dpi).save(p)
        out.append(p)
    return out


def _safe(name):
    return "".join(c if (c.isalnum() or c in "-_") else "_" for c in name)

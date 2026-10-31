# -*- coding: utf-8 -*-
"""Turn a built .docx into PDF + one contact sheet, using the local Word install."""
import io, os, sys, tempfile
import pymupdf


def docx_to_pdf(docx_path, pdf_path=None):
    import win32com.client as win32
    docx_path = os.path.abspath(docx_path)
    pdf_path = os.path.abspath(pdf_path or os.path.splitext(docx_path)[0] + ".pdf")
    if os.path.exists(pdf_path):
        os.remove(pdf_path)
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(docx_path, ReadOnly=True, AddToRecentFiles=False)
        doc.SaveAs(pdf_path, FileFormat=17)   # wdFormatPDF
        doc.Close(False)
    finally:
        word.Quit()
    return pdf_path


def contact_sheet(pdf_path, out_png, pages=None, dpi=100, cols=3, tile_w=560):
    from PIL import Image, ImageDraw
    doc = pymupdf.open(pdf_path)
    idx = pages if pages else list(range(len(doc)))
    tiles = []
    for i in idx:
        if i >= len(doc):
            continue
        pm = doc[i].get_pixmap(dpi=dpi)
        im = Image.frombytes("RGB", (pm.width, pm.height), pm.samples)
        im = im.resize((tile_w, int(im.height * tile_w / im.width)), Image.LANCZOS)
        tiles.append((i + 1, im))
    rows = (len(tiles) + cols - 1) // cols
    ch = max(t[1].height for t in tiles) + 24
    sheet = Image.new("RGB", (cols * (tile_w + 12) + 12, rows * ch + 12), "white")
    d = ImageDraw.Draw(sheet)
    for k, (pno, im) in enumerate(tiles):
        cx, cy = 12 + (k % cols) * (tile_w + 12), 12 + (k // cols) * ch
        d.text((cx, cy), "page %d" % pno, fill="black")
        sheet.paste(im, (cx, cy + 16))
    sheet.save(out_png)
    return out_png, len(doc)

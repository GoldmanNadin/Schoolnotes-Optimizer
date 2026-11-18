#!/usr/bin/env python3
"""Baut aus einer Markdown-Mitschrift ein gestaltetes Word-Dokument (+ optional PDF).

  build_skript.py mitschrift.md --farbe 7B4FA3 [--out ordner] [--pdf]

Kopf der .md (Frontmatter):
  ---
  fach: Mathematik
  klasse: 3. Klasse
  titel: Funktionen            # optional, sonst Fachname
  untertitel: Lernunterlage    # optional
  autor: Max Muster
  deckblatt: true              # Standard true
  inhaltsverzeichnis: true     # Standard true
  kapitelseiten: false         # true = '# ' bekommt eigene Titelseite
  ---
Unterstuetzte Markdown-Teile: #/##/### , Absaetze, - Listen, 1. Listen, **fett**, *kursiv*,
==markiert==, Tabellen, ![Bildtext](pfad), > [!merke] / [!achtung] / [!fazit] Boxen, --- Seitenumbruch.
"""
import argparse, re, subprocess, sys, pathlib, shutil, tempfile
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT = "Arial"
CREAM = "F2F0DE"
TEXT_W = 17.0  # cm (A4 21 - 2*2)
ORANGE = "C8741A"
ORANGE_BG = "FDF1E3"


def hx(c): return RGBColor.from_string(c.upper())
def tint(c, f):
    c = c.lstrip("#"); r, g, b = (int(c[i:i + 2], 16) for i in (0, 2, 4))
    return "%02X%02X%02X" % tuple(round(v + (255 - v) * f) for v in (r, g, b))


# ---------- Low-Level-Helfer ----------
def shade(el_pr, fill):
    s = OxmlElement("w:shd"); s.set(qn("w:val"), "clear"); s.set(qn("w:color"), "auto"); s.set(qn("w:fill"), fill)
    el_pr.append(s)

def cell_shade(cell, fill): shade(cell._tc.get_or_add_tcPr(), fill)

def cell_borders(cell, color="BFBFBF", sz=4, sides=("top", "left", "bottom", "right"), left_sz=None, left_color=None):
    pr = cell._tc.get_or_add_tcPr(); b = OxmlElement("w:tcBorders")
    for s in ("top", "left", "bottom", "right"):
        e = OxmlElement(f"w:{s}")
        if s in sides:
            e.set(qn("w:val"), "single"); e.set(qn("w:sz"), str(left_sz if (s == "left" and left_sz) else sz))
            e.set(qn("w:color"), left_color if (s == "left" and left_color) else color)
        else:
            e.set(qn("w:val"), "nil")
        b.append(e)
    pr.append(b)

def cell_margins(cell, top=80, bottom=80, left=140, right=140):
    pr = cell._tc.get_or_add_tcPr(); m = OxmlElement("w:tcMar")
    for k, v in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        e = OxmlElement(f"w:{k}"); e.set(qn("w:w"), str(v)); e.set(qn("w:type"), "dxa"); m.append(e)
    pr.append(m)

def set_col_widths(table, widths_cm):
    table.autofit = False
    tblPr = table._tbl.tblPr
    lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblPr.append(lay)
    for old in tblPr.findall(qn("w:tblW")): tblPr.remove(old)
    tw = OxmlElement("w:tblW"); tw.set(qn("w:w"), str(int(sum(widths_cm) * 567))); tw.set(qn("w:type"), "dxa"); tblPr.append(tw)
    ti = OxmlElement("w:tblInd"); ti.set(qn("w:w"), "0"); ti.set(qn("w:type"), "dxa"); tblPr.append(ti)
    for row in table.rows:
        for c, w in zip(row.cells, widths_cm): c.width = Cm(w)

def para_border_bottom(p, color, sz=8, space=2):
    pPr = p._p.get_or_add_pPr(); b = OxmlElement("w:pBdr"); e = OxmlElement("w:bottom")
    e.set(qn("w:val"), "single"); e.set(qn("w:sz"), str(sz)); e.set(qn("w:space"), str(space)); e.set(qn("w:color"), color)
    b.append(e); pPr.append(b)

def para_border_top(p, color, sz=6, space=4):
    pPr = p._p.get_or_add_pPr(); b = OxmlElement("w:pBdr"); e = OxmlElement("w:top")
    e.set(qn("w:val"), "single"); e.set(qn("w:sz"), str(sz)); e.set(qn("w:space"), str(space)); e.set(qn("w:color"), color)
    b.append(e); pPr.append(b)

def fmt(p, before=0, after=4, line=1.15, keep_next=False):
    pf = p.paragraph_format; pf.space_before = Pt(before); pf.space_after = Pt(after); pf.line_spacing = line
    if keep_next: pf.keep_with_next = True

def run_font(r, size=None, bold=None, italic=None, color=None):
    r.font.name = FONT
    r._r.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), FONT)
    if size: r.font.size = Pt(size)
    if bold is not None: r.font.bold = bold
    if italic is not None: r.font.italic = italic
    if color: r.font.color.rgb = hx(color)

def add_field(run, instr):
    for t, txt in (("begin", None), (None, instr), ("separate", None), ("text", "1"), ("end", None)):
        if t == "text":
            tt = OxmlElement("w:t"); tt.text = txt; run._r.append(tt)
        elif t is None:
            it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = txt; run._r.append(it)
        else:
            fc = OxmlElement("w:fldChar"); fc.set(qn("w:fldCharType"), t); run._r.append(fc)


# ---------- Inline-Markdown ----------
INLINE = re.compile(r"(\*\*.+?\*\*|==.+?==|\*[^*\s][^*]*?\*)")

def add_inline(p, text, size=11, color=None, base_bold=False, hl=None):
    for part in INLINE.split(text):
        if not part: continue
        if part.startswith("**") and part.endswith("**"):
            r = p.add_run(part[2:-2]); run_font(r, size, True, None, color)
        elif part.startswith("==") and part.endswith("=="):
            r = p.add_run(part[2:-2]); run_font(r, size, True, None, color)
            if hl: shade(r._r.get_or_add_rPr(), hl)
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            r = p.add_run(part[1:-1]); run_font(r, size, base_bold, True, color)
        else:
            r = p.add_run(part); run_font(r, size, base_bold, None, color)


# ---------- Dokument ----------
class Builder:
    def __init__(self, meta, farbe, basedir):
        self.m, self.c, self.base = meta, farbe.upper().lstrip("#"), basedir
        self.light, self.mid = tint(self.c, 0.90), tint(self.c, 0.70)
        self.doc = Document()
        self.headings = []  # (level, text)
        self.setup()

    def setup(self):
        d = self.doc; s = d.sections[0]
        s.page_width, s.page_height = Cm(21), Cm(29.7)
        s.left_margin = s.right_margin = Cm(2); s.top_margin = Cm(2); s.bottom_margin = Cm(2)
        st = d.styles["Normal"]; st.font.name = FONT; st.font.size = Pt(11)
        st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
        for lvl, size in ((1, 15), (2, 12.5), (3, 11)):
            h = d.styles[f"Heading {lvl}"]; h.font.name = FONT; h.font.size = Pt(size); h.font.bold = True
            h.font.color.rgb = hx(self.c); h.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
            h.element.rPr.rFonts.set(qn("w:asciiTheme"), "") if False else None
            h.paragraph_format.space_before = Pt({1: 16, 2: 12, 3: 8}[lvl]); h.paragraph_format.space_after = Pt(4)
            h.paragraph_format.keep_with_next = True

    # --- Footer ---
    def footer(self, section):
        section.footer.is_linked_to_previous = False
        p = section.footer.paragraphs[0]; p.text = ""; p.style = self.doc.styles["Normal"]
        para_border_top(p, self.c)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(TEXT_W), WD_TAB_ALIGNMENT.RIGHT)
        r = p.add_run(self.m.get("autor", "")); run_font(r, 8, True, None, self.c)
        r = p.add_run(f"  ·  {self.m['fach']} {self.m.get('klasse', '')}".rstrip()); run_font(r, 8, False, None, self.c)
        r = p.add_run("\tSeite "); run_font(r, 8, True, None, self.c)
        r = p.add_run(); run_font(r, 8, True, None, self.c); add_field(r, "PAGE")
        r = p.add_run(" / "); run_font(r, 8, True, None, self.c)
        r = p.add_run(); run_font(r, 8, True, None, self.c); add_field(r, "NUMPAGES")

    # --- Deckblatt: Fachfarbe, heller Bogen-Ersatz, Creme-Box mit Fach in Grossbuchstaben ---
    def cover(self):
        d = self.doc; s = d.sections[0]
        s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Cm(0)
        s.footer_distance = Cm(0)
        t = d.add_table(rows=1, cols=1); t.alignment = WD_TABLE_ALIGNMENT.LEFT
        set_col_widths(t, [21.0])
        row = t.rows[0]; row.height = Cm(29.2); row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        c = row.cells[0]; cell_shade(c, self.c); c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        cell_borders(c, sides=())
        c.paragraphs[0].text = ""
        inner = c.add_table(rows=1, cols=1); inner.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_col_widths(inner, [12.5])
        ic = inner.rows[0].cells[0]; cell_shade(ic, CREAM); cell_borders(ic, sides=()); cell_margins(ic, 400, 400, 300, 300)
        p = ic.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER; fmt(p, 6, 14)
        r = p.add_run(self.m["fach"].upper()); run_font(r, 24, True, None, "000000")
        p2 = ic.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER; fmt(p2, 0, 4)
        if self.m.get("titel"):
            r = p2.add_run(self.m["titel"]); run_font(r, 13, False, None, "333333")
            p2 = ic.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER; fmt(p2, 0, 4)
        r = p2.add_run(self.m.get("klasse", "")); run_font(r, 16, False, None, "000000")
        # kleiner Autor-Hinweis unten in heller Farbe
        p3 = c.add_paragraph(); p3.alignment = WD_ALIGN_PARAGRAPH.LEFT; fmt(p3, 160, 0)
        p3.paragraph_format.left_indent = Cm(1.8)
        r = p3.add_run(self.m.get("autor", "")); run_font(r, 13, True, None, tint(self.c, 0.75))
        # nach dem Deckblatt: neue Section mit normalen Raendern
        from docx.enum.section import WD_SECTION
        ns = d.add_section(WD_SECTION.NEW_PAGE)
        ns.left_margin = ns.right_margin = ns.top_margin = ns.bottom_margin = Cm(2); ns.footer_distance = Cm(1.0)
        self.footer(ns)
        # Raender der ersten Section-Fussnote leer lassen
        s.footer.is_linked_to_previous = False

    # --- Inhaltsverzeichnis (statisch, Seitenzahlen aus 2. Durchlauf) ---
    def toc_placeholder(self):
        self.toc_pos = len(self.doc.paragraphs)
        p = self.doc.add_paragraph(); r = p.add_run("Inhaltsverzeichnis"); run_font(r, 17, False, None, self.c)
        para_border_bottom(p, self.c, 8); fmt(p, 0, 10)
        self.toc_anchor = p
        self.toc_paras = []

    def build_toc(self, pages):
        """Fuellt das Inhaltsverzeichnis; pages: Liste Seitenzahl je Eintrag (oder '')."""
        # vorhandene Zeilen entfernen
        for tp in self.toc_paras: tp._p.getparent().remove(tp._p)
        self.toc_paras = []
        anchor = self.toc_anchor._p
        for i, (lvl, txt) in enumerate(self.headings):
            p = self.doc.add_paragraph()
            anchor.addnext(p._p) if i == 0 else self.toc_paras[-1]._p.addnext(p._p)
            fmt(p, 6 if lvl == 1 else 0, 1, 1.1)
            p.paragraph_format.left_indent = Cm({1: 0, 2: 0.6, 3: 1.2}[lvl])
            p.paragraph_format.tab_stops.add_tab_stop(Cm(TEXT_W), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
            size, col, bold = (10, self.c, True) if lvl == 1 else ((9, "000000", False) if lvl == 2 else (8.5, "333333", False))
            r = p.add_run(txt); run_font(r, size, bold, None, col)
            r = p.add_run("\t" + str(pages[i] if i < len(pages) else "")); run_font(r, size, bold, None, col)
            self.toc_paras.append(p)
        if self.toc_paras:
            self.toc_paras[-1].add_run().add_break(WD_BREAK.PAGE)

    # --- Bausteine ---
    def heading(self, lvl, text):
        self.headings.append((lvl, text))
        if lvl == 1 and self.m.get("kapitelseiten"):
            self.doc.add_page_break()
            for _ in range(9): self.doc.add_paragraph()
            p = self.doc.add_heading(level=1); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(text); run_font(r, 28, False, None, self.c)
            self.page_break_after = True
            return
        if lvl == 1 and getattr(self, "first_h1_done", False) and not self.m.get("kapitelseiten"):
            self.doc.add_page_break()
        if lvl == 1: self.first_h1_done = True
        p = self.doc.add_heading(level=lvl)
        r = p.add_run(text); run_font(r, {1: 15, 2: 12.5, 3: 11}[lvl], True, None, self.c)
        if lvl == 1: para_border_bottom(p, self.c, 6, 2)

    def para(self, text, size=11):
        p = self.doc.add_paragraph(); fmt(p, 0, 6, 1.2)
        add_inline(p, text, size, hl=self.mid); return p

    def bullet(self, text, level=0, num=None):
        p = self.doc.add_paragraph(); fmt(p, 0, 3, 1.15)
        ind = 0.65 + 0.6 * level
        p.paragraph_format.left_indent = Cm(ind); p.paragraph_format.first_line_indent = Cm(-0.5)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(ind))
        r = p.add_run((f"{num}" if num else "•") + "\t"); run_font(r, 11, bool(num), None, self.c if not num else "000000")
        add_inline(p, text, 11, hl=self.mid)

    def box(self, kind, lines):
        spec = {"merke": ("Merke", self.light, self.c, self.c), "achtung": ("Achtung", ORANGE_BG, ORANGE, ORANGE),
                "fazit": ("Fazit", tint(self.c, 0.93), self.c, self.c),
                "beispiel": ("Beispiel", "F5F5F5", "7F7F7F", "555555")}[kind]
        label, bg, border, lab = spec
        t = self.doc.add_table(rows=1, cols=1); set_col_widths(t, [TEXT_W])
        c = t.rows[0].cells[0]; cell_shade(c, bg); cell_borders(c, border, 6, left_sz=24); cell_margins(c, 100, 100, 200, 160)
        p = c.paragraphs[0]; fmt(p, 0, 3)
        r = p.add_run(label); run_font(r, 10.5, True, None, lab)
        for ln in lines:
            if not ln.strip(): continue
            q = c.add_paragraph(); fmt(q, 0, 2, 1.15)
            if ln.lstrip().startswith("- "):
                q.paragraph_format.left_indent = Cm(0.5); q.paragraph_format.first_line_indent = Cm(-0.4)
                rr = q.add_run("•  "); run_font(rr, 10, False, None, lab); add_inline(q, ln.lstrip()[2:], 10, hl=self.mid)
            else:
                add_inline(q, ln, 10, hl=self.mid)
        sp = self.doc.add_paragraph(); fmt(sp, 0, 4)

    def table(self, rows):
        hdr = rows[0]; n = len(hdr)
        t = self.doc.add_table(rows=len(rows), cols=n); t.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_col_widths(t, [TEXT_W / n] * n)
        for i, row in enumerate(rows):
            tr = t.rows[i]
            trPr = tr._tr.get_or_add_trPr(); cs = OxmlElement("w:cantSplit"); trPr.append(cs)
            if i == 0:
                th = OxmlElement("w:tblHeader"); trPr.append(th)
            for j in range(n):
                c = tr.cells[j]; txt = row[j] if j < len(row) else ""
                cell_borders(c, "BFBFBF", 4); cell_margins(c, 60, 60, 100, 100)
                p = c.paragraphs[0]; fmt(p, 0, 0, 1.1)
                if i == 0:
                    cell_shade(c, self.c); add_inline(p, txt, 10, "FFFFFF", base_bold=True)
                else:
                    if j == 0: cell_shade(c, self.light)
                    add_inline(p, txt, 10, base_bold=(j == 0), hl=self.mid)
        sp = self.doc.add_paragraph(); fmt(sp, 0, 4)

    def image(self, alt, path):
        pth = (self.base / path) if not pathlib.Path(path).is_absolute() else pathlib.Path(path)
        if not pth.exists():
            self.para(f"*[Bild fehlt: {path}]*"); return
        p = self.doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; fmt(p, 4, 2)
        p.add_run().add_picture(str(pth), width=Cm(11))
        if alt:
            q = self.doc.add_paragraph(); q.alignment = WD_ALIGN_PARAGRAPH.CENTER; fmt(q, 0, 8)
            r = q.add_run(alt); run_font(r, 8.5, False, True, "666666")


# ---------- Parser ----------
def parse_front(text):
    meta = {}
    m = re.match(r"^---\s*\n(.*?)\n---\s*\n", text, re.S)
    if m:
        for ln in m.group(1).splitlines():
            if ":" in ln:
                k, v = ln.split(":", 1); v = v.split("#")[0].strip().strip('"')
                meta[k.strip()] = {"true": True, "false": False}.get(v.lower(), v)
        text = text[m.end():]
    return meta, text

def render(b, body, with_toc):
    lines = body.splitlines(); i = 0; first_content = True
    while i < len(lines):
        ln = lines[i]
        if not ln.strip(): i += 1; continue
        if ln.strip() == "---": b.doc.add_page_break(); i += 1; continue
        m = re.match(r"^(#{1,3})\s+(.*)$", ln)
        if m: b.heading(len(m.group(1)), m.group(2).strip()); i += 1; continue
        if ln.startswith(">"):
            blk = []
            while i < len(lines) and lines[i].startswith(">"):
                blk.append(lines[i][1:].lstrip()); i += 1
            k = re.match(r"\[!(\w+)\]\s*(.*)", blk[0]) if blk else None
            if k and k.group(1).lower() in ("merke", "achtung", "fazit", "beispiel"):
                rest = ([k.group(2)] if k.group(2) else []) + blk[1:]
                b.box(k.group(1).lower(), rest)
            else:
                b.box("merke", blk)
            continue
        if ln.lstrip().startswith("|"):
            rows = []
            while i < len(lines) and lines[i].lstrip().startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                if not all(re.fullmatch(r":?-{2,}:?", c) for c in cells): rows.append(cells)
                i += 1
            b.table(rows); continue
        m = re.match(r"^!\[(.*?)\]\((.*?)\)", ln)
        if m: b.image(m.group(1), m.group(2)); i += 1; continue
        m = re.match(r"^(\s*)[-*]\s+(.*)$", ln)
        if m: b.bullet(m.group(2), len(m.group(1)) // 2); i += 1; continue
        m = re.match(r"^(\s*)(\d+)[.)]\s+(.*)$", ln)
        if m: b.bullet(m.group(3), len(m.group(1)) // 2, num=m.group(2) + "."); i += 1; continue
        # Absatz (aufeinanderfolgende Zeilen zusammenfassen)
        buf = [ln.strip()]; i += 1
        while i < len(lines) and lines[i].strip() and not re.match(r"^(#|>|\||!\[|\s*[-*]\s|\s*\d+[.)]\s|---)", lines[i]):
            buf.append(lines[i].strip()); i += 1
        b.para(" ".join(buf))


def build(md_path, farbe, out_dir, pdf):
    md_path = pathlib.Path(md_path); text = md_path.read_text(encoding="utf8")
    meta, body = parse_front(text)
    meta.setdefault("fach", md_path.stem); meta.setdefault("autor", "")
    out_dir = pathlib.Path(out_dir); out_dir.mkdir(parents=True, exist_ok=True)
    name = re.sub(r"[^\w\-]+", "_", f"{meta['fach']}_{meta.get('titel') or 'Mitschrift'}")
    docx_path = out_dir / f"{name}.docx"
    want_toc = meta.get("inhaltsverzeichnis", True) is not False
    want_cover = meta.get("deckblatt", True) is not False

    def make(pages):
        b = Builder(meta, farbe, md_path.parent)
        if want_cover: b.cover()
        else: b.footer(b.doc.sections[0])
        if want_toc: b.toc_placeholder()
        render(b, body, want_toc)
        if want_toc: b.build_toc(pages)
        b.doc.core_properties.author = meta["autor"]; b.doc.core_properties.title = meta.get("titel") or meta["fach"]
        b.doc.save(docx_path); return b

    heads = make([]).headings
    pages = []
    if want_toc and heads:
        make([0] * len(heads))                       # Platzhalterzahlen, damit Layout stimmt
        pdf_path = to_pdf(docx_path)
        pages = find_pages(pdf_path, heads, skip=(2 if want_cover else 1))
        make(pages)
    if pdf: to_pdf(docx_path)
    elif (out_dir / (docx_path.stem + ".pdf")).exists() and not pdf:
        (out_dir / (docx_path.stem + ".pdf")).unlink()
    return docx_path


def to_pdf(docx_path):
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(docx_path.parent), str(docx_path)],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
    return docx_path.with_suffix(".pdf")

def find_pages(pdf_path, heads, skip):
    n = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", str(pdf_path)], capture_output=True, text=True).stdout).group(1))
    texts = [subprocess.run(["pdftotext", "-f", str(p), "-l", str(p), "-layout", str(pdf_path), "-"], capture_output=True, text=True).stdout
             for p in range(1, n + 1)]
    pages, start = [], 1
    toc_pages = 0
    for p, t in enumerate(texts, 1):
        if "Inhaltsverzeichnis" in t or re.search(r"\.{8,}", t): toc_pages = max(toc_pages, p)
    start = toc_pages + 1 if toc_pages else 1
    # TOC kann mehrere Seiten haben: erste Seite nach der letzten TOC-Zeile suchen
    cur = start
    for lvl, txt in heads:
        key = re.sub(r"\s+", " ", txt)
        found = None
        for p in range(cur, n + 1):
            if key in re.sub(r"\s+", " ", texts[p - 1]): found = p; break
        if found is None: pages.append(""); continue
        pages.append(found); cur = found
    return pages


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("md"); ap.add_argument("--farbe", required=True)
    ap.add_argument("--out", default="."); ap.add_argument("--pdf", action="store_true")
    a = ap.parse_args()
    print(build(a.md, a.farbe, a.out, a.pdf))

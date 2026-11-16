#!/usr/bin/env python3
"""Liest ein bestehendes Word-Dokument (.docx) als Text + Bilder aus.
Wird nur einmal pro Fach gebraucht, um eine alte Mitschrift zu übernehmen.

  python docx_lesen.py <datei.docx> <zielordner>
      -> <zielordner>/inhalt.md und <zielordner>/bilder/...

Nur Python-Standardbibliothek.
"""
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
V = "{urn:schemas-microsoft-com:vml}"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def stilnamen(z):
    namen = {}
    try:
        s = ET.fromstring(z.read("word/styles.xml"))
        for st in s.iter(W + "style"):
            n = st.find(W + "name")
            namen[st.get(W + "styleId")] = (n.get(W + "val") if n is not None else "").lower()
    except KeyError:
        pass
    return namen


def lauf_text(r):
    t = "".join((x.text or "") for x in r.iter() if x.tag in (W + "t",))
    if r.find(f".//{W}tab") is not None:
        t += " "
    rpr = r.find(W + "rPr")
    if t.strip() and rpr is not None:
        if rpr.find(W + "b") is not None and rpr.find(W + "b").get(W + "val", "1") not in ("0", "false"):
            t = f"**{t.strip()}** " if t.endswith(" ") else f"**{t.strip()}**"
        if rpr.find(W + "highlight") is not None:
            t = f"=={t}=="
    return t


class Leser:
    def __init__(self, z, ziel):
        self.z, self.ziel = z, Path(ziel)
        self.stile = stilnamen(z)
        self.rels = {}
        try:
            for rel in ET.fromstring(z.read("word/_rels/document.xml.rels")).iter(REL + "Relationship"):
                self.rels[rel.get("Id")] = rel.get("Target")
        except KeyError:
            pass
        self.bilder = 0
        self.zeilen = []

    def bild(self, rid):
        ziel = self.rels.get(rid)
        if not ziel:
            return None
        pfad = "word/" + ziel.lstrip("/").replace("word/", "", 1) if not ziel.startswith("word/") else ziel
        try:
            daten = self.z.read(pfad)
        except KeyError:
            return None
        self.bilder += 1
        endung = Path(ziel).suffix.lower() or ".png"
        name = f"bilder/alt-{self.bilder:02d}{endung}"
        (self.ziel / "bilder").mkdir(parents=True, exist_ok=True)
        (self.ziel / name).write_bytes(daten)
        hinweis = " (EMF/WMF – Word-Grafik, evtl. nicht anzeigbar)" if endung in (".emf", ".wmf") else ""
        return f"[BILD {self.bilder}: {name}{hinweis}]"

    def absatz(self, p):
        stil = ""
        ppr = p.find(W + "pPr")
        if ppr is not None and ppr.find(W + "pStyle") is not None:
            sid = ppr.find(W + "pStyle").get(W + "val")
            stil = self.stile.get(sid, sid.lower())
        teile, bilder = [], []
        for el in p.iter():
            if el.tag == W + "r":
                teile.append(lauf_text(el))
            elif el.tag == A + "blip":
                b = self.bild(el.get(R + "embed"))
                if b:
                    bilder.append(b)
            elif el.tag == V + "imagedata":
                b = self.bild(el.get(R + "id"))
                if b:
                    bilder.append(b)
        text = re.sub(r"\s+", " ", "".join(teile)).strip()
        text = text.replace("****", "")
        prefix = ""
        if re.search(r"heading 1|überschrift 1|^title|titel", stil):
            prefix = "## "
        elif re.search(r"heading [2-9]|überschrift [2-9]", stil):
            prefix = "### "
        elif ppr is not None and ppr.find(W + "numPr") is not None:
            ebene = ppr.find(f"{W}numPr/{W}ilvl")
            tiefe = int(ebene.get(W + "val")) if ebene is not None else 0
            prefix = "   " * tiefe + "- "
        elif re.search(r"toc|verzeichnis", stil):
            return []  # Inhaltsverzeichnis überspringen
        aus = []
        if text:
            aus.append(prefix + text)
        aus += bilder
        return aus

    def tabelle(self, tbl):
        zeilen = []
        for tr in tbl.findall(W + "tr"):
            zellen = []
            for tc in tr.findall(W + "tc"):
                inhalt = []
                for p in tc.iter(W + "p"):
                    inhalt += self.absatz(p)
                zellen.append(" / ".join(x.lstrip("#- ") for x in inhalt if x).replace("|", "\\|"))
            zeilen.append(zellen)
        if not zeilen:
            return []
        if len(zeilen) == 1 and len(zeilen[0]) == 1:  # Kasten mit einer Zelle (z. B. „Merke“)
            return ["[KASTEN]", zeilen[0][0], "[/KASTEN]"]
        breite = max(len(z) for z in zeilen)
        zeilen = [z + [""] * (breite - len(z)) for z in zeilen]
        aus = ["| " + " | ".join(zeilen[0]) + " |", "|" + "---|" * breite]
        aus += ["| " + " | ".join(z) + " |" for z in zeilen[1:]]
        return aus

    def lesen(self):
        body = ET.fromstring(self.z.read("word/document.xml")).find(W + "body")
        for el in body:
            if el.tag == W + "p":
                self.zeilen += self.absatz(el)
            elif el.tag == W + "tbl":
                self.zeilen += [""] + self.tabelle(el) + [""]
            elif el.tag == W + "sdt":  # z. B. Inhaltsverzeichnis
                continue
        return "\n".join(self.zeilen)


def main(datei, ziel):
    Path(ziel).mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(datei) as z:
        l = Leser(z, ziel)
        text = l.lesen()
    text = re.sub(r"\n{3,}", "\n\n", text).strip() + "\n"
    (Path(ziel) / "inhalt.md").write_text(f"# Alte Mitschrift: {Path(datei).name}\n\n" + text, encoding="utf-8")
    print(json.dumps({"ok": True, "inhalt": str(Path(ziel) / "inhalt.md"), "bilder": l.bilder,
                      "zeichen": len(text)}, ensure_ascii=False))


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])

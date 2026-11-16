#!/usr/bin/env python3
"""Liest Exporte aus dem Notizblock (Mitschrift_*.html).

  python export_lesen.py finden [startordner]
      -> listet alle unverarbeiteten Exporte (JSON) aus _Eingang-Ordnern und Downloads

  python export_lesen.py auspacken <export.html> <zielordner>
      -> schreibt <zielordner>/inhalt.md (lesbarer Text mit Markierungen),
         <zielordner>/meta.json und <zielordner>/bilder/bild-01.png ...

Nur Python-Standardbibliothek, keine Installation nötig.
"""
import base64
import json
import os
import re
import sys
from html.parser import HTMLParser
from pathlib import Path

LEGENDE = """<!-- Legende der Markierungen:
## Überschrift 1 / ### Überschrift 2 / #### Überschrift 3   (von der Schülerin gesetzt)
**fett**  *kursiv*  __unterstrichen__  ==mit Textmarker markiert==  [farbig:rot]text[/farbig] (farbiger Text)
[MERKE-KASTEN] ... [/MERKE-KASTEN]   = von der Schülerin als „Merke“ markiert (wichtig!)
[BEISPIEL-KASTEN] ... [/BEISPIEL-KASTEN]  = als Beispiel markiert
[ZITAT] ... [/ZITAT]                   = Zitat / Definition
[HANDSCHRIFT n: datei | ...]           = direkt aufs Blatt gezeichnet/geschrieben (Mindmap, Pfeile, Kreise,
                                         Notizen). Der Ausschnitt zeigt auch den Text darunter -> UNBEDINGT ansehen!
[ZEICHNUNG n: datei | Breite x%]       = mit der Maus gezeichnet -> Bild ansehen!
[BILD n: datei | Breite x%]            = eingefügtes Bild/Screenshot -> Bild ansehen!
[BEMALTES-BILD n: datei | Breite x%]   = Bild, auf das gezeichnet wurde -> Bild ansehen!
-->
"""

FARBNAMEN = {
    (211, 47, 47): "rot", (21, 101, 192): "blau", (46, 125, 50): "grün",
    (31, 35, 40): None, (0, 0, 0): None,
}


def rgb(wert):
    wert = wert.strip().lower()
    m = re.match(r"rgba?\((\d+),\s*(\d+),\s*(\d+)", wert)
    if m:
        return tuple(int(x) for x in m.groups())
    m = re.match(r"#([0-9a-f]{6})", wert)
    if m:
        n = int(m.group(1), 16)
        return (n >> 16, (n >> 8) & 255, n & 255)
    return None


def farbname(c):
    if c is None:
        return None
    if c in FARBNAMEN:
        return FARBNAMEN[c]
    r, g, b = c
    if max(c) - min(c) < 40 and max(c) < 90:
        return None  # schwarz/grau = normaler Text
    if r > g and r > b:
        return "rot"
    if b > r and b > g:
        return "blau"
    if g > r and g > b:
        return "grün"
    return "farbig"


class Leser(HTMLParser):
    BLOCK = {"p", "div", "h1", "h2", "h3", "h4", "li", "blockquote", "tr", "table", "ul", "ol", "main", "body"}

    def __init__(self, bildordner, version=1):
        super().__init__(convert_charrefs=True)
        self.version = version
        self.kaesten = []
        self.listenklassen = []
        self.quelle = False
        self.zeilen = []
        self.puffer = ""
        self.listen = []          # Stapel: "ul"/"ol" mit Zähler
        self.zaehler = []
        self.inline = []          # Stapel der Endmarker für Inline-Tags
        self.tabelle = None       # Liste von Zeilen (Liste von Zellen)
        self.zelle = None
        self.in_main = False
        self.bildordner = Path(bildordner)
        self.bilder = []
        self.ignorieren = 0

    # --- Hilfen ---
    def schreibe(self, text):
        if self.zelle is not None:
            self.zelle.append(text)
        else:
            self.puffer += text

    def zeile_beenden(self):
        t = self.puffer.rstrip()
        if t.strip() and t.strip() not in ("-", "##", "###") and not re.fullmatch(r"\s*\d+\.", t):
            listenzeile = re.match(r"\s*(- |– |\d+\. |[a-z]\) |[ivxlcdmIVXLCDM]+\. |#)", t)
            self.zeilen.append(t if listenzeile else t.strip())
        self.puffer = ""

    def leerzeile(self):
        self.zeile_beenden()
        if self.zeilen and self.zeilen[-1] != "":
            self.zeilen.append("")

    # --- Parser ---
    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "main" and a.get("id") == "notizen":
            self.in_main = True
            return
        if not self.in_main:
            return
        if tag in ("script", "style"):
            self.ignorieren += 1
            return
        if tag == "svg":
            self.ignorieren += 1
            return
        if tag in ("h1", "h2", "h3", "h4"):
            self.leerzeile()
            if self.version >= 2:
                self.puffer = {"h1": "## ", "h2": "### "}.get(tag, "#### ")
            else:
                self.puffer = "## " if tag in ("h1", "h2") else "### "
        elif tag == "p" and "bildquelle" in (a.get("class") or ""):
            self.zeile_beenden(); self.puffer = "(Bildquelle) "
        elif tag in ("p", "div"):
            if self.zelle is None:
                self.zeile_beenden()
            else:
                if self.zelle:
                    self.zelle.append(" / ")
        elif tag == "br":
            if self.zelle is not None:
                self.zelle.append(" / ")
            else:
                self.zeile_beenden()
        elif tag in ("ul", "ol"):
            self.zeile_beenden()
            self.listen.append(tag); self.zaehler.append(0); self.listenklassen.append(a.get("class") or "")
        elif tag == "li":
            self.zeile_beenden()
            tiefe = max(0, len(self.listen) - 1)
            if self.listen and self.listen[-1] == "ol":
                self.zaehler[-1] += 1
                self.puffer = "   " * tiefe + nummer(self.zaehler[-1], self.listen, self.listenklassen) + " "
            else:
                strich = any("striche" in k for k in self.listenklassen)
                self.puffer = "   " * tiefe + ("– " if strich else "- ")
        elif tag == "blockquote":
            art = {"beispiel": "BEISPIEL-KASTEN", "zitat": "ZITAT"}.get((a.get("class") or "").strip(), "MERKE-KASTEN")
            self.kaesten.append(art)
            self.leerzeile(); self.zeilen.append(f"[{art}]")
        elif tag == "table":
            self.leerzeile(); self.tabelle = []
        elif tag == "tr" and self.tabelle is not None:
            self.tabelle.append([])
        elif tag in ("td", "th") and self.tabelle is not None:
            self.zelle = []
        elif tag in ("b", "strong"):
            self.schreibe("**"); self.inline.append("**")
        elif tag in ("i", "em"):
            self.schreibe("*"); self.inline.append("*")
        elif tag == "u":
            self.schreibe("__"); self.inline.append("__")
        elif tag == "mark":
            self.schreibe("=="); self.inline.append("==")
        elif tag in ("span", "font"):
            stil = a.get("style", "") or ""
            ende = ""
            bg = re.search(r"background(?:-color)?:\s*([^;]+)", stil)
            if bg and rgb(bg.group(1)) not in (None, (255, 255, 255)):
                self.schreibe("=="); ende = "==" + ende
            fg = re.search(r"(?<!-)color:\s*([^;]+)", stil)
            name = farbname(rgb(fg.group(1))) if fg else farbname(rgb(a.get("color", ""))) if a.get("color") else None
            if name:
                self.schreibe(f"[farbig:{name}]"); ende = "[/farbig]" + ende
            if "font-weight: bold" in stil or "font-weight: 700" in stil:
                self.schreibe("**"); ende = "**" + ende
            if "font-style: italic" in stil:
                self.schreibe("*"); ende = "*" + ende
            if "underline" in stil:
                self.schreibe("__"); ende = "__" + ende
            self.inline.append(ende)
        elif tag == "img":
            self.bild(a)

    def handle_endtag(self, tag):
        if tag == "main":
            self.in_main = False
            self.zeile_beenden()
            return
        if not self.in_main:
            return
        if tag in ("script", "style", "svg"):
            self.ignorieren = max(0, self.ignorieren - 1)
            return
        if tag in ("b", "strong", "i", "em", "u", "mark", "span", "font"):
            if self.inline:
                self.schreibe(self.inline.pop())
        elif tag in ("h1", "h2", "h3", "h4"):
            self.leerzeile()
        elif tag in ("p", "div"):
            if self.zelle is None:
                self.zeile_beenden()
        elif tag in ("ul", "ol"):
            self.zeile_beenden()
            if self.listen:
                self.listen.pop(); self.zaehler.pop(); self.listenklassen.pop()
            if not self.listen:
                self.leerzeile()
        elif tag == "li":
            self.zeile_beenden()
        elif tag == "blockquote":
            art = self.kaesten.pop() if self.kaesten else "MERKE-KASTEN"
            self.zeile_beenden(); self.zeilen.append(f"[/{art}]"); self.zeilen.append("")
        elif tag in ("td", "th") and self.zelle is not None and self.tabelle is not None:
            text = re.sub(r"\s+", " ", "".join(self.zelle)).strip(" /")
            if self.tabelle:
                self.tabelle[-1].append(text.replace("|", "\\|"))
            self.zelle = None
        elif tag == "table" and self.tabelle is not None:
            zeilen = [z for z in self.tabelle if z]
            if zeilen:
                breite = max(len(z) for z in zeilen)
                zeilen = [z + [""] * (breite - len(z)) for z in zeilen]
                self.zeilen.append("| " + " | ".join(zeilen[0]) + " |")
                self.zeilen.append("|" + "---|" * breite)
                for z in zeilen[1:]:
                    self.zeilen.append("| " + " | ".join(z) + " |")
            self.tabelle = None
            self.zeilen.append("")

    def handle_data(self, data):
        if not self.in_main or self.ignorieren:
            return
        text = re.sub(r"\s+", " ", data)
        if not text.strip() and not self.puffer and self.zelle is None:
            return
        self.schreibe(text)

    def bild(self, a):
        src = a.get("src", "")
        m = re.match(r"data:image/([a-zA-Z0-9.+-]+);base64,(.*)", src, re.S)
        if not m:
            return
        endung = {"jpeg": "jpg", "svg+xml": "svg"}.get(m.group(1).lower(), m.group(1).lower())
        nr = len(self.bilder) + 1
        name = f"bild-{nr:02d}.{endung}"
        self.bildordner.mkdir(parents=True, exist_ok=True)
        (self.bildordner / name).write_bytes(base64.b64decode(m.group(2)))
        art = a.get("data-art", "bild")
        marke = {"zeichnung": "ZEICHNUNG", "bemaltes-bild": "BEMALTES-BILD", "handschrift": "HANDSCHRIFT"}.get(art, "BILD")
        breite = a.get("data-breite", "")
        info = f"[{marke} {nr}: bilder/{name}" + (f" | Breite {breite}" if breite else "")
        if a.get("data-bezug"):
            info += f" | berührt den Text: „{a['data-bezug']}“"
        if a.get("data-texte"):
            info += f" | aufs Blatt geschriebene Wörter: {a['data-texte']}"
        if a.get("data-quelle"):
            info += f" | Quelle: {a['data-quelle']}"
        info += "]"
        self.bilder.append({"nr": nr, "datei": f"bilder/{name}", "art": art, "breite": breite})
        if self.zelle is not None:
            self.zelle.append(" " + info + " ")
        else:
            self.zeile_beenden(); self.zeilen.append(info)


def roemisch(n):
    w = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]
    r = ""
    for wert, z in w:
        while n >= wert:
            r += z; n -= wert
    return r


def nummer(n, listen, klassen):
    """Nummer wie im Notizblock: 1. / a) / i. – je nach Art und Ebene."""
    ol_tiefe = sum(1 for t in listen if t == "ol") - 1
    klasse = next((k for t, k in zip(listen, klassen) if t == "ol" and k), "")
    if "roemisch" in klasse and ol_tiefe == 0:
        return roemisch(n) + "."
    if "buchstaben" in klasse and ol_tiefe == 0:
        return chr(96 + min(n, 26)) + ")"
    return [f"{n}.", chr(96 + min(n, 26)) + ")", roemisch(n).lower() + "."][min(ol_tiefe, 2)]


def meta_lesen(html):
    meta = {}
    for name, wert in re.findall(r'<meta name="mitschrift-([a-z]+)" content="([^"]*)"', html):
        import html as h
        meta[name] = h.unescape(wert)
    return meta


def auspacken(datei, ziel):
    datei, ziel = Path(datei), Path(ziel)
    ziel.mkdir(parents=True, exist_ok=True)
    html = datei.read_text(encoding="utf-8")
    meta = meta_lesen(html)
    if "<main" not in html:  # notfalls: ganzer body
        html = re.sub(r"<body[^>]*>", '<body><main id="notizen">', html).replace("</body>", "</main></body>")
    try:
        version = int(meta.get("version", "1"))
    except ValueError:
        version = 1
    leser = Leser(ziel / "bilder", version)
    leser.feed(html)
    leser.zeile_beenden()
    text = "\n".join(leser.zeilen)
    text = re.sub(r"\n{3,}", "\n\n", text).strip() + "\n"
    kopf = f"# {meta.get('fach', '?')} – {meta.get('thema') or '(ohne Thema)'} – {meta.get('datum', '?')}\n"
    (ziel / "inhalt.md").write_text(LEGENDE + kopf + "\n" + text, encoding="utf-8")
    meta["quelle"] = datei.name
    meta["bilder"] = leser.bilder
    (ziel / "meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"ok": True, "inhalt": str(ziel / "inhalt.md"), "fach": meta.get("fach"),
                      "thema": meta.get("thema"), "datum": meta.get("datum"),
                      "bilder": [str(ziel / b["datei"]) for b in leser.bilder]}, ensure_ascii=False, indent=2))


def eingangsordner(start):
    start = Path(start).resolve()
    gefunden = []
    kandidaten = [start] + list(start.parents)[:3]
    for k in kandidaten:
        for ordner in [k / "_Eingang"] + [p / "_Eingang" for p in sorted(k.iterdir()) if p.is_dir()] if k.is_dir() else []:
            try:
                if ordner.is_dir() and ordner not in gefunden:
                    gefunden.append(ordner)
            except OSError:
                pass
    return gefunden


def finden(start="."):
    exporte = []
    for ordner in eingangsordner(start):
        for f in sorted(ordner.glob("*.htm*")):
            exporte.append(f)
    dl = Path.home() / "Downloads"
    if dl.is_dir():
        exporte += sorted(dl.glob("Mitschrift_*.htm*"))
    liste = []
    for f in exporte:
        try:
            meta = meta_lesen(f.read_text(encoding="utf-8", errors="ignore")[:4000])
        except OSError:
            continue
        if not meta.get("fach"):
            continue
        liste.append({"datei": str(f), "fach": meta.get("fach"), "thema": meta.get("thema"),
                      "datum": meta.get("datum"), "erstellt": meta.get("erstellt"),
                      "groesse_kb": round(f.stat().st_size / 1024)})
    print(json.dumps({"eingang": [str(o) for o in eingangsordner(start)], "exporte": liste},
                     ensure_ascii=False, indent=2))


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) >= 2 and sys.argv[1] == "finden":
        finden(sys.argv[2] if len(sys.argv) > 2 else ".")
    elif len(sys.argv) == 4 and sys.argv[1] == "auspacken":
        auspacken(sys.argv[2], sys.argv[3])
    else:
        print(__doc__)
        sys.exit(1)

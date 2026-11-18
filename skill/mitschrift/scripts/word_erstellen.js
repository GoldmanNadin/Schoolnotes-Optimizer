#!/usr/bin/env node
/*
 * Baut das komplette Word-Dokument eines Fachs aus mitschrift.json.
 *
 *   node word_erstellen.js <Fach>/_Mitschrift-Daten/mitschrift.json [--ausgabe <datei.docx>]
 *
 * - Design (Farben pro Fach) kommt aus ../faecher.json
 * - Vor dem Überschreiben wird die alte Word-Datei nach _Mitschrift-Daten/backup/ kopiert (die letzten 10 bleiben).
 * - Exit-Code 3 = Word-Datei ist gerade geöffnet (bitte schließen und nochmal versuchen).
 */
"use strict";
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
  BorderStyle, AlignmentType, HeadingLevel, ImageRun, Footer, PageNumber, PositionalTab,
  PositionalTabAlignment, PositionalTabRelativeTo, PositionalTabLeader, LevelFormat, Bookmark,
  InternalHyperlink, LineRuleType, TableLayoutType, Tab,
  Math: OMath, MathRun, MathFraction, MathSuperScript, MathSubScript, MathSubSuperScript, MathRadical,
} = require("docx");
const JSZip = require("jszip");

/* ---------------- Eingabe ---------------- */
const args = process.argv.slice(2);
if (!args[0]) { console.error("Aufruf: node word_erstellen.js <mitschrift.json> [--ausgabe datei.docx]"); process.exit(1); }
const jsonPfad = path.resolve(args[0]);
const D = JSON.parse(fs.readFileSync(jsonPfad, "utf8"));
// EINZEL = eine Stunde pro Word-Datei (Standard): { fach, datum, thema, datei, bloecke }
// sonst alte Form mit Titelseite: { fach, datei, stunden: [...] }
const EINZEL = Array.isArray(D.bloecke);
// Bilder liegen relativ zum Ordner „_Mitschrift-Daten“ (die JSON kann auch in …/stunden/ liegen)
let datenOrdner = path.dirname(jsonPfad);
for (let d = datenOrdner; path.dirname(d) !== d; d = path.dirname(d)) { if (path.basename(d) === "_Mitschrift-Daten") { datenOrdner = d; break; } }
const fachOrdner = path.dirname(datenOrdner);
const dateiSauber = s => String(s).replace(/[\\/:*?"<>|]+/g, "").replace(/\s+/g, " ").trim().slice(0, 80);
const ausIdx = args.indexOf("--ausgabe");
const ausgabe = ausIdx >= 0 ? path.resolve(args[ausIdx + 1])
  : path.join(fachOrdner, D.datei || (EINZEL ? `${dateiSauber(`${D.datum || ""} ${D.thema || "Mitschrift"}`)}.docx` : `Mitschrift ${D.fach}.docx`));

const faecher = JSON.parse(fs.readFileSync(path.join(__dirname, "..", "faecher.json"), "utf8"));
const fachKey = Object.keys(faecher).find(k => k.toLowerCase() === String(D.fach || "").toLowerCase());
const fachInfo = Object.assign({}, faecher._standard, fachKey ? faecher[fachKey] : {});
const FACHNAME = D.name || fachInfo.name || D.fach || "Mitschrift";
const sprache = D.sprache || fachInfo.sprache || "de";

/* ---------------- Farben ---------------- */
function hex2rgb(h) { const n = parseInt(String(h).replace("#", ""), 16); return [n >> 16, (n >> 8) & 255, n & 255]; }
function rgb2hex(c) { return c.map(x => Math.max(0, Math.min(255, Math.round(x))).toString(16).padStart(2, "0")).join("").toUpperCase(); }
function mische(h, ziel, anteil) { const a = hex2rgb(h), b = hex2rgb(ziel); return rgb2hex(a.map((x, i) => x + (b[i] - x) * anteil)); }
const BASIS = (D.farbe || fachInfo.farbe || "#455A64");
const F = {
  akzent: rgb2hex(hex2rgb(BASIS)),
  dunkel: mische(BASIS, "#000000", 0.25),
  hell: mische(BASIS, "#FFFFFF", 0.90),
  mittel: mische(BASIS, "#FFFFFF", 0.75),
  text: "262626", grau: "6B7280", linie: "D0D5DB",
};

/* ---------------- Maße ---------------- */
const SEITE_B = 11906, RAND = 1134;               // A4, 2 cm Rand
const BREITE = SEITE_B - 2 * RAND;                // 9638 DXA ≈ 17 cm
const FONT = "Arial";

/* ---------------- Text mit **fett**, *kursiv*, ==markiert==, __unterstrichen__ ---------------- */
/* ---------------- Formeln: LaTeX (aus dem Notizblock) → echte Word-Formel ---------------- */
const SYMBOLE = {
  cdot: "·", times: "×", div: "÷", pm: "±", mp: "∓", le: "≤", leq: "≤", ge: "≥", geq: "≥", ne: "≠", neq: "≠",
  approx: "≈", equiv: "≡", infty: "∞", to: "→", rightarrow: "→", leftarrow: "←", Rightarrow: "⇒", Leftrightarrow: "⇔",
  leftrightarrow: "↔", degree: "°", circ: "°", percent: "%", ldots: "…", dots: "…", cdots: "⋯", in: "∈", notin: "∉",
  subset: "⊂", cup: "∪", cap: "∩", emptyset: "∅", forall: "∀", exists: "∃", partial: "∂", nabla: "∇", sum: "∑",
  prod: "∏", int: "∫", angle: "∠", perp: "⊥", parallel: "∥", lt: "<", gt: ">", mid: "|", vert: "|",
  alpha: "α", beta: "β", gamma: "γ", delta: "δ", epsilon: "ε", varepsilon: "ε", zeta: "ζ", eta: "η", theta: "θ",
  vartheta: "ϑ", iota: "ι", kappa: "κ", lambda: "λ", mu: "μ", nu: "ν", xi: "ξ", pi: "π", rho: "ρ", sigma: "σ",
  tau: "τ", upsilon: "υ", phi: "φ", varphi: "φ", chi: "χ", psi: "ψ", omega: "ω", Gamma: "Γ", Delta: "Δ",
  Theta: "Θ", Lambda: "Λ", Xi: "Ξ", Pi: "Π", Sigma: "Σ", Phi: "Φ", Psi: "Ψ", Omega: "Ω",
  exponentialE: "e", imaginaryI: "i", differentialD: "d",
};
const FUNKTIONEN = new Set(["sin", "cos", "tan", "cot", "arcsin", "arccos", "arctan", "sinh", "cosh", "tanh", "log", "ln", "lg", "exp", "lim", "max", "min", "det", "mod"]);
function latexZuMath(src) {
  let i = 0;
  const flach = a => a.flat(Infinity);
  const mindestens = a => { const f = flach(a); return f.length ? f : [new MathRun(" ")]; };
  const leer = () => { while (i < src.length && /\s/.test(src[i])) i++; };
  function rohGruppe() { // Text in {…} unverändert
    leer(); if (src[i] !== "{") return "";
    let tiefe = 0, start = ++i;
    for (; i < src.length; i++) { if (src[i] === "{") tiefe++; else if (src[i] === "}") { if (!tiefe) break; tiefe--; } }
    return src.slice(start, i++);
  }
  // Argument ohne {} = genau EIN Zeichen/Befehl (LaTeX: \frac34 = 3/4, x^23 = x²3)
  function argument() { leer(); if (src[i] === "{") { i++; return liste("}"); } if (/[0-9]/.test(src[i] || "")) return [new MathRun(src[i++])]; const a = atom(); return a == null ? [] : [a]; }
  function befehl() {
    i++; // Backslash
    let name = "";
    if (/[a-zA-Z]/.test(src[i] || "")) { while (/[a-zA-Z]/.test(src[i] || "")) name += src[i++]; }
    else name = src[i++] || "";
    switch (name) {
      case "frac": case "dfrac": case "tfrac": case "cfrac": { const z = argument(), n = argument(); return new MathFraction({ numerator: mindestens(z), denominator: mindestens(n) }); }
      case "sqrt": { let grad = null; leer(); if (src[i] === "[") { let t = 0, s = ++i; for (; i < src.length; i++) { if (src[i] === "[") t++; else if (src[i] === "]") { if (!t) break; t--; } } grad = latexZuListe(src.slice(s, i++)); } const inh = argument(); return new MathRadical({ children: mindestens(inh), ...(grad && flach(grad).length ? { degree: flach(grad) } : {}) }); }
      case "left": case "right": case "big": case "Big": case "bigl": case "bigr": case "Bigl": case "Bigr": { leer(); if (src[i] === "\\") { i++; let d = ""; while (/[a-zA-Z]/.test(src[i] || "")) d += src[i++]; if (!d) d = src[i++]; return new MathRun({ lbrace: "{", rbrace: "}", vert: "|", langle: "⟨", rangle: "⟩", "{": "{", "}": "}", "|": "‖" }[d] || ""); } const c = src[i++]; return c === "." ? null : new MathRun(c); }
      case "text": case "textrm": case "mathrm": case "operatorname": case "mathit": case "mathbf": case "textbf": case "mbox": return new MathRun(rohGruppe().replace(/\\,|\\;|\\ /g, " "));
      case "placeholder": { leer(); if (src[i] === "[") { while (i < src.length && src[i] !== "]") i++; i++; } rohGruppe(); return null; }
      case ",": case ";": case ":": case " ": case "quad": case "qquad": return new MathRun(" ");
      case "!": return null;
      case "{": case "}": case "%": case "$": case "#": case "&": case "_": return new MathRun(name);
      case "\\": return new MathRun(" ");
      default:
        if (SYMBOLE[name]) return new MathRun(SYMBOLE[name]);
        if (FUNKTIONEN.has(name)) return new MathRun(name);
        return new MathRun(name);
    }
  }
  function atom() {
    leer();
    if (i >= src.length) return null;
    const c = src[i];
    if (c === "\\") return befehl();
    if (c === "{") { i++; return liste("}"); }
    if (/[0-9.,]/.test(c)) { let z = ""; while (/[0-9.,]/.test(src[i] || "")) z += src[i++]; return new MathRun(z); }
    i++;
    return new MathRun({ "-": "−", "*": "·", "'": "′" }[c] || c);
  }
  function liste(ende) {
    const out = [];
    while (i < src.length) {
      leer();
      if (ende && src[i] === ende) { i++; break; }
      if (src[i] === "}") { i++; continue; }
      if (src[i] === "^" || src[i] === "_") {
        let hoch = null, tief = null;
        while (src[i] === "^" || src[i] === "_") { const c = src[i++]; const a = argument(); if (c === "^") hoch = a; else tief = a; leer(); }
        let basis = out.pop(); basis = basis == null ? [new MathRun(" ")] : flach([basis]);
        if (hoch && tief) out.push(new MathSubSuperScript({ children: basis, subScript: mindestens(tief), superScript: mindestens(hoch) }));
        else if (hoch) out.push(new MathSuperScript({ children: basis, superScript: mindestens(hoch) }));
        else out.push(new MathSubScript({ children: basis, subScript: mindestens(tief) }));
        continue;
      }
      const a = atom();
      if (a != null) out.push(a);
    }
    return out;
  }
  function latexZuListe(s) { const alt = [src, i]; src = s; i = 0; const r = liste(null); [src, i] = alt; return r; }
  return new OMath({ children: mindestens(liste(null)) });
}

// Textstück; \t wird zu einem echten Word-Tabulator
function lauf(text, basis) {
  if (!text.includes("\t")) return new TextRun({ text, ...basis });
  const kinder = [];
  text.split("\t").forEach((teil, i) => { if (i) kinder.push(new Tab()); if (teil) kinder.push(teil); });
  return new TextRun({ ...basis, children: kinder });
}
function laeufe(text, basis = {}) {
  const out = [];
  const s = String(text ?? "");
  const re = /(\[\[[^\]]+?\]\]|\$[^$]+?\$|\^\{[^}]*\}|_\{[^}]*\}|\*\*[^*]+?\*\*|==[^=]+?==|__[^_]+?__|\*[^*\s][^*]*?\*)/g;
  let pos = 0, m;
  while ((m = re.exec(s))) {
    if (m.index > pos) out.push(lauf(s.slice(pos, m.index), basis));
    const t = m[0];
    // verschachtelt erlaubt, z. B. **fett mit ==Markierung==**
    if (t.startsWith("[[")) out.push(...laeufe(t.slice(2, -2), { ...basis, border: { style: BorderStyle.SINGLE, size: 8, color: "262626", space: 1 } }));
    else if (t.startsWith("$")) out.push(latexZuMath(t.slice(1, -1)));
    else if (t.startsWith("^{")) out.push(...laeufe(t.slice(2, -1), { ...basis, superScript: true }));
    else if (t.startsWith("_{")) out.push(...laeufe(t.slice(2, -1), { ...basis, subScript: true }));
    else if (t.startsWith("**")) out.push(...laeufe(t.slice(2, -2), { ...basis, bold: true }));
    else if (t.startsWith("==")) out.push(...laeufe(t.slice(2, -2), { ...basis, bold: true, color: basis.color || F.dunkel, shading: { type: ShadingType.CLEAR, color: "auto", fill: F.mittel } }));
    else if (t.startsWith("__")) out.push(...laeufe(t.slice(2, -2), { ...basis, underline: {} }));
    else out.push(...laeufe(t.slice(1, -1), { ...basis, italics: true }));
    pos = m.index + t.length;
  }
  if (pos < s.length) out.push(lauf(s.slice(pos), basis));
  return out;
}
const absatzText = (text, opt = {}) => String(text ?? "").split("\n").map(z => new Paragraph({ children: laeufe(z, opt.run), ...(opt.para || {}) }));

/* ---------------- Listen ---------------- */
let nummernInstanz = 0;
// art: Aufzählung "punkte" (•) oder "striche" (–); Nummerierung "nummern" (1.), "buchstaben" (a)), "roemisch" (I.)
function liste(punkte, nummeriert, ebene = 0, instanz = null, art = null) {
  if (nummeriert && instanz === null) instanz = ++nummernInstanz;
  const ref = nummeriert ? ({ buchstaben: "buchstaben", roemisch: "roemisch" }[art] || "nummern") : (art === "striche" ? "striche" : "punkte");
  const out = [];
  for (const p of punkte || []) {
    const text = typeof p === "string" ? p : (p.text ?? "");
    out.push(new Paragraph({
      children: laeufe(text),
      numbering: nummeriert ? { reference: ref, level: Math.min(ebene, 1), instance: instanz } : { reference: ref, level: Math.min(ebene, 2) },
      spacing: { after: 60 },
    }));
    const unter = typeof p === "object" && (p.unterpunkte || p.punkte);
    if (unter && unter.length) out.push(...liste(unter, nummeriert, ebene + 1, instanz, art));
  }
  return out;
}

/* ---------------- Kästen ---------------- */
const KEIN = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
function abstand(pt = 4) { return new Paragraph({ spacing: { before: 0, after: 0, line: pt * 20, lineRule: LineRuleType.EXACT }, children: [] }); }
function kasten({ titel, titelFarbe, fill, rand, rundum = false, kinder }) {
  const linie = { style: BorderStyle.SINGLE, size: rundum ? 8 : 24, color: rand };
  const zelle = new TableCell({
    width: { size: BREITE, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill },
    borders: rundum ? { top: linie, bottom: linie, left: linie, right: linie } : { top: KEIN, bottom: KEIN, right: KEIN, left: linie },
    margins: { top: 110, bottom: 110, left: 220, right: 220 },
    children: [
      ...(titel ? [new Paragraph({ spacing: { after: 60 }, keepNext: true, children: [new TextRun({ text: titel.toUpperCase(), bold: true, size: 17, color: titelFarbe, characterSpacing: 30 })] })] : []),
      ...kinder,
    ],
  });
  return [abstand(4), new Table({ width: { size: BREITE, type: WidthType.DXA }, columnWidths: [BREITE], layout: TableLayoutType.FIXED, rows: [new TableRow({ cantSplit: false, children: [zelle] })] }), abstand(8)];
}
function kastenInhalt(b) {
  const k = [];
  if (b.text) k.push(...absatzText(b.text, { para: { spacing: { after: 60 } } }));
  if (b.punkte && b.punkte.length) k.push(...liste(b.punkte, false));
  if (b.nummern && b.nummern.length) k.push(...liste(b.nummern, true));
  if (!k.length) k.push(new Paragraph({ children: [] }));
  return k;
}

/* ---------------- Tabellen ---------------- */
function zellAbsaetze(inhalt, run) {
  return String(inhalt ?? "").split("\n").map(z => {
    const punkt = z.match(/^\s*[-•]\s+(.*)$/);
    return punkt
      ? new Paragraph({ children: laeufe(punkt[1], run), numbering: { reference: "punkte", level: 0 }, spacing: { after: 40 } })
      : new Paragraph({ children: laeufe(z, run), spacing: { after: 40 } });
  });
}
function spaltenBreiten(anzahl, prozente) {
  let p = (prozente && prozente.length === anzahl) ? prozente.map(Number) : Array(anzahl).fill(100 / anzahl);
  const summe = p.reduce((a, b) => a + b, 0);
  const w = p.map(x => Math.floor(BREITE * x / summe));
  w[w.length - 1] += BREITE - w.reduce((a, b) => a + b, 0);
  return w;
}
function tabelle(b) {
  const kopf = b.kopf || null;
  const zeilen = b.zeilen || [];
  const anzahl = Math.max(kopf ? kopf.length : 0, ...zeilen.map(z => z.length), 1);
  const breiten = spaltenBreiten(anzahl, b.breiten);
  const rahmen = { style: BorderStyle.SINGLE, size: 4, color: F.linie };
  const rand = { top: rahmen, bottom: rahmen, left: rahmen, right: rahmen };
  const zelle = (inhalt, i, fill, run) => new TableCell({
    width: { size: breiten[i], type: WidthType.DXA }, borders: rand,
    shading: { type: ShadingType.CLEAR, color: "auto", fill }, margins: { top: 70, bottom: 70, left: 120, right: 120 },
    children: zellAbsaetze(inhalt, run),
  });
  const rows = [];
  if (kopf) rows.push(new TableRow({ tableHeader: true, cantSplit: true, children: Array.from({ length: anzahl }, (_, i) => zelle(kopf[i] ?? "", i, F.akzent, { bold: true, color: "FFFFFF" })) }));
  zeilen.forEach((z, r) => rows.push(new TableRow({ cantSplit: true, children: Array.from({ length: anzahl }, (_, i) => zelle(z[i] ?? "", i, r % 2 ? F.hell : "FFFFFF", i === 0 && b.erste_spalte_fett !== false && anzahl > 1 ? { bold: true } : {})) })));
  const out = [abstand(4), new Table({ width: { size: BREITE, type: WidthType.DXA }, columnWidths: breiten, layout: TableLayoutType.FIXED, rows })];
  if (b.unterschrift) out.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60 }, children: [new TextRun({ text: b.unterschrift, italics: true, size: 18, color: F.grau })] }));
  out.push(abstand(8));
  return out;
}
function begriffe(b) {
  const eintraege = b.eintraege || [];
  const breiten = spaltenBreiten(2, [26, 74]);
  const unten = { style: BorderStyle.SINGLE, size: 4, color: F.mittel };
  const rand = { top: unten, bottom: unten, left: KEIN, right: KEIN };
  const out = [];
  if (b.titel !== null) out.push(new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(b.titel || (sprache === "en" ? "Key terms" : "Wichtige Begriffe"))] }));
  out.push(new Table({
    width: { size: BREITE, type: WidthType.DXA }, columnWidths: breiten, layout: TableLayoutType.FIXED,
    rows: eintraege.map(e => new TableRow({ cantSplit: true, children: [
      new TableCell({ width: { size: breiten[0], type: WidthType.DXA }, borders: rand, shading: { type: ShadingType.CLEAR, color: "auto", fill: F.hell }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: laeufe(e.begriff, { bold: true, color: F.dunkel }) })] }),
      new TableCell({ width: { size: breiten[1], type: WidthType.DXA }, borders: rand, margins: { top: 70, bottom: 70, left: 160, right: 120 }, children: zellAbsaetze(e.erklaerung) }),
    ] })),
  }), abstand(8));
  return out;
}

/* ---------------- Bilder ---------------- */
function bildMasse(buf) {
  if (buf[0] === 0x89 && buf[1] === 0x50) return { typ: "png", w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) };
  if (buf[0] === 0xff && buf[1] === 0xd8) {
    let i = 2;
    while (i < buf.length) {
      if (buf[i] !== 0xff) { i++; continue; }
      const m = buf[i + 1], len = buf.readUInt16BE(i + 2);
      if (m >= 0xc0 && m <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(m)) return { typ: "jpg", h: buf.readUInt16BE(i + 5), w: buf.readUInt16BE(i + 7) };
      i += 2 + len;
    }
  }
  if (buf.slice(0, 3).toString() === "GIF") return { typ: "gif", w: buf.readUInt16LE(6), h: buf.readUInt16LE(8) };
  if (buf[0] === 0x42 && buf[1] === 0x4d) return { typ: "bmp", w: buf.readInt32LE(18), h: Math.abs(buf.readInt32LE(22)) };
  return null;
}
function bild(b) {
  const datei = path.resolve(datenOrdner, b.datei);
  if (!fs.existsSync(datei)) { console.warn(`WARNUNG: Bild fehlt: ${b.datei}`); return [new Paragraph({ children: [new TextRun({ text: `[Bild fehlt: ${b.datei}]`, color: "C00000" })] })]; }
  const buf = fs.readFileSync(datei);
  const m = bildMasse(buf);
  if (!m) { console.warn(`WARNUNG: Bildformat nicht unterstützt: ${b.datei}`); return []; }
  const pxProCm = 96 / 2.54;
  let wCm = Math.min(Number(b.breite_cm) || 12, 17);
  let hCm = wCm * m.h / m.w;
  if (hCm > 20) { wCm = wCm * 20 / hCm; hCm = 20; }
  const out = [new Paragraph({
    alignment: AlignmentType.CENTER, keepNext: !!b.unterschrift, spacing: { before: 120, after: b.unterschrift ? 40 : 160 },
    children: [new ImageRun({ type: m.typ, data: buf, transformation: { width: Math.round(wCm * pxProCm), height: Math.round(hCm * pxProCm) }, altText: { title: b.unterschrift || "Bild", description: b.unterschrift || "Bild aus der Mitschrift", name: path.basename(datei) } })],
  })];
  if (b.unterschrift) out.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: b.unterschrift, italics: true, size: 18, color: F.grau })] }));
  return out;
}

/* ---------------- Bausteine ---------------- */
const T = sprache === "en"
  ? { merke: "Remember", beispiel: "Example", achtung: "Important", zusammenfassung: "In a nutshell", uebersicht: "Overview", stunde: "Lesson", datum: "Date", thema: "Topic" }
  : { merke: "Merke", beispiel: "Beispiel", achtung: "Achtung", zusammenfassung: "Kurz zusammengefasst", uebersicht: "Übersicht", stunde: "Nr.", datum: "Datum", thema: "Thema" };

const AUSRICHTUNG = { mitte: AlignmentType.CENTER, zentriert: AlignmentType.CENTER, rechts: AlignmentType.RIGHT, blocksatz: AlignmentType.JUSTIFIED, links: AlignmentType.LEFT };
function absatzOpt(b) {
  const o = {};
  if (b.ausrichtung && AUSRICHTUNG[b.ausrichtung]) o.alignment = AUSRICHTUNG[b.ausrichtung];
  if (b.einzug_cm) o.indent = { left: Math.round(Number(b.einzug_cm) * 567) };
  // Tabulatoren (\t im Text) springen wie in Word auf Tabstopps alle 1,25 cm
  return o;
}
function baustein(b) {
  switch (b.typ) {
    case "abschnitt": return [new Paragraph({ heading: HeadingLevel.HEADING_2, ...absatzOpt(b), children: laeufe(b.text) })];
    case "unterabschnitt": return [new Paragraph({ heading: HeadingLevel.HEADING_3, ...absatzOpt(b), children: laeufe(b.text) })];
    case "text": return absatzText(b.text, { para: absatzOpt(b) });
    case "liste": return [...liste(b.punkte, false, 0, null, b.art), abstand(4)];
    case "nummern": return [...liste(b.punkte, true, 0, null, b.art), abstand(4)];
    case "tabelle": return tabelle(b);
    case "begriffe": return begriffe(b);
    case "bild": return bild(b);
    case "merke": return kasten({ titel: b.titel || T.merke, titelFarbe: F.akzent, fill: F.hell, rand: F.akzent, kinder: kastenInhalt(b) });
    case "beispiel": return kasten({ titel: b.titel || T.beispiel, titelFarbe: "5B6470", fill: "F4F5F7", rand: "9AA3AE", kinder: kastenInhalt(b) });
    case "achtung": return kasten({ titel: b.titel || T.achtung, titelFarbe: "B42318", fill: "FDECEA", rand: "D92D20", kinder: kastenInhalt(b) });
    case "rahmen": return kasten({ titel: b.titel || null, titelFarbe: F.dunkel, fill: "FFFFFF", rand: "262626", rundum: true, kinder: kastenInhalt(b) });
    case "zusammenfassung": return kasten({ titel: b.titel || T.zusammenfassung, titelFarbe: F.dunkel, fill: "FFFFFF", rand: F.akzent, rundum: true, kinder: kastenInhalt(b) });
    case "formel": {
      if (b.latex) { // echte Word-Formel, mittig
        const out = [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 160 }, children: [latexZuMath(b.latex)] })];
        if (b.erklaerung) out.push(...absatzText(b.erklaerung, { para: { alignment: AlignmentType.CENTER, spacing: { after: 160 } }, run: { size: 19, color: F.grau } }));
        return out;
      }
      const out = [new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { before: 120, after: b.erklaerung ? 40 : 160 }, keepNext: !!b.erklaerung,
        shading: { type: ShadingType.CLEAR, color: "auto", fill: F.hell },
        border: { top: { style: BorderStyle.SINGLE, size: 6, color: F.hell, space: 4 }, bottom: { style: BorderStyle.SINGLE, size: 6, color: F.hell, space: 4 } },
        children: laeufe(b.text, { font: "Cambria Math", size: 28, color: F.dunkel }),
      })];
      if (b.erklaerung) out.push(...absatzText(b.erklaerung, { para: { alignment: AlignmentType.CENTER, spacing: { after: 160 } }, run: { size: 19, color: F.grau } }));
      return out;
    }
    default:
      console.warn(`WARNUNG: unbekannter Baustein-Typ "${b.typ}" – als Text ausgegeben`);
      return absatzText(b.text || JSON.stringify(b));
  }
}

/* ---------------- Datum ---------------- */
function datumLang(iso) {
  const d = new Date(iso + "T12:00:00");
  if (isNaN(d)) return iso || "";
  const tage = sprache === "en" ? ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"] : ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"];
  const monate = sprache === "en" ? ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"] : ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"];
  return sprache === "en" ? `${tage[d.getDay()]}, ${d.getDate()} ${monate[d.getMonth()]} ${d.getFullYear()}` : `${tage[d.getDay()]}, ${d.getDate()}. ${monate[d.getMonth()]} ${d.getFullYear()}`;
}
const datumKurz = iso => { const m = String(iso || "").match(/^(\d{4})-(\d{2})-(\d{2})$/); return m ? `${m[3]}.${m[2]}.${m[1]}` : (iso || "–"); };

/* ---------------- Dokument ---------------- */
const stunden = EINZEL ? [D] : (D.stunden || []).slice().sort((a, b) => String(a.datum || "").localeCompare(String(b.datum || "")));
const marke = i => `stunde_${i + 1}`;
const kinder = [];

if (EINZEL) {
  // Nur ihre Stunde: kleines graues Datum oben, dann genau ihr Inhalt – keine Titelseite, kein Kapitel
  if (D.datum) kinder.push(new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: datumLang(D.datum), size: 18, color: F.grau })] }));
  for (const b of D.bloecke) kinder.push(...baustein(b));
} else {
// Titelseite mit Übersicht
kinder.push(new Paragraph({ spacing: { before: 1400, after: 0 }, children: [new TextRun({ text: FACHNAME, bold: true, size: 60, color: F.akzent })] }));
kinder.push(new Paragraph({
  spacing: { after: 480 }, border: { bottom: { style: BorderStyle.SINGLE, size: 24, color: F.akzent, space: 8 } },
  children: [new TextRun({ text: D.untertitel || (sprache === "en" ? "Class notes" : "Mitschrift"), size: 28, color: F.grau })],
}));
if (stunden.length) {
  kinder.push(new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: T.uebersicht.toUpperCase(), bold: true, size: 20, color: F.akzent, characterSpacing: 40 })] }));
  const breiten = spaltenBreiten(3, [9, 18, 73]);
  const rahmen = { style: BorderStyle.SINGLE, size: 4, color: F.linie };
  const rand = { top: KEIN, left: KEIN, right: KEIN, bottom: rahmen };
  const z = (inhalt, i, opt = {}) => new TableCell({ width: { size: breiten[i], type: WidthType.DXA }, borders: rand, shading: opt.fill ? { type: ShadingType.CLEAR, color: "auto", fill: opt.fill } : undefined, margins: { top: 60, bottom: 60, left: 120, right: 120 }, children: [new Paragraph({ children: Array.isArray(inhalt) ? inhalt : [new TextRun({ text: inhalt, ...(opt.run || {}) })] })] });
  const rows = [new TableRow({ tableHeader: true, children: [z(T.stunde, 0, { fill: F.akzent, run: { bold: true, color: "FFFFFF" } }), z(T.datum, 1, { fill: F.akzent, run: { bold: true, color: "FFFFFF" } }), z(T.thema, 2, { fill: F.akzent, run: { bold: true, color: "FFFFFF" } })] })];
  stunden.forEach((s, i) => rows.push(new TableRow({ cantSplit: true, children: [
    z(String(i + 1), 0, { run: { color: F.grau } }), z(datumKurz(s.datum), 1, { run: { color: F.grau } }),
    z([new InternalHyperlink({ anchor: marke(i), children: [new TextRun({ text: s.thema || "Mitschrift", color: F.dunkel })] })], 2),
  ] })));
  kinder.push(new Table({ width: { size: BREITE, type: WidthType.DXA }, columnWidths: breiten, layout: TableLayoutType.FIXED, rows }));
}
const heuteIso = new Date().toISOString().slice(0, 10);
kinder.push(new Paragraph({ spacing: { before: 240 }, children: [new TextRun({ text: (sprache === "en" ? "Last updated: " : "Stand: ") + datumKurz(heuteIso), size: 16, color: F.grau })] }));

// Stunden
stunden.forEach((s, i) => {
  kinder.push(new Paragraph({ pageBreakBefore: true, spacing: { after: 60 }, keepNext: true, children: [new TextRun({ text: datumLang(s.datum).toUpperCase(), bold: true, size: 17, color: F.akzent, characterSpacing: 30 })] }));
  kinder.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new Bookmark({ id: marke(i), children: [new TextRun(s.thema || "Mitschrift")] })] }));
  if (s.einleitung) kinder.push(...absatzText(s.einleitung, { run: { italics: true, color: "4B5563" }, para: { spacing: { after: 160 } } }));
  for (const b of s.bloecke || []) kinder.push(...baustein(b));
});
}

const doc = new Document({
  creator: "Notizblock → Claude",
  title: EINZEL ? `${FACHNAME} – ${D.thema || "Mitschrift"} (${datumKurz(D.datum)})` : `${FACHNAME} – Mitschrift`,
  description: "Automatisch gestaltete Unterrichtsmitschrift",
  styles: {
    default: { document: { run: { font: FONT, size: 22, color: F.text }, paragraph: { spacing: { after: 120, line: 288 } } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 36, bold: true, color: "FFFFFF" },
        paragraph: { spacing: { before: 0, after: 280 }, keepNext: true, outlineLevel: 0,
          shading: { type: ShadingType.CLEAR, color: "auto", fill: F.akzent },
          border: ["top", "bottom", "left", "right"].reduce((o, k) => (o[k] = { style: BorderStyle.SINGLE, size: 6, color: F.akzent, space: k === "left" || k === "right" ? 8 : 6 }, o), {}),
          indent: { left: 170, right: 170 } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 28, bold: true, color: F.dunkel },
        paragraph: { spacing: { before: 360, after: 140 }, keepNext: true, outlineLevel: 1, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: F.akzent, space: 3 } } } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 24, bold: true, color: F.akzent },
        paragraph: { spacing: { before: 240, after: 80 }, keepNext: true, outlineLevel: 2 } },
    ],
  },
  numbering: {
    config: [
      { reference: "punkte", levels: [
        { level: 0, format: LevelFormat.BULLET, text: "●", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 280 } }, run: { color: F.akzent, size: 14 } } },
        { level: 1, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 800, hanging: 280 } }, run: { color: F.akzent } } },
        { level: 2, format: LevelFormat.BULLET, text: "·", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1200, hanging: 280 } }, run: { color: F.akzent } } },
      ] },
      { reference: "nummern", levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 320 } }, run: { color: F.akzent, bold: true } } },
        { level: 1, format: LevelFormat.LOWER_LETTER, text: "%2)", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 800, hanging: 320 } }, run: { color: F.akzent } } },
      ] },
      { reference: "striche", levels: [
        { level: 0, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 280 } }, run: { color: F.akzent } } },
        { level: 1, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 800, hanging: 280 } }, run: { color: F.akzent } } },
        { level: 2, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1200, hanging: 280 } }, run: { color: F.akzent } } },
      ] },
      { reference: "buchstaben", levels: [
        { level: 0, format: LevelFormat.LOWER_LETTER, text: "%1)", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 320 } }, run: { color: F.akzent, bold: true } } },
        { level: 1, format: LevelFormat.LOWER_ROMAN, text: "%2.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 800, hanging: 320 } }, run: { color: F.akzent } } },
      ] },
      { reference: "roemisch", levels: [
        { level: 0, format: LevelFormat.UPPER_ROMAN, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 480, hanging: 400 } }, run: { color: F.akzent, bold: true } } },
        { level: 1, format: LevelFormat.LOWER_LETTER, text: "%2)", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 880, hanging: 320 } }, run: { color: F.akzent } } },
      ] },
    ],
  },
  sections: [{
    properties: { page: { size: { width: SEITE_B, height: 16838 }, margin: { top: RAND, bottom: RAND, left: RAND, right: RAND, footer: 567 } } },
    footers: { default: new Footer({ children: [EINZEL ? new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ size: 16, color: F.grau, children: [sprache === "en" ? "Page " : "Seite ", PageNumber.CURRENT, sprache === "en" ? " of " : " von ", PageNumber.TOTAL_PAGES] })],
    }) : new Paragraph({
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: F.mittel, space: 6 } },
      children: [
        new TextRun({ text: `${FACHNAME} · ${sprache === "en" ? "Class notes" : "Mitschrift"}`, size: 16, color: F.grau }),
        new TextRun({ size: 16, color: F.grau, children: [new PositionalTab({ alignment: PositionalTabAlignment.RIGHT, relativeTo: PositionalTabRelativeTo.MARGIN, leader: PositionalTabLeader.NONE }), sprache === "en" ? "Page " : "Seite ", PageNumber.CURRENT, sprache === "en" ? " of " : " von ", PageNumber.TOTAL_PAGES] }),
      ],
    })] }) },
    children: kinder,
  }],
});

/* ---------------- Schreiben (mit Sicherung) ---------------- */
async function ohneOrdnerEintraege(buf) {
  // Ordner-Einträge im ZIP entfernen (OneDrive/Word-Online mögen sie nicht)
  const zip = await JSZip.loadAsync(buf);
  Object.keys(zip.files).filter(n => zip.files[n].dir).forEach(n => zip.remove(n));
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}
function sichern() {
  if (!fs.existsSync(ausgabe)) return null;
  const ordner = path.join(datenOrdner, "backup");
  fs.mkdirSync(ordner, { recursive: true });
  const stempel = new Date().toISOString().replace(/[-:]/g, "").replace("T", "-").slice(0, 15);
  const ziel = path.join(ordner, `${path.basename(ausgabe, ".docx")}_${stempel}.docx`);
  fs.copyFileSync(ausgabe, ziel);
  const alle = fs.readdirSync(ordner).filter(n => n.startsWith(path.basename(ausgabe, ".docx") + "_2") && n.endsWith(".docx")).sort();
  alle.slice(0, Math.max(0, alle.length - 10)).forEach(n => fs.unlinkSync(path.join(ordner, n)));
  return ziel;
}

(async () => {
  const buf = await ohneOrdnerEintraege(await Packer.toBuffer(doc));
  // Testen, ob die Datei beschreibbar ist (Word offen?)
  if (fs.existsSync(ausgabe)) {
    try { const fd = fs.openSync(ausgabe, "r+"); fs.closeSync(fd); }
    catch (e) { console.error(`FEHLER_DATEI_OFFEN: ${ausgabe} ist gerade geöffnet (Word schließen und nochmal versuchen).`); process.exit(3); }
  }
  const sicherung = sichern();
  try { fs.writeFileSync(ausgabe, buf); }
  catch (e) {
    if (["EBUSY", "EPERM", "EACCES"].includes(e.code)) { console.error(`FEHLER_DATEI_OFFEN: ${ausgabe} ist gerade geöffnet (Word schließen und nochmal versuchen).`); process.exit(3); }
    throw e;
  }
  console.log(JSON.stringify({ ok: true, datei: ausgabe, stunden: stunden.length, sicherung, farbe: "#" + F.akzent }, null, 2));
})().catch(e => { console.error("FEHLER:", e.message); process.exit(1); });

---
name: mitschrift
description: Verarbeitet Unterrichtsnotizen aus dem Notizblock (Dateien „Mitschrift_*.html“ im gewählten Speicherort/_Eingang oder in Downloads) und macht aus jeder Stunde ein eigenes Word-Dokument im Fach-Ordner – wortgetreu, schön gestaltet in der Farbe des Fachs, Zeichnungen sauber nachgezeichnet, ohne Titelseite. Verwenden bei „/mitschrift“, „Mitschrift verarbeiten“, „Notizen in Word“, „neue Stunde eintragen“ oder wenn eine Mitschrift_*.html-Datei auftaucht. Auch zum Ändern des Designs der Mitschriften.
---

# Mitschrift → gestaltetes Word-Dokument

Die Schülerin schreibt im Unterricht im **Notizblock** (Desktop-App) und drückt am Ende
„Stunde beenden → Claude“. Dadurch landet eine Datei `Mitschrift_<Fach>_<Datum>_<Thema>.html`
im Ordner `_Eingang` (oder in `Downloads`). Deine Aufgabe: daraus eine übersichtliche,
gestaltete Mitschrift machen – **wortgetreu, ohne etwas dazuzuschreiben** – als **eigenes
Word-Dokument für genau diese Stunde** (keine Titelseite, kein Inhaltsverzeichnis, kein Kapitel-Balken).

**Sprich mit ihr einfach, kurz und auf Deutsch.** Sie hat keinen technischen Hintergrund:
keine Fachbegriffe wie JSON, Skript, Node; Fortschritt in ein bis zwei kurzen Sätzen.

`SKILL` = der Ordner, in dem diese Datei liegt (Skripte unter `SKILL/scripts/`).

## Wie die Daten liegen

```
Schule/2026_27/                        ← Schuljahres-Ordner (enthält _Eingang)
├─ _Eingang/                           ← neue Exporte aus dem Notizblock
│  └─ erledigt/                        ← verarbeitete Exporte (nie löschen)
├─ EBC/
│  ├─ 2026-11-16 Das Internet.docx     ← EIN Word-Dokument PRO STUNDE
│  ├─ 2026-11-18 Netzwerke.docx
│  └─ _Mitschrift-Daten/
│     ├─ stunden/2026-11-16_Das-Internet.json   ← Inhalt dieser Stunde (Quelle der .docx)
│     ├─ bilder/                       ← nachgezeichnete Skizzen, Bilder
│     └─ backup/                       ← automatische Sicherungen, falls eine .docx neu gebaut wird
└─ Mathe/ …
```

Jedes Word-Dokument wird aus seiner Stunden-Datei (`stunden/….json`) erzeugt. Darum: **nie direkt
in der .docx herumbasteln**, sondern die Stunden-Datei ändern und neu bauen.
Ältere Word-Dateien im Fach-Ordner (z. B. eine frühere `Mitschrift.docx`) **nicht anfassen**.

## Ablauf

### 1. Vorbereitung (einmalig, schnell prüfen)
- Falls `SKILL/scripts/node_modules` fehlt: `npm install --prefix "SKILL/scripts"` ausführen.

### 2. Exporte finden
```bash
python "SKILL/scripts/export_lesen.py" finden "<aktueller Ordner>"
```
Liefert alle offenen Exporte (aus `_Eingang`-Ordnern und `Downloads/Mitschrift_*.html`).
- Gesucht wird zuerst in dem Ordner, den sie im Notizblock über den 📁-Knopf gewählt hat
  (steht in `%APPDATA%/Notizblock/einstellungen.json`), dann in `_Eingang`-Ordnern in der Nähe.
- Keine gefunden → freundlich fragen, wo sie gespeichert hat (im Notizblock oben beim 📁-Knopf zu sehen).
- Mehrere → alle der Reihe nach verarbeiten (nach Datum sortiert).
- Der **Schuljahres-Ordner** (mit den Fach-Ordnern) ist der Ordner, in dem `_Eingang` liegt.
  Liegt ein Export in Downloads, gehört er zum Schuljahres-Ordner des gewählten Eingangs.

### 3. Export auspacken und GENAU lesen
```bash
python "SKILL/scripts/export_lesen.py" auspacken "<export.html>" "<Fach>/_Mitschrift-Daten/eingang-tmp"
```
- `inhalt.md` vollständig lesen (Legende der Markierungen steht oben in der Datei).
  Überschriften: `##` = Überschrift 1, `###` = Überschrift 2, `####` = Überschrift 3 –
  das ist ihre eigene Gliederung, daran orientieren.
- **Jedes Bild und jede Zeichnung mit dem Read-Tool ansehen.** Zeichnungen enthalten oft
  wichtigen Inhalt (Skizzen, Diagramme, Pfeile, Rechnungen).
- **`[HANDSCHRIFT n: …]`** = sie hat direkt aufs Blatt gezeichnet oder geschrieben (Mindmap um
  ein Wort, Pfeile zwischen Begriffen, Kreise, Randnotizen). Der Ausschnitt zeigt die Zeichnung
  **zusammen mit dem getippten Text darunter**; „berührt den Text“ sagt, zu welchen Absätzen sie
  gehört, „aufs Blatt geschriebene Wörter“ sind getippte Beschriftungen. Bedeutung erschließen und
  sauber nachzeichnen (siehe Schritt 5b).
- `(Bildquelle) …` bzw. „Quelle:“ bei einem Bild aus dem Internet → als Bildunterschrift übernehmen
  (Lizenzangabe gehört dazu).
- Bilder, die ins Dokument sollen, nach `<Fach>/_Mitschrift-Daten/bilder/` kopieren mit
  sprechendem Namen: `<datum>_<kurzname>.png` (z. B. `2026-11-16_client-server.png`).
- Danach `eingang-tmp` löschen.

### 4. Fach-Ordner
Fach-Ordner = `<Schuljahres-Ordner>/<fach aus dem Export>` (anlegen, falls er fehlt), darin
`_Mitschrift-Daten/stunden/` und `_Mitschrift-Daten/bilder/` (anlegen, falls sie fehlen).

### 5. Neue Stunde gestalten
Aus den Notizen die Stunden-Datei `<Fach>/_Mitschrift-Daten/stunden/<datum>_<Thema-mit-Bindestrichen>.json`
schreiben (Regeln: siehe „Gestaltungsregeln“ – **wortgetreu, nichts dazuschreiben**):
```json
{ "fach": "EBC", "datum": "2026-11-16", "thema": "Das Internet", "farbe": "#1F7A5C",
  "datei": "2026-11-16 Das Internet.docx",
  "quelle": "Mitschrift_EBC_2026-11-16_Das-Internet.html",
  "bloecke": [ … ] }
```
- `datei` = Name des Word-Dokuments im Fach-Ordner: `<datum> <Thema>.docx`.
- `farbe` = die Fachfarbe aus dem Export (`farbe` in der Ausgabe von „auspacken“) – die hat sie beim
  Anlegen des Fachs im Notizblock gewählt. Fehlt sie, gilt `faecher.json`.
- **Gibt es die Stunden-Datei schon** (gleiches Datum + Thema)? Dann wurde die Stunde nochmal
  exportiert (z. B. nach dem Weiterschreiben) → **überschreiben**; das Word-Dokument wird ersetzt
  (die alte Version landet automatisch in `backup/`).
- Kein Thema angegeben → ihre erste Überschrift als Thema nehmen; gibt es keine, „Mitschrift“.

### 5b. Zeichnungen sauber nachzeichnen
Für jede Skizze/Mindmap/jedes Diagramm aus `[HANDSCHRIFT …]` oder `[ZEICHNUNG …]`:
1. Ausschnitt genau ansehen: welche Wörter, Kästen, Pfeile (Richtung!), Verbindungen, Farben.
2. Eine **saubere SVG** schreiben, die **genau dasselbe** zeigt – gleiche Wörter (nur Tippfehler
   korrigieren), gleiche Pfeilrichtungen, gleiche Anordnung (oben/unten/links/rechts), **nichts
   hinzufügen, nichts weglassen**, keine Überschrift oder Erklärung im Bild.
   Stil: `viewBox` ca. 640–900 breit, weißer Hintergrund, Schrift Arial 16–19 px, Linien 2–2,5 px,
   Pfeilspitzen per `<marker>`, Hauptbegriff in einem abgerundeten Kasten mit heller Fachfarbe,
   weitere Begriffe in schlichten Kästen; Farbe = Fachfarbe aus `faecher.json` (hat sie Farben
   bewusst zur Unterscheidung benutzt, die Unterscheidung beibehalten).
   Datei: `<Fach>/_Mitschrift-Daten/bilder/<datum>_<kurzname>.svg`
3. In ein Bild umwandeln:
```bash
node "SKILL/scripts/svg_zu_png.js" "<…>.svg" "<…>.png"
```
4. Das PNG mit dem Read-Tool ansehen und mit dem Original vergleichen (alles drin? Pfeile richtig?
   nichts überlappt?). Dann als `bild` einfügen – **ohne Bildunterschrift**.
- Nur Unterstreichungen, Kreise um Wörter, Häkchen → **kein** Bild, sondern das Wort ==hervorheben==.
- Steht ein getipptes Wort mitten in der Zeichnung (z. B. „Internet“ als Mitte der Mindmap), gehört
  es in die Zeichnung und steht dann nicht zusätzlich als eigener Absatz da.
- Fotos/Screenshots/Bilder aus dem Internet bleiben unverändert.

### 6. Word-Dokument bauen
```bash
node "SKILL/scripts/word_erstellen.js" "<Fach>/_Mitschrift-Daten/stunden/<datum>_<Thema>.json"
```
- Exit-Code 3 / `FEHLER_DATEI_OFFEN` → sagen: *„Bitte schließ kurz die Word-Datei <name>,
  dann mache ich weiter.“* und danach erneut ausführen.
- Gibt es die Word-Datei schon, wird die alte automatisch gesichert (`backup/`, die letzten 10).

### 7. Kontrolle (immer)
```bash
powershell -ExecutionPolicy Bypass -File "SKILL/scripts/vorschau.ps1" "<datei.docx>" "<scratchpad>/vorschau.pdf"
```
Die PDF-Seiten der neuen Stunde ansehen: Ist **jedes ihrer Wörter** drin – und **kein fremdes**?
Tabellen lesbar? Zeichnungen richtig groß? Falls nötig die Stunden-Datei korrigieren und neu bauen. (Falls Word nicht
verfügbar ist, diesen Schritt auslassen.)

### 8. Aufräumen und kurz melden
- Export nach `_Eingang/erledigt/` verschieben (aus Downloads ebenfalls dorthin).
- Kurze Rückmeldung, z. B.:
  *„Fertig ✓ EBC → „2026-11-16 Das Internet.docx“ (1 Mindmap sauber nachgezeichnet,
  2 Tippfehler korrigiert).“*
  Wenn etwas unklar war (unleserliches Wort), das in einem Satz erwähnen.

## Gestaltungsregeln (so sollen die Mitschriften IMMER aussehen)

**Ziel:** Ihre Mitschrift – **genau ihr Text, so wie der Lehrer ihn gegeben hat** – übersichtlich
und schön gestaltet. Claude ist hier Layouter, nicht Autor.

**Text: wortgetreu**
1. **Jedes Wort bleibt ihres.** Nichts umformulieren, nichts zusammenfassen, Stichworte bleiben
   Stichworte, Abkürzungen bleiben Abkürzungen.
2. **Nur Rechtschreib- und Tippfehler** still korrigieren (z. B. „Netzwerkkatre“ → „Netzwerkkarte“,
   Groß-/Kleinschreibung). Sonst nichts verändern.
3. **Nichts dazuschreiben:** keine Einleitung, keine Zusammenfassung, keine Begriffsliste, keine
   Erklärungen, keine Beispiele, keine Bildunterschriften, keine Abschnitts-Nummern, die sie nicht
   selbst geschrieben hat. Nichts weglassen.
4. Unleserliches/Unklares nicht raten: so lassen wie geschrieben und nur in der Rückmeldung an sie
   erwähnen (nicht ins Dokument schreiben).
5. Ausnahme: Quellenangaben von Bildern aus dem Internet (`(Bildquelle) …`) bleiben als kleine
   Bildunterschrift stehen – die gehören zur Lizenz.

**Aufbau: ordnen ist erlaubt (ohne neue Wörter)**
6. Ihre Überschriften → `abschnitt` (Überschrift 1) / `unterabschnitt` (Überschrift 2 und 3),
   genau mit ihrem Wortlaut.
7. Ihre Listen bleiben Listen in ihrer Art (`liste` für • und –, `nummern` für 1. / a) / I.).
8. Was sie mit `⇥` (Tab) in Spalten geschrieben hat („Begriff ⇥ Erklärung“) → `tabelle`
   (Kopfzeile nur, wenn sie selbst eine Kopfzeile geschrieben hat; sonst `kopf` weglassen).
9. Ihre Kästen: `[MERKE-KASTEN]` → `merke`, `[BEISPIEL-KASTEN]` → `beispiel`, `[ZITAT]` → `merke`
   mit Titel „Zitat“. Zusätzlich darf Wichtiges, das sie **markiert, farbig, unterstrichen, eingekreist
   oder mit „!!“** geschrieben hat, in einen `merke`-Kasten oder ==hervorgehoben== werden.
   „Achtung“-Kasten nur, wenn sie selbst „Achtung“/„Vorsicht“/„Test!“ o. Ä. geschrieben hat.
10. **Formeln** aus dem Notizblock stehen als `$…$` (LaTeX) im Text: **genau so übernehmen** – mitten im
    Satz bleibt `$…$` im `text` stehen (wird in Word eine echte Formel). `[Formelzeile] $…$` → Baustein
    `formel` mit `latex` (ohne `$`), ohne `erklaerung`, außer sie hat eine geschrieben. Nichts an der
    Formel verändern oder „verbessern“. Hoch-/tiefgestellter Text `^{…}`/`_{…}` bleibt so stehen.
11. `[zentriert]`/`[rechtsbündig]`/`[eingerückt]` übernehmen (`ausrichtung`, `einzug_cm`).
12. Zeichnungen → sauber nachgezeichnet (Schritt 5b), an der Stelle, an der sie gezeichnet wurden.
13. Sprache bleibt, wie sie geschrieben hat (Englisch bleibt Englisch).

**Aussehen** (macht das Skript automatisch – nicht von Hand nachbauen)
- **Eine Word-Datei pro Stunde.** Ganz oben nur eine kleine graue Datumszeile, dann **genau ihr
  Inhalt**. Keine Titelseite, keine Übersicht, kein Inhaltsverzeichnis, kein Kapitel-/Titelbalken,
  keine Seitenumbrüche zwischen Abschnitten. Unten nur „Seite X von Y“.
- Jedes Fach hat eine eigene Farbe (`SKILL/faecher.json`): ihre Überschriften, Kästen,
  Tabellenköpfe und Aufzählungszeichen in dieser Farbe. Schrift Arial 11.

Will sie das Design ändern (z. B. „mach Mathe lila“, „größere Schrift“): Farben in
`faecher.json`, Layout in `scripts/word_erstellen.js` ändern, dann die betroffenen Stunden-Dateien
neu bauen (nur auf Wunsch auch ältere) und mit der Vorschau kontrollieren. Neues Fach: Zeile in `faecher.json` ergänzen.

## Bausteine (Felder in `bloecke`)

| typ | Felder | Wofür |
|---|---|---|
| `abschnitt` | `text`, optional `ausrichtung` | ihre Überschrift 1 (genau ihr Wortlaut) |
| `unterabschnitt` | `text` | Überschrift 2. Ebene |
| `text` | `text`, optional `ausrichtung` ("mitte" / "rechts" / "blocksatz"), `einzug_cm` | Absatz (Zeilenumbruch `\n` = neuer Absatz, `\t` = Tab) |
| `liste` | `punkte`: Liste aus Text oder `{ "text": "…", "unterpunkte": [ … ] }`, `art`: "punkte" (•) oder "striche" (–) – wie sie es geschrieben hat | Aufzählung |
| `nummern` | `punkte` (wie oben), `art`: "nummern" (1.), "buchstaben" (a)) oder "roemisch" (I.) | Nummerierung |
| `tabelle` | `kopf`: [..], `zeilen`: [[..],[..]], optional `breiten`: [30,70] (Prozent), `unterschrift`, `erste_spalte_fett` (Standard true) | Vergleiche |
| `merke` | `text` und/oder `punkte`, optional `titel` | Definition, Kernaussage |
| `beispiel` | `text` und/oder `punkte`/`nummern`, optional `titel` | Beispiel, Rechenbeispiel |
| `achtung` | `text` und/oder `punkte`, optional `titel` | nur, wenn sie selbst „Achtung“ o. Ä. geschrieben hat |
| `formel` | `latex` (z. B. "\\frac{3}{4}+x^{2}") – oder `text` für einfache Formeln, optional `erklaerung` | Formel in eigener Zeile |
| `bild` | `datei` ("bilder/…png", relativ zu `_Mitschrift-Daten`), `breite_cm` (Standard 12, max. 17), `unterschrift` (nur Bildquelle) | nachgezeichnete Skizze, Foto, Screenshot |
| `begriffe` | `eintraege`: [{ "begriff": "…", "erklaerung": "…" }], `titel: null` | nur, wenn SIE so eine Liste geschrieben hat |
| `zusammenfassung` | `punkte` | nur, wenn SIE eine Zusammenfassung geschrieben hat |

In allen Texten: `**fett**`, `*kursiv*`, `__unterstrichen__`, `==hervorgehoben==` (farbig hinterlegt),
`^{hoch}`, `_{tief}`, `$LaTeX-Formel$` (echte Word-Formel mitten im Satz).
In Tabellenzellen: `\n` = neue Zeile, Zeile beginnt mit `- ` = Aufzählungspunkt.

Bildgrößen-Richtwert: Zeichnungen mit „Breite x %“ im Export → `breite_cm` ≈ x % von 17 cm
(mind. 6 cm). Kleine Skizzen 6–9 cm, normale 10–13 cm, große Diagramme bis 17 cm.

Ein vollständiges Beispiel steht in `SKILL/beispiel/2026-11-16_Das-Internet.json`.

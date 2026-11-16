---
name: mitschrift
description: Verarbeitet Unterrichtsnotizen aus dem Notizblock (Dateien „Mitschrift_*.html“ im gewählten Speicherort/_Eingang oder in Downloads) und hängt sie wortgetreu, aber schön gestaltet an das Word-Dokument des jeweiligen Fachs an – jedes Fach mit eigener Farbe, Zeichnungen sauber nachgezeichnet. Verwenden bei „/mitschrift“, „Mitschrift verarbeiten“, „Notizen in Word“, „neue Stunde eintragen“ oder wenn eine Mitschrift_*.html-Datei auftaucht. Auch zum Ändern des Designs der Mitschriften.
---

# Mitschrift → gestaltetes Word-Dokument

Die Schülerin schreibt im Unterricht im **Notizblock** (Desktop-App) und drückt am Ende
„Stunde beenden → Claude“. Dadurch landet eine Datei `Mitschrift_<Fach>_<Datum>_<Thema>.html`
im Ordner `_Eingang` (oder in `Downloads`). Deine Aufgabe: daraus eine übersichtliche,
gestaltete Mitschrift machen – **wortgetreu, ohne etwas dazuzuschreiben** – und sie an das
Word-Dokument des Fachs anhängen.

**Sprich mit ihr einfach, kurz und auf Deutsch.** Sie hat keinen technischen Hintergrund:
keine Fachbegriffe wie JSON, Skript, Node; Fortschritt in ein bis zwei kurzen Sätzen.

`SKILL` = der Ordner, in dem diese Datei liegt (Skripte unter `SKILL/scripts/`).

## Wie die Daten liegen

```
Schule/2026_27/                 ← Schuljahres-Ordner (enthält _Eingang)
├─ _Eingang/                    ← neue Exporte aus dem Notizblock
│  └─ erledigt/                 ← verarbeitete Exporte (nie löschen)
├─ EBC/
│  ├─ Mitschrift.docx           ← das Word-Dokument (wird jedes Mal komplett neu gebaut)
│  └─ _Mitschrift-Daten/
│     ├─ mitschrift.json        ← Inhalt ALLER Stunden (die „Quelle“ des Word-Dokuments)
│     ├─ bilder/                ← Bilder und Zeichnungen
│     └─ backup/                ← automatische Sicherungen der Word-Datei
└─ Mathe/ …
```

Das Word-Dokument wird immer aus `mitschrift.json` erzeugt. Darum: **nie direkt in der
.docx herumbasteln**, sondern immer `mitschrift.json` ändern und neu bauen. So sehen alle
Stunden immer gleich aus, und Design-Änderungen gelten automatisch für alles.

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

### 4. Fach-Daten öffnen (oder beim ersten Mal anlegen)
Fach-Ordner = `<Schuljahres-Ordner>/<fach aus dem Export>` (anlegen, falls er fehlt).

**Gibt es `_Mitschrift-Daten/mitschrift.json` schon?** → laden, weiter mit Schritt 5.

**Erstes Mal für dieses Fach:**
1. Nachsehen, welche .docx-Dateien im Fach-Ordner liegen.
2. Mit AskUserQuestion einfach fragen: *„In welches Word-Dokument sollen die
   <Fach>-Mitschriften?“* – Optionen: die passenden vorhandenen Dateien (max. 3, die
   wahrscheinlichste zuerst, z. B. eine namens „Mitschrift“ oder wie das Fach) und
   „Neues Dokument ‚Mitschrift <Fach>.docx‘“.
3. Bei einer vorhandenen Datei:
   - Kopie nach `_Mitschrift-Daten/backup/ORIGINAL_<name>.docx` legen (nie überschreiben).
   - `python "SKILL/scripts/docx_lesen.py" "<datei.docx>" "<Fach>/_Mitschrift-Daten/alt-tmp"`
   - Den alten Inhalt **vollständig** in Stunden-Bausteine übertragen (siehe unten):
     pro großem Thema eine „Stunde“; Datum unbekannt → Änderungsdatum der Datei bzw.
     Reihenfolge beibehalten. Bilder übernehmen (EMF/WMF-Grafiken weglassen und kurz erwähnen).
     Nichts weglassen – im Zweifel lieber mehr übernehmen.
4. `mitschrift.json` anlegen:
```json
{ "fach": "EBC", "datei": "Mitschrift.docx", "untertitel": "Mitschrift · Schuljahr 2026/27", "stunden": [] }
```

### 5. Neue Stunde gestalten
Aus den Notizen eine Stunde bauen und in `stunden` einfügen (Regeln: siehe „Gestaltungsregeln“ –
**wortgetreu, nichts dazuschreiben**):
```json
{ "datum": "2026-11-16", "thema": "Das Internet", "quelle": "Mitschrift_EBC_2026-11-16_Das-Internet.html",
  "bloecke": [ … ] }
```
- **Gleiches Fach + gleiches Datum + gleiches Thema schon vorhanden?** Dann wurde die Stunde
  nochmal exportiert (z. B. nach dem Weiterschreiben) → die alte Stunde **ersetzen**, nicht doppelt anlegen.
- Kein Thema angegeben → ihre erste Überschrift als Thema nehmen; gibt es keine, kurz fragen (nichts erfinden).
- Die Stunden werden automatisch nach Datum sortiert.

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
node "SKILL/scripts/word_erstellen.js" "<Fach>/_Mitschrift-Daten/mitschrift.json"
```
- Exit-Code 3 / `FEHLER_DATEI_OFFEN` → sagen: *„Bitte schließ kurz die Word-Datei <name>,
  dann mache ich weiter.“* und danach erneut ausführen.
- Die alte Word-Datei wird automatisch gesichert (`backup/`, die letzten 10).

### 7. Kontrolle (bei der ersten Stunde eines Fachs und bei Design-Änderungen Pflicht)
```bash
powershell -ExecutionPolicy Bypass -File "SKILL/scripts/vorschau.ps1" "<datei.docx>" "<scratchpad>/vorschau.pdf"
```
Die PDF-Seiten der neuen Stunde ansehen: Ist **jedes ihrer Wörter** drin – und **kein fremdes**?
Tabellen lesbar? Zeichnungen richtig groß? Falls nötig `mitschrift.json` korrigieren und neu bauen. (Falls Word nicht
verfügbar ist, diesen Schritt auslassen.)

### 8. Aufräumen und kurz melden
- Export nach `_Eingang/erledigt/` verschieben (aus Downloads ebenfalls dorthin).
- Kurze Rückmeldung, z. B.:
  *„Fertig ✓ EBC → Mitschrift.docx: Stunde vom 16.11. „Das Internet“ ist drin (2 Überschriften,
  1 Mindmap sauber nachgezeichnet, 2 Tippfehler korrigiert).“*
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
10. Formeln, die sie geschrieben hat → `formel` (ohne `erklaerung`, außer sie hat die Erklärung
    selbst geschrieben).
11. `[zentriert]`/`[rechtsbündig]`/`[eingerückt]` übernehmen (`ausrichtung`, `einzug_cm`).
12. Zeichnungen → sauber nachgezeichnet (Schritt 5b), an der Stelle, an der sie gezeichnet wurden.
13. Sprache bleibt, wie sie geschrieben hat (Englisch bleibt Englisch).

**Aussehen** (macht das Skript automatisch – nicht von Hand nachbauen)
- Jedes Fach hat eine eigene Farbe (`SKILL/faecher.json`): Überschriften, Kästen,
  Tabellenköpfe und Aufzählungszeichen in dieser Farbe. Schrift Arial 11.
- Titelseite mit Fachname und anklickbarer Übersicht aller Stunden.
- Jede Stunde beginnt auf einer neuen Seite mit Datum und farbigem Titelbalken.
- Fußzeile mit Fach und „Seite X von Y“.

Will sie das Design ändern (z. B. „mach Mathe lila“, „größere Schrift“): Farben in
`faecher.json`, Layout in `scripts/word_erstellen.js` ändern, dann alle betroffenen Fächer
neu bauen und mit der Vorschau kontrollieren. Neues Fach: Zeile in `faecher.json` ergänzen.

## Bausteine (Felder in `bloecke`)

| typ | Felder | Wofür |
|---|---|---|
| `abschnitt` | `text`, optional `ausrichtung` | ihre Überschrift 1 (genau ihr Wortlaut) |
| `unterabschnitt` | `text` | Überschrift 2. Ebene |
| `text` | `text`, optional `ausrichtung` ("mitte" / "rechts" / "blocksatz"), `einzug_cm` | Absatz (Zeilenumbruch `\n` = neuer Absatz, `\t` = Tab) |
| `liste` | `punkte`: Liste aus Text oder `{ "text": "…", "unterpunkte": [ … ] }` | Aufzählung |
| `nummern` | `punkte` (wie oben) | nummerierte Schritte |
| `tabelle` | `kopf`: [..], `zeilen`: [[..],[..]], optional `breiten`: [30,70] (Prozent), `unterschrift`, `erste_spalte_fett` (Standard true) | Vergleiche |
| `merke` | `text` und/oder `punkte`, optional `titel` | Definition, Kernaussage |
| `beispiel` | `text` und/oder `punkte`/`nummern`, optional `titel` | Beispiel, Rechenbeispiel |
| `achtung` | `text` und/oder `punkte`, optional `titel` | nur, wenn sie selbst „Achtung“ o. Ä. geschrieben hat |
| `formel` | `text` (z. B. "Δf = f(b) − f(a)"), optional `erklaerung` | Formel |
| `bild` | `datei` ("bilder/…png", relativ zu `_Mitschrift-Daten`), `breite_cm` (Standard 12, max. 17), `unterschrift` (nur Bildquelle) | nachgezeichnete Skizze, Foto, Screenshot |
| `begriffe` | `eintraege`: [{ "begriff": "…", "erklaerung": "…" }], `titel: null` | nur, wenn SIE so eine Liste geschrieben hat |
| `zusammenfassung` | `punkte` | nur, wenn SIE eine Zusammenfassung geschrieben hat |

In allen Texten: `**fett**`, `*kursiv*`, `__unterstrichen__`, `==hervorgehoben==` (farbig hinterlegt).
In Tabellenzellen: `\n` = neue Zeile, Zeile beginnt mit `- ` = Aufzählungspunkt.

Bildgrößen-Richtwert: Zeichnungen mit „Breite x %“ im Export → `breite_cm` ≈ x % von 17 cm
(mind. 6 cm). Kleine Skizzen 6–9 cm, normale 10–13 cm, große Diagramme bis 17 cm.

Ein vollständiges Beispiel steht in `SKILL/beispiel/mitschrift.json`.

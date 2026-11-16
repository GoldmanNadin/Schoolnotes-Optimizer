---
name: mitschrift
description: Verarbeitet Unterrichtsnotizen aus dem Notizblock (Dateien „Mitschrift_*.html“ im Ordner _Eingang oder in Downloads) und hängt sie schön gestaltet an das Word-Dokument des jeweiligen Fachs an – jedes Fach mit eigener Farbe. Verwenden bei „/mitschrift“, „Mitschrift verarbeiten“, „Notizen in Word“, „neue Stunde eintragen“ oder wenn eine Mitschrift_*.html-Datei auftaucht. Auch zum Ändern des Designs der Mitschriften.
---

# Mitschrift → gestaltetes Word-Dokument

Die Schülerin schreibt im Unterricht im **Notizblock** (Desktop-App) und drückt am Ende
„Stunde beenden → Claude“. Dadurch landet eine Datei `Mitschrift_<Fach>_<Datum>_<Thema>.html`
im Ordner `_Eingang` (oder in `Downloads`). Deine Aufgabe: daraus eine übersichtliche,
verständliche Lernunterlage machen und sie an das Word-Dokument des Fachs anhängen.

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
- Keine gefunden → freundlich sagen, wo die Datei liegen sollte (`Schule › 2026_27 › _Eingang`)
  und fragen, ob sie woanders gespeichert wurde.
- Mehrere → alle der Reihe nach verarbeiten (nach Datum sortiert).
- Liegt ein Export in Downloads, gehört er zum Schuljahres-Ordner, in dem `_Eingang` liegt.

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
  einbauen (siehe Gestaltungsregel 7).
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
Aus den Notizen eine Stunde bauen und in `stunden` einfügen:
```json
{ "datum": "2026-11-16", "thema": "Das Internet", "quelle": "Mitschrift_EBC_2026-11-16_Das-Internet.html",
  "einleitung": "Ein Satz, worum es in dieser Stunde ging.",
  "bloecke": [ … ] }
```
- **Gleiches Fach + gleiches Datum + gleiches Thema schon vorhanden?** Dann wurde die Stunde
  nochmal exportiert (z. B. nach dem Weiterschreiben) → die alte Stunde **ersetzen**, nicht doppelt anlegen.
- Kein Thema angegeben → ein kurzes, passendes Thema aus dem Inhalt wählen.
- Die Stunden werden automatisch nach Datum sortiert.

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
Die PDF-Seiten der neuen Stunde ansehen: Ist alles drin? Tabellen lesbar? Bilder richtig
groß? Falls nötig `mitschrift.json` korrigieren und neu bauen. (Falls Word nicht
verfügbar ist, diesen Schritt auslassen.)

### 8. Aufräumen und kurz melden
- Export nach `_Eingang/erledigt/` verschieben (aus Downloads ebenfalls dorthin).
- Kurze Rückmeldung, z. B.:
  *„Fertig ✓ EBC → Mitschrift.docx: Stunde vom 16.11. „Das Internet“ ist drin (3 Abschnitte,
  1 Zeichnung, Merke-Kasten zu DNS).“*
  Wenn etwas unklar war (unleserliches Wort, unklare Abkürzung), das in einem Satz erwähnen.

## Gestaltungsregeln (so sollen die Mitschriften IMMER aussehen)

**Ziel:** Eine Lernunterlage, mit der sie vor dem Test gut lernen kann – verständlich,
übersichtlich, vollständig. Wie ein gutes Skriptum, nicht wie ein Roman.

**Inhalt**
1. **Nichts geht verloren.** Jeder Begriff, jede Zahl, jedes Beispiel, jede Skizze aus den
   Notizen kommt vor. Lieber zu viel als zu wenig.
2. **Stichworte → kurze, verständliche Sätze.** Abkürzungen beim ersten Vorkommen
   ausschreiben, z. B. „NIC (Network Interface Card – Netzwerkkarte)“.
3. Rechtschreibung und Tippfehler still korrigieren.
4. Kurze Ergänzungen sind erlaubt, wenn sie das Verstehen erleichtern und sicher stimmen
   (z. B. ein Alltagsbeispiel). Keinen neuen Prüfungsstoff erfinden. Ist etwas in den Notizen
   unklar, nicht raten, sondern im Text mit „(?)“ markieren und in der Rückmeldung erwähnen.
5. Was sie **markiert, farbig, unterstrichen, mit „!!“ oder im Merke-Kasten** geschrieben hat,
   ist wichtig → in einen `merke`-Kasten oder ==hervorheben==. Ihr `[BEISPIEL-KASTEN]` → `beispiel`,
   `[ZITAT]` → `merke` mit Titel „Definition“ (oder kursiver `text`). Ihre Listenart (1. / a) / I. /
   – / •) beibehalten: Nummerierungen bleiben `nummern`.
6. Sprache: Deutsch. Im Fach Englisch bleibt der Inhalt auf Englisch (Erklärungen dürfen kurz
   auf Deutsch dazu).

**Aufbau jeder Stunde**
1. `einleitung`: ein Satz, worum es geht.
2. 2–6 `abschnitt`e mit nummerierten Überschriften („1. …“, „2. …“), bei Bedarf `unterabschnitt`.
3. Aufzählungen als `liste`/`nummern` (Reihenfolgen, Schritte → `nummern`).
4. Vergleiche, Gegenüberstellungen, „A vs. B“ → `tabelle`.
5. Definitionen und Kernregeln → `merke`. Beispiele → `beispiel`. Typische Fehler /
   „kommt zum Test“ → `achtung`. Formeln → `formel` (mit Erklärung der Variablen).
6. Zeichnungen und Bilder → `bild` mit Bildunterschrift an der passenden Stelle. Steckt in einer
   Skizze Inhalt (z. B. ein Ablauf), diesen zusätzlich kurz in Worten oder als Tabelle erklären.
7. **Handschrift (`[HANDSCHRIFT …]`):** Den Inhalt immer in den Text übernehmen – eine Mindmap
   um „Internet“ mit „WWW“, „E-Mail“, „FTP“ wird z. B. eine Liste „Dienste des Internets“ (mit
   Unterpunkten) oder eine Tabelle; Pfeile werden zu „→“-Beziehungen; eingekreiste/unterstrichene
   Wörter sind wichtig (→ hervorheben). Ist es eine echte Skizze/Mindmap, zusätzlich den Ausschnitt
   als `bild` einfügen (Unterschrift z. B. „Mindmap aus dem Unterricht“). Nur Unterstreichungen,
   Kreise oder Häkchen → kein Bild nötig, nur hervorheben.
8. Ab 3 neuen Fachbegriffen am Ende: `begriffe` („Wichtige Begriffe“).
9. Zum Schluss immer: `zusammenfassung` mit 3–5 Punkten.

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
| `abschnitt` | `text` | Überschrift 1. Ebene in der Stunde („1. Grundbegriffe“) |
| `unterabschnitt` | `text` | Überschrift 2. Ebene |
| `text` | `text` | Absatz (Zeilenumbruch `\n` = neuer Absatz) |
| `liste` | `punkte`: Liste aus Text oder `{ "text": "…", "unterpunkte": [ … ] }` | Aufzählung |
| `nummern` | `punkte` (wie oben) | nummerierte Schritte |
| `tabelle` | `kopf`: [..], `zeilen`: [[..],[..]], optional `breiten`: [30,70] (Prozent), `unterschrift`, `erste_spalte_fett` (Standard true) | Vergleiche |
| `merke` | `text` und/oder `punkte`, optional `titel` | Definition, Kernaussage |
| `beispiel` | `text` und/oder `punkte`/`nummern`, optional `titel` | Beispiel, Rechenbeispiel |
| `achtung` | `text` und/oder `punkte`, optional `titel` | Fehlerquelle, Testhinweis |
| `formel` | `text` (z. B. "Δf = f(b) − f(a)"), optional `erklaerung` | Formel |
| `bild` | `datei` ("bilder/…png", relativ zu `_Mitschrift-Daten`), `breite_cm` (Standard 12, max. 17), `unterschrift` | Zeichnung, Foto, Screenshot |
| `begriffe` | `eintraege`: [{ "begriff": "…", "erklaerung": "…" }], optional `titel` (`null` = ohne Überschrift) | Glossar |
| `zusammenfassung` | `punkte` | „Kurz zusammengefasst“ am Ende |

In allen Texten: `**fett**`, `*kursiv*`, `__unterstrichen__`, `==hervorgehoben==` (farbig hinterlegt).
In Tabellenzellen: `\n` = neue Zeile, Zeile beginnt mit `- ` = Aufzählungspunkt.

Bildgrößen-Richtwert: Zeichnungen mit „Breite x %“ im Export → `breite_cm` ≈ x % von 17 cm
(mind. 6 cm). Kleine Skizzen 6–9 cm, normale 10–13 cm, große Diagramme bis 17 cm.

Ein vollständiges Beispiel steht in `SKILL/beispiel/mitschrift.json`.

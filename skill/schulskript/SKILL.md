---
name: schulskript
description: "Erstellt Schul-Mitschriften, Skripte und Lernunterlagen im festen Design - Inhaltsverzeichnis, Merke-Boxen, Tabellen, Fusszeile - mit einer eigenen Farbe pro Fach. Nutze diesen Skill immer, wenn jemand eine neue Mitschrift, ein Skript, eine Lernunterlage, Zusammenfassung oder ein neues Fach fuer die Schule anlegen, einen Mitschrift-Inhalt (Foto, Text, Stichworte) schoen aufbereiten oder ein bestehendes Fach-Dokument im gleichen Stil fortsetzen will, auch wenn das Wort Skill oder Design nicht vorkommt."
---

# Schulskript

Baut aus dem Inhalt einer Schulstunde ein gestaltetes Lerndokument. Der Stil ist aus zwei bestehenden Mitschriften abgeleitet (EBC in Gruen, Naturwissenschaften in Blau) und soll bei jedem Fach gleich aussehen. Nur die **Fachfarbe** wechselt. Dadurch erkennt man ein Fach sofort am Dokument.

## Der gelieferte Inhalt bleibt unveraendert (wichtigste Regel)

Wenn eine Datei, Zeichnung, ein Foto oder ein Text geliefert wird, wird der Text **einfach abgetippt**, genau so wie er dasteht. Grafisch verbessert werden nur **Zeichnungen und Formate** (Layout, Ausrichtung, Boxen, Tabellen, Schaubilder). Der Inhalt gehoert der Person, die ihn geliefert hat, und wird nicht angefasst.
- Woerter und Formulierungen exakt uebernehmen. Keine Saetze umschreiben, kuerzen oder "verbessern".
- Keine Saetze, Erklaerungen, Beispiele, Definitionen oder sonstigen Kontext hinzufuegen.
- **Keine Aufzaehlungspunkte, Nummerierungen oder Ueberschriften erfinden.** Aufzaehlungspunkte nur dort, wo in der Vorlage selbst welche stehen. Steht etwas als normaler Text oder Zeile da, bleibt es normaler Text oder Zeile.
- Keine Merke-Box, Kurzuebersicht, Fazit oder Tabelle, die in der Vorlage nicht vorgesehen ist. Hervorhebungen (fett, markiert) nur, wo sie schon in der Vorlage sind.
- Bei einer Zeichnung (Skizze, Schaubild, Mindmap, Tabelle): Aufbau, Beschriftungen und Aussage uebernehmen, nur sauberer zeichnen, ausrichten und gliedern. Nichts Neues einzeichnen oder beschriften.
- Offensichtliche Tippfehler duerfen still korrigiert werden. Alles andere (Unklarheiten, moegliche Fehler, Luecken) nicht selbst aendern, sondern am Ende in einem Satz nennen.
- Reihenfolge der Inhalte beibehalten, ausser es wird um Umordnung gebeten.
- Nach dem Bauen den Text im Ergebnis mit der Vorlage vergleichen und sicherstellen, dass nichts dazugekommen oder weggefallen ist, auch keine Aufzaehlungspunkte.

## Formatierung und Zeichnungen komplett uebernehmen

Nicht nur der Text, auch die **Formatierung der Vorlage** wird uebernommen. Es reicht nicht, alles untereinander, linksbuendig und schwarz zu schreiben.
- **Position und Anordnung 1:1:** Was in der Vorlage rechts steht, steht rechts. Was nebeneinander steht, bleibt nebeneinander. Eingerueckte Zeilen, Abstaende, Zeilenumbrueche und Reihenfolge bleiben wie in der Vorlage.
- **Schriftfarben und Auszeichnungen 1:1:** Ein rotes Wort (z. B. ein rotes k in einer Formel) bleibt rot. Fett, kursiv, unterstrichen und unterschiedliche Groessen bleiben wie in der Vorlage. Die Fachfarbe gilt nur fuer das Dokumentdesign (Ueberschriften, Boxen, Tabellenkoepfe), nicht fuer Farben, die in der Vorlage selbst gesetzt wurden.
- **Text und Zeichnung gehoeren zusammen:** Wurde zu getipptem Text etwas dazugezeichnet (Pfeile, Kreise, Unterstreichungen, Klammern, Beschriftungen wie "Konstante"), wird dieser ganze Bereich als **ein Block** nachgebaut, in dem Text und Zeichnung zusammen sitzen. Die Zeichnung ist nie ein loses Bild unter oder neben dem Text.
- **Pfeile und Markierungen treffen dieselben Stellen wie in der Vorlage:** Zeigt ein Pfeil in der Vorlage auf das e in der Formel, zeigt er auch im Ergebnis auf genau dieses e. Zeigt ein zweiter auf das e in der Erklaerung rechts, zeigt er dort hin. Pfeilform und -farbe wie in der Vorlage, nur sauberer gezeichnet. Beschriftungen (z. B. "Konstante") behalten Position, Farbe und ungefaehre Groesse.
- **Umsetzung solcher Bloecke:** Den Block als SVG oder HTML bauen (Text an den Positionen der Vorlage, Pfeile als saubere Vektoren, Koordinaten aus der Vorlage ablesen), in hoher Aufloesung (mindestens doppelt) als PNG rendern und ins Markdown als Bild einbinden. Schrift passend zum Dokument. Der Text im Block ist exakt abgetippt. Fuer Bloecke keine Bildunterschrift hinzufuegen. Alles ohne Sonderlayout bleibt normaler Text im Dokument.
- **Pruefen:** Das gebaute Bild neben die Vorlage legen und vergleichen: Positionen, Farben, Auszeichnungen, Zielpunkte der Pfeile, Text unveraendert.

## Kein Deckblatt

Es wird **nie** ein Deckblatt erstellt, bei keinem Fach und keinem Dokument, auch nicht als Vorschlag oder Angebot. Das Dokument beginnt direkt mit dem Inhalt (bzw. dem Inhaltsverzeichnis). Im Markdown-Kopf immer `deckblatt: false` setzen.

## Ablauf

1. **Fach und Inhalt klaeren.** Fach, Klasse und Thema stehen meist in der Anfrage. Fehlt etwas Wesentliches (Fach!), eine kurze Frage stellen. Der Inhalt kommt aus Text, Stichworten, Fotos oder Dokumenten. Fotos und PDFs zuerst lesen und dabei auch auf Position, Farben und Zeichnungen achten. Wurde der Inhalt selbst geliefert (Zeichnung, Foto, Text, Datei), gelten die Regeln "Der gelieferte Inhalt bleibt unveraendert" und "Formatierung und Zeichnungen komplett uebernehmen" oben.
2. **Fachfarbe bestimmen.**
   - Zuerst im Gedaechtnis nachsehen, ob es `/areas/schulskript-faecher.md` gibt, und dort die Fach-Farb-Zuordnung lesen.
   - Dann `python scripts/pick_color.py --fach "<Fach>" --used <Hex-Werte aus dem Gedaechtnis>` ausfuehren. Hat das Fach schon eine Farbe, wird sie wiederverwendet; sonst waehlt das Skript zufaellig eine freie Farbe aus der Palette (`faecher.json` kennt EBC = Gruen `1F7A5C`, Naturwissenschaften = Blau `0070C0`).
   - Bei einem neuen Fach die Zuordnung in `/areas/schulskript-faecher.md` ergaenzen (Datei neu anlegen, falls es sie nicht gibt). Skill-Dateien sind nicht dauerhaft beschreibbar, das Gedaechtnis ist der Ort, an dem die Farben ueber Sitzungen hinweg erhalten bleiben. Die gewaehlte Farbe in einem Satz nennen.
3. **Markdown schreiben** (siehe Format unten). Die `.md` ist die bearbeitbare Quelle. Bloecke mit Sonderlayout oder Zeichnung als Bild einbinden (siehe oben).
4. **Dokument bauen:** `python scripts/build_skript.py mitschrift.md --farbe <HEX> --out <ordner> --pdf` erzeugt Word und PDF mit Inhaltsverzeichnis samt Seitenzahlen und Fusszeile, ohne Deckblatt (`deckblatt: false` im Kopf).
5. **Pruefen.** PDF in Bilder umwandeln (`pdftoppm -r 60 -png`) und ansehen: Layout, Tabellen, Boxen, Bilder, Seitenzahlen im Inhaltsverzeichnis, und dass **kein Deckblatt** vorne steht. Dazu alles mit der Vorlage vergleichen: Text (nichts hinzugefuegt, umgeschrieben oder weggelassen, keine zusaetzlichen Aufzaehlungspunkte), Positionen, Farben, Auszeichnungen und Pfeilziele. Erst dann abgeben.
6. **Abgeben** mit SendUserFile: `.md`, `.docx` und `.pdf`. Kurz sagen, welche Farbe das Fach bekommen hat.

## Markdown-Format

Kopf (Frontmatter) und Inhalt:

```
---
fach: Mathematik
klasse: 3. Klasse
titel: Lineare Funktionen      # optional
autor: Max Muster
kapitelseiten: false           # true = jedes "# " bekommt eine eigene Titelseite (ausfuehrliche Skripte)
deckblatt: false               # IMMER false, es gibt kein Deckblatt
inhaltsverzeichnis: true       # optional, Standard true
---
# Kapitel          -> Hauptueberschrift (neue Seite)
## Abschnitt
### Unterabschnitt
**fett** fuer Fachbegriffe, ==markiert== fuer das Allerwichtigste (leicht getoent)
- Aufzaehlung / 1. Nummerierung
| Tabelle | mit | Kopfzeile |
![Bildtext](bild.png)
> [!merke] Kurze Regel, Definition oder Formel zum Auswendiglernen
> [!achtung] Typische Fehler und Korrekturen
> [!fazit] Zusammenfassung am Kapitelende
> [!beispiel] Rechenbeispiel
---                    -> Seitenumbruch
```

## Inhaltlich gut schreiben (nur wenn Inhalt von Grund auf erstellt werden soll)

Diese Hinweise gelten nur, wenn ausdruecklich gewuenscht ist, dass Inhalt neu geschrieben oder ausgebaut wird (z. B. nur Thema oder Stichworte, "schreib mir ein Skript zu ..."). Bei gelieferten Dateien oder Zeichnungen gelten die Regeln oben. Gute Ergebnisse haben dann:
- kurze Saetze und Stichpunkte statt Fliesstext, Fachbegriffe **fett**, damit man sie beim Lernen findet
- pro Kapitel mindestens eine Merke-Box (Kernaussage) und, wo es passt, eine Tabelle zum Vergleichen
- bei Mathe/Rechnen jeden Schritt mit konkreten Zahlen zeigen, keinen Zwischenschritt ueberspringen
- Grafiken nur einbauen, wenn sie vorhanden oder wirklich hilfreich sind; Bildunterschrift immer setzen
- am Ende bei laengeren Dokumenten eine Kurzuebersicht (Tabelle Begriff / kurz erklaert) wie in der EBC-Mitschrift

Design-Details (Masse, Schrift, Boxen) stehen in `references/design.md`, nur bei Layout-Fragen oder Anpassungen lesen. Die dortigen Angaben zum Deckblatt gelten nicht mehr.
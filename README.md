# ✎ Unterrichts-Notizblock → Claude → Word

Im Unterricht schnell mitschreiben und zeichnen – am Ende der Stunde **ein Klick**, und Claude
macht daraus eine schön gestaltete, übersichtliche Word-Lernunterlage. Jedes Fach bekommt
sein eigenes Design mit eigener Farbe, und alle Stunden eines Fachs landen in **einem**
Word-Dokument.

```
Notizblock (Browser)  ──„Stunde beenden → Claude“──▶  _Eingang/Mitschrift_….html
                                                              │
                                     in Claude: /mitschrift   ▼
                                  Fach/Mitschrift.docx  ◀── Skill „mitschrift“
```

## Was kann der Notizblock?

- Schreiben wie in Word: Überschriften, **fett**, *kursiv*, Farben, Textmarker, Aufzählungen,
  Tabellen und „💡 Merke“-Kästen
- **Mit der Maus zeichnen**: Stift, Textmarker, Radierer, Linie, Pfeil, Rechteck, Kreis, Text –
  auf weißem, kariertem oder liniertem Papier
- Bilder einfügen: Knopf „Bild“, **Strg+V** (z. B. Screenshots) oder einfach hineinziehen –
  und auf Bildern herummalen („✎ Bemalen“)
- Speichert automatisch im Browser – nichts geht verloren, auch wenn das Fenster zugeht
- Fach, Thema und Datum oben auswählen

## So benutzt du es

1. **Notizblock öffnen** – `index.html` doppelklicken (am besten in Edge oder Chrome)
   oder die Online-Version über GitHub Pages verwenden.
2. **Mitschreiben**, zeichnen, Bilder einfügen.
3. Am Ende: **„Stunde beenden → Claude“** klicken.
   Beim ersten Mal einmalig den Ordner `Schule › 2026_27 › _Eingang` auswählen.
4. In der **Claude-App** (Ordner `Schule/2026_27`) eintippen: **`/mitschrift`**
   → Claude hängt die Stunde gestaltet an das Word-Dokument des Fachs an.

Tipp: Wenn du nach dem Export weiterschreibst und nochmal exportierst, ersetzt Claude die
Stunde einfach – es entsteht nichts doppelt.

## Einrichtung des Claude-Skills (einmalig)

Der Ordner [`skill/mitschrift`](skill/mitschrift) enthält den Skill, der festlegt, wie die
Word-Dokumente aussehen. Installieren: Rechtsklick auf
[`skill/installieren.ps1`](skill/installieren.ps1) → **„Mit PowerShell ausführen“**.
Voraussetzung: Python und Node.js sind installiert.

## Design anpassen

- **Farben pro Fach:** [`skill/mitschrift/faecher.json`](skill/mitschrift/faecher.json)
- **Regeln, wie Claude gestaltet** (Merke-Kästen, Zusammenfassung, Tabellen …):
  Abschnitt „Gestaltungsregeln“ in [`skill/mitschrift/SKILL.md`](skill/mitschrift/SKILL.md)
- Oder einfach Claude sagen: *„Mach Mathe lila“* oder *„Lass die Zusammenfassung weg“*.

Nach Änderungen im Ordner `skill` nochmal `installieren.ps1` ausführen.

## Wo liegt was?

| Datei / Ordner | Inhalt |
|---|---|
| `index.html` | der Notizblock (eine einzige Datei, läuft offline im Browser) |
| `skill/mitschrift/SKILL.md` | Anleitung + Gestaltungsregeln für Claude |
| `skill/mitschrift/scripts/` | Hilfsprogramme: Export lesen, Word bauen, Vorschau |
| `skill/mitschrift/beispiel/` | Beispiel-Export und Beispiel-Inhalt |

Im Schulordner legt der Skill pro Fach einen Ordner `_Mitschrift-Daten` an (Inhalt aller
Stunden, Bilder, automatische Sicherungen der Word-Datei). Den bitte nicht löschen.

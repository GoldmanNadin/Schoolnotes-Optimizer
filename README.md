# ✎ Notizblock → Claude → Word

Eine Desktop-App für den Unterricht: tippen wie in Word, **mit der Maus direkt aufs Blatt
zeichnen** (z. B. eine Mindmap um ein getipptes Wort), Bilder einfügen – und am Ende der Stunde
**ein Klick**: Claude macht daraus eine schön gestaltete Word-Lernunterlage. Jedes Fach hat sein
eigenes Design; jede Stunde wird ein eigenes Word-Dokument im Fach-Ordner – nur dein Text, ohne Titelseite.

```
Notizblock (Desktop-App) ──„Stunde beenden → Claude“──▶ Schule/<Schuljahr>/_Eingang/Mitschrift_….html
                                                                  │
                                         in Claude: /mitschrift   ▼
                         <Fach>/2026-11-16 Das Internet.docx ◀── Skill „mitschrift“
```

## Herunterladen und installieren

1. Rechts unter **„Releases“** die neueste Version öffnen.
2. **`Notizblock-Setup-….exe`** herunterladen und doppelklicken.
   (Windows fragt eventuell „Unbekannter Herausgeber“ → „Weitere Informationen“ → „Trotzdem ausführen“.)
3. Fertig – auf dem Desktop liegt jetzt **Notizblock**.

Die **Claude-Skills** liegen dort ebenfalls als eigene Dateien:
- **`Skill-mitschrift.zip`** – macht aus einer Notizblock-Stunde ein Word-Dokument
- **`Skill-schulskript.zip`** – macht aus Foto, Text oder Stichworten ein gestaltetes Skript

In claude.ai unter *Einstellungen → Fähigkeiten* hochladen, oder in Claude Code entpacken und den
Ordner nach `~/.claude/skills/` kopieren.

Oder selbst bauen: Rechtsklick auf [`bauen.ps1`](bauen.ps1) → „Mit PowerShell ausführen“
(braucht [Node.js](https://nodejs.org)).

## Was kann der Notizblock?

**Schreiben wie in Word**
- Formatvorlagen: *Standard*, *Überschrift 1–3*, *Merke*, *Beispiel*, *Zitat*
  (Strg+Alt+1/2/3, Strg+Umschalt+N)
- Aufzählungen mit **Punkten** oder **Strichen**, Nummerierungen **1. 2. 3.**, **a) b) c)**,
  **I. II. III.** – mit Unterebenen (Tab / Umschalt+Tab)
- Wie in Word: „- “ oder „1. “ am Zeilenanfang startet automatisch eine Liste,
  zweimal Enter beendet sie
- **Tab-Taste wie in Word**: springt auf Tabstopps (alle 1,25 cm) – gut für Spalten wie „Begriff ⇥ Erklärung“
- Ausrichtung **links, zentriert, rechts, Blocksatz** (Strg+L / E / R / J), Einzug vergrößern/verkleinern
- **Überall schreiben**: Doppelklick auf eine freie Stelle → dort ein Textfeld (z. B. um Pfeile zu
  beschriften). Textfelder anklicken = bearbeiten, mit der Maus ziehen = verschieben
- **Hochzahlen im Text**: einfach „^“ tippen (m^2 → m²), Leertaste oder Tab macht normal weiter
- **Einrahmen** wie die Lehrerin: Formatvorlage **Rahmen** für ganze Absätze, Knopf **[ab]** für markierte Wörter
- Fett, kursiv, unterstrichen, **hoch-/tiefgestellt** (m², H₂O – Strg+Umschalt+Plus / Strg+#),
  Textfarben, Textmarker, Tabellen (Tab springt zur nächsten Zelle)

**Eigene Fächer mit Typ**
- Fächer legst du selbst an: Name, Farbe und Art – **📝 Theorie** (viel Text) oder
  **🔢 Mathe / Naturwissenschaft** (zusätzlich die Formel-Leiste). Verwalten unter „⋯ Mehr“.

**Formeln (in Mathe-Fächern)**
- Knopf **√x Formel** (oder Strg+M / Alt+=) öffnet ein Formel-Feld direkt im Text
- darin: **„/“ macht einen Bruch** (was davor steht, kommt nach oben), **„^“ eine Hochzahl**, „_“ einen
  Index – **Tab** springt wieder heraus, ganz außen zurück in den normalen Text
- Knöpfe für Bruch, Hochzahl, Wurzel, Klammer, ·, ±, ≤, ≥, ≠, ≈, π, Δ, ∞ und eine Formel-Tastatur ⌨
- Steht die Formel allein in einer Zeile, wird sie automatisch zentriert
- Fertige Formel **anklicken**, um sie zu ändern
- Im Word-Dokument werden daraus **echte, bearbeitbare Word-Formeln**
- Rechtschreibprüfung mit Vorschlägen per Rechtsklick

**Zeichnen direkt auf dem Blatt** – ohne extra Fenster
- Stift, Textmarker, Pfeil, Linie, Kreis/Rahmen, Radierer, Farben, Strichstärken
- **T**: Text irgendwo aufs Blatt schreiben – ideal für Mindmaps mit der Maus
- Zeichnungen „kleben“ an ihrem Absatz und wandern mit, wenn darüber Text dazukommt;
  zeichnest du unter den Text, rutscht weiterer Text automatisch darunter
- **Esc** = zurück zum Tippen · mit einem Grafik-Stift kann man jederzeit direkt schreiben

**Bilder**
- vom Computer, mit **Strg+V** (Screenshots) oder hineinziehen
- **Onlinebild**: freie Bilder suchen (Wikimedia Commons, mit Quellenangabe) oder eine
  Bild-Adresse aus dem Internet einfügen; Bilder aus dem Browser kann man auch einfach hineinziehen

**Sicher**: speichert laufend automatisch – nichts geht verloren, auch nicht beim Schließen.

## So benutzt du es

1. **Notizblock** auf dem Desktop öffnen (beim ersten Mal deine Fächer anlegen), Fach + Thema wählen, mitschreiben und zeichnen.
2. Am Ende: **„Stunde beenden → Claude“**. Die Stunde landet automatisch in
   den Ordner, den du beim ersten Start wählst (oben über den 📁-Knopf jederzeit änderbar).
3. In der **Claude-App** (Ordner `Schule/<Schuljahr>`) eintippen: **`/mitschrift`**.

Claude bekommt den Text **und** Ausschnitte deiner Zeichnungen samt dem Text darunter – so
versteht es z. B. eine Mindmap und baut sie sinnvoll ins Word-Dokument ein.

## Einrichtung des Claude-Skills (einmalig)

Rechtsklick auf [`skill/installieren.ps1`](skill/installieren.ps1) → **„Mit PowerShell ausführen“**.
Voraussetzung: Python und Node.js.

## Zweiter Skill: „schulskript“

Im Ordner [`skill/schulskript`](skill/schulskript) liegt ein zusätzlicher Claude-Skill. Er macht aus
einem **Foto, Text oder Stichworten** einer Schulstunde ein gestaltetes Skript (Word + PDF) – mit
Inhaltsverzeichnis, Merke-Boxen, Tabellen und einer eigenen Farbe pro Fach, **ohne Deckblatt**.
Auch hier gilt: Der gelieferte Text wird **wortgetreu** übernommen, nur Layout und Zeichnungen
werden schöner.

- Ablauf und Regeln: [`skill/schulskript/SKILL.md`](skill/schulskript/SKILL.md)
- Benötigt Python mit `python-docx`; für das PDF und das Inhaltsverzeichnis mit Seitenzahlen LibreOffice
  (`soffice`) und Poppler (`pdfinfo`, `pdftotext`)
- Installieren: den Ordner `skill/schulskript` nach `~/.claude/skills/` kopieren – oder als ZIP
  in claude.ai unter *Einstellungen → Fähigkeiten* hochladen

## Design anpassen

- **Farben pro Fach:** [`skill/mitschrift/faecher.json`](skill/mitschrift/faecher.json)
- **Gestaltungsregeln** (wortgetreu, nichts dazuschreiben, Zeichnungen sauber nachzeichnen …):
  [`skill/mitschrift/SKILL.md`](skill/mitschrift/SKILL.md)
- Oder einfach Claude sagen: *„Mach Mathe lila“*. Danach `skill/installieren.ps1` nochmal ausführen.

## Aufbau

| Ordner / Datei | Inhalt |
|---|---|
| `app/index.html` | der Editor (läuft notfalls auch im Browser, dann ohne Zeichnungs-Übergabe) |
| `desktop/` | Desktop-App (Electron): Fenster, Speichern, Export, Bilder aus dem Internet |
| `build/` | App-Symbol |
| `bauen.ps1` | baut die App und installiert sie |
| `skill/mitschrift/` | Claude-Skill: Anleitung, Gestaltungsregeln, Skripte, Beispiel |
| `skill/schulskript/` | Claude-Skill: Skript aus Foto/Text/Stichworten (Word + PDF) |
| `.github/workflows/installer.yml` | baut bei neuer Version automatisch den Installer für „Releases“ |

## Lizenz

Dieses Projekt ist **Open Source** unter der [MIT-Lizenz](LICENSE) – © 2026 Nadin Zerlauth.
Du darfst es frei benutzen, verändern und weitergeben, solange der Lizenzhinweis erhalten bleibt.

Verwendete fremde Bausteine:
- [MathLive](https://mathlive.io) (Formel-Feld) – MIT-Lizenz, siehe [`app/vendor/mathlive/LICENSE.txt`](app/vendor/mathlive/LICENSE.txt)
- [Electron](https://www.electronjs.org) und [docx](https://github.com/dolanmiu/docx) – MIT-Lizenz (werden beim Bauen geladen)
- Bilder aus der Onlinebild-Suche stammen von [Wikimedia Commons](https://commons.wikimedia.org) und behalten ihre eigene Lizenz (wird beim Einfügen angezeigt)

## Mitmachen

Ideen, Fehler oder Verbesserungen? Einfach unter **Issues** melden oder einen **Pull Request** schicken.

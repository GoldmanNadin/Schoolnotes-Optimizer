# ✎ Notizblock → Claude → Word

Eine Desktop-App für den Unterricht: tippen wie in Word, **mit der Maus direkt aufs Blatt
zeichnen** (z. B. eine Mindmap um ein getipptes Wort), Bilder einfügen – und am Ende der Stunde
**ein Klick**: Claude macht daraus eine schön gestaltete Word-Lernunterlage. Jedes Fach hat sein
eigenes Design, alle Stunden eines Fachs landen in **einem** Word-Dokument.

```
Notizblock (Desktop-App) ──„Stunde beenden → Claude“──▶ Schule/<Schuljahr>/_Eingang/Mitschrift_….html
                                                                  │
                                         in Claude: /mitschrift   ▼
                                       <Fach>/Mitschrift.docx ◀── Skill „mitschrift“
```

## Herunterladen und installieren

1. Rechts unter **„Releases“** die neueste Version öffnen.
2. **`Notizblock-Setup-….exe`** herunterladen und doppelklicken.
   (Windows fragt eventuell „Unbekannter Herausgeber“ → „Weitere Informationen“ → „Trotzdem ausführen“.)
3. Fertig – auf dem Desktop liegt jetzt **Notizblock**.

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
- Fett, kursiv, unterstrichen, Textfarben, Textmarker, Tabellen (Tab springt zur nächsten Zelle)
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

1. **Notizblock** auf dem Desktop öffnen, Fach + Thema wählen, mitschreiben und zeichnen.
2. Am Ende: **„Stunde beenden → Claude“**. Die Stunde landet automatisch in
   `OneDrive › Schule › <neuestes Schuljahr> › _Eingang` (änderbar unter „⋯ Mehr“).
3. In der **Claude-App** (Ordner `Schule/<Schuljahr>`) eintippen: **`/mitschrift`**.

Claude bekommt den Text **und** Ausschnitte deiner Zeichnungen samt dem Text darunter – so
versteht es z. B. eine Mindmap und baut sie sinnvoll ins Word-Dokument ein.

## Einrichtung des Claude-Skills (einmalig)

Rechtsklick auf [`skill/installieren.ps1`](skill/installieren.ps1) → **„Mit PowerShell ausführen“**.
Voraussetzung: Python und Node.js.

## Design anpassen

- **Farben pro Fach:** [`skill/mitschrift/faecher.json`](skill/mitschrift/faecher.json)
- **Gestaltungsregeln** (Merke-Kästen, Zusammenfassung, Umgang mit Zeichnungen …):
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
| `.github/workflows/installer.yml` | baut bei neuer Version automatisch den Installer für „Releases“ |

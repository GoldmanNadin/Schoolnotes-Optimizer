"use strict";
/*
 * Notizblock – Desktop-App (Electron)
 * Hauptprozess: Fenster, Speichern auf der Festplatte, Export in den _Eingang-Ordner,
 * Bilder aus dem Internet laden, Handschrift-Ausschnitte für Claude rendern.
 */
const { app, BrowserWindow, ipcMain, dialog, shell, Menu, net, session, clipboard, nativeImage } = require("electron");
const path = require("path");
const fs = require("fs");
const os = require("os");

app.setAppUserModelId("io.github.goldmannadin.notizblock");
// Für Tests: eigener Datenordner / Eingang, damit echte Notizen unberührt bleiben
if (process.env.NOTIZBLOCK_TEST_DATEN) app.setPath("userData", process.env.NOTIZBLOCK_TEST_DATEN);
if (!app.requestSingleInstanceLock()) { app.quit(); }

const DATEN = app.getPath("userData");
const DATEI = {
  entwurf: path.join(DATEN, "entwurf.json"),
  letzte: path.join(DATEN, "letzte-stunde.json"),
  einstellungen: path.join(DATEN, "einstellungen.json"),
};
const USER_AGENT = "Notizblock/2.0 (Schul-Notizen-App; Electron)";

/* ---------------- Dateien ---------------- */
function lesen(datei, standard = null) {
  try { return JSON.parse(fs.readFileSync(datei, "utf8")); } catch { return standard; }
}
function schreiben(datei, daten) {
  fs.mkdirSync(path.dirname(datei), { recursive: true });
  const tmp = datei + ".tmp";
  fs.writeFileSync(tmp, typeof daten === "string" ? daten : JSON.stringify(daten), "utf8");
  fs.renameSync(tmp, datei);
}
const einstellungen = () => lesen(DATEI.einstellungen, {});
const einstellungSetzen = (k, v) => { const e = einstellungen(); e[k] = v; schreiben(DATEI.einstellungen, e); };

/* Eingangsordner automatisch finden: <OneDrive>/Schule/<Schuljahr>/_Eingang (neuestes Schuljahr) */
function eingangFinden() {
  if (process.env.NOTIZBLOCK_EINGANG) return process.env.NOTIZBLOCK_EINGANG;
  const gespeichert = einstellungen().eingang;
  if (gespeichert && fs.existsSync(gespeichert)) return gespeichert;
  const wurzeln = [process.env.OneDriveCommercial, process.env.OneDrive, path.join(os.homedir(), "OneDrive"), path.join(os.homedir(), "Documents")].filter(Boolean);
  for (const w of wurzeln) {
    const schule = path.join(w, "Schule");
    if (!fs.existsSync(schule)) continue;
    const jahre = fs.readdirSync(schule).filter(n => /^\d{4}_\d{2}$/.test(n)).sort().reverse();
    for (const j of jahre) {
      const e = path.join(schule, j, "_Eingang");
      if (fs.existsSync(e)) { einstellungSetzen("eingang", e); return e; }
    }
  }
  return null;
}
async function eingangWaehlen(fenster) {
  const r = await dialog.showOpenDialog(fenster, {
    title: "Ordner _Eingang wählen (Schule › Schuljahr › _Eingang)",
    properties: ["openDirectory", "createDirectory"],
    defaultPath: process.env.OneDriveCommercial || process.env.OneDrive || os.homedir(),
  });
  if (r.canceled || !r.filePaths[0]) return null;
  einstellungSetzen("eingang", r.filePaths[0]);
  return r.filePaths[0];
}

/* ---------------- Fenster ---------------- */
let fenster = null;
function fensterErstellen() {
  fenster = new BrowserWindow({
    width: 1320, height: 920, minWidth: 900, minHeight: 600,
    title: "Notizblock", backgroundColor: "#F3F4F6", show: false, autoHideMenuBar: true,
    icon: path.join(__dirname, "..", "build", "icon.png"),
    webPreferences: { preload: path.join(__dirname, "preload.js"), contextIsolation: true, nodeIntegration: false, sandbox: true, spellcheck: true },
  });
  Menu.setApplicationMenu(null);
  fenster.loadFile(path.join(__dirname, "..", "app", "index.html"));
  fenster.once("ready-to-show", () => { fenster.maximize(); fenster.show(); });

  // Links nie im App-Fenster öffnen, sondern im normalen Browser
  fenster.webContents.setWindowOpenHandler(({ url }) => { if (/^https?:/.test(url)) shell.openExternal(url); return { action: "deny" }; });
  fenster.webContents.on("will-navigate", (e, url) => { if (!url.startsWith("file:")) { e.preventDefault(); if (/^https?:/.test(url)) shell.openExternal(url); } });

  // Rechtsklick-Menü (Rechtschreibung, Ausschneiden/Kopieren/Einfügen, Bild kopieren)
  fenster.webContents.on("context-menu", (_e, p) => {
    const vorlage = [];
    if (p.misspelledWord) {
      for (const v of p.dictionarySuggestions.slice(0, 6)) vorlage.push({ label: v, click: () => fenster.webContents.replaceMisspelling(v) });
      if (!p.dictionarySuggestions.length) vorlage.push({ label: "Keine Vorschläge", enabled: false });
      vorlage.push({ label: "Zum Wörterbuch hinzufügen", click: () => fenster.webContents.session.addWordToSpellCheckerDictionary(p.misspelledWord) });
      vorlage.push({ type: "separator" });
    }
    if (p.mediaType === "image") {
      vorlage.push({ label: "Bild kopieren", click: () => fenster.webContents.copyImageAt(p.x, p.y) }, { type: "separator" });
    }
    if (p.isEditable || p.selectionText) {
      vorlage.push(
        { label: "Ausschneiden", role: "cut", enabled: p.editFlags.canCut },
        { label: "Kopieren", role: "copy", enabled: p.editFlags.canCopy },
        { label: "Einfügen", role: "paste", enabled: p.editFlags.canPaste },
        { label: "Als reinen Text einfügen", role: "pasteAndMatchStyle", enabled: p.editFlags.canPaste },
        { type: "separator" },
        { label: "Alles auswählen", role: "selectAll" },
      );
    }
    if (vorlage.length) Menu.buildFromTemplate(vorlage).popup({ window: fenster });
  });

  // Zoom mit Strg + Mausrad / Strg +/-/0
  fenster.webContents.on("before-input-event", (e, input) => {
    if (input.type !== "keyDown" || !input.control) return;
    const wc = fenster.webContents;
    if (input.key === "+" || input.key === "=") { wc.setZoomLevel(wc.getZoomLevel() + 0.5); e.preventDefault(); }
    if (input.key === "-") { wc.setZoomLevel(wc.getZoomLevel() - 0.5); e.preventDefault(); }
    if (input.key === "0") { wc.setZoomLevel(0); e.preventDefault(); }
    if (input.key.toLowerCase() === "r" && input.shift) { wc.reload(); e.preventDefault(); }
    if (input.key === "F12" || (input.shift && input.key.toLowerCase() === "i")) { wc.toggleDevTools(); e.preventDefault(); }
  });
}

app.on("second-instance", () => { if (fenster) { if (fenster.isMinimized()) fenster.restore(); fenster.focus(); } });
app.whenReady().then(() => {
  try {
    const verfuegbar = session.defaultSession.availableSpellCheckerLanguages || [];
    const wunsch = ["de-AT", "de-DE", "de", "en-US", "en-GB"].filter(l => verfuegbar.includes(l));
    if (wunsch.length) session.defaultSession.setSpellCheckerLanguages(wunsch);
  } catch { /* Windows nutzt die eingebaute Rechtschreibprüfung */ }
  fensterErstellen();
});
app.on("window-all-closed", () => app.quit());

/* ---------------- Entwurf (automatisches Speichern) ---------------- */
ipcMain.handle("entwurf:laden", () => lesen(DATEI.entwurf));
ipcMain.handle("entwurf:speichern", (_e, d) => { schreiben(DATEI.entwurf, d); return true; });
ipcMain.on("entwurf:speichern-sofort", (e, d) => { try { schreiben(DATEI.entwurf, d); e.returnValue = true; } catch { e.returnValue = false; } });
ipcMain.handle("letzte:laden", () => lesen(DATEI.letzte));
ipcMain.handle("letzte:speichern", (_e, d) => { schreiben(DATEI.letzte, d); return true; });
ipcMain.handle("einstellungen:laden", () => ({ ...einstellungen(), eingang: eingangFinden() }));
ipcMain.handle("einstellungen:setzen", (_e, k, v) => { einstellungSetzen(k, v); return true; });

/* ---------------- Export ---------------- */
ipcMain.handle("eingang:waehlen", async () => eingangWaehlen(fenster));
ipcMain.handle("eingang:oeffnen", async () => { const e = eingangFinden(); if (e) shell.openPath(e); return e; });
ipcMain.handle("export:speichern", async (_e, name, html) => {
  let ordner = eingangFinden();
  if (!ordner) ordner = await eingangWaehlen(fenster);
  if (!ordner) return { ok: false, grund: "abgebrochen" };
  const sauber = name.replace(/[\\/:*?"<>|]+/g, "");
  const ziel = path.join(ordner, sauber);
  schreiben(ziel, html);
  return { ok: true, pfad: ziel, ordner };
});

/* Handschrift-Ausschnitte: Seite unsichtbar rendern und Bereiche fotografieren */
ipcMain.handle("export:ausschnitte", async (_e, { html, breite, hoehe, bereiche }) => {
  const SKALA = 2;
  const tmp = path.join(os.tmpdir(), `notizblock-seite-${process.pid}-${Date.now()}.html`);
  fs.writeFileSync(tmp, html, "utf8");
  // doppelt so groß rendern (Zoom 2) → scharfe Ausschnitte für Word
  const h = Math.min(Math.ceil(hoehe) + 40, 7800);
  const w = new BrowserWindow({
    show: false, width: Math.ceil(breite * SKALA), height: h * SKALA, useContentSize: true, enableLargerThanScreen: true,
    webPreferences: { offscreen: true, sandbox: true, contextIsolation: true, javascript: false, zoomFactor: SKALA },
  });
  try {
    w.webContents.setFrameRate(5);
    await w.loadFile(tmp);
    w.webContents.setZoomFactor(SKALA);
    w.setContentSize(Math.ceil(breite * SKALA), h * SKALA);
    await new Promise(r => setTimeout(r, 800));
    const ergebnis = [];
    for (const b of bereiche) {
      const y = Math.max(0, Math.floor(b.y)), hh = Math.ceil(Math.min(b.h, h - y));
      if (hh <= 0) { ergebnis.push(null); continue; }
      const rect = { x: Math.max(0, Math.floor(b.x)) * SKALA, y: y * SKALA, width: Math.ceil(b.w) * SKALA, height: hh * SKALA };
      const bild = await w.webContents.capturePage(rect);
      ergebnis.push(bild.isEmpty() ? null : bild.toDataURL());
    }
    return ergebnis;
  } finally {
    w.destroy();
    fs.rm(tmp, { force: true }, () => {});
  }
});

/* ---------------- Bilder aus dem Internet ---------------- */
async function holen(url, maxBytes = 20 * 1024 * 1024) {
  if (!/^https?:\/\//i.test(url)) throw new Error("Nur http(s)-Adressen");
  const r = await net.fetch(url, { headers: { "User-Agent": USER_AGENT } });
  if (!r.ok) throw new Error(`Fehler ${r.status}`);
  const typ = (r.headers.get("content-type") || "").split(";")[0].trim();
  const buf = Buffer.from(await r.arrayBuffer());
  if (buf.length > maxBytes) throw new Error("Bild ist zu groß");
  return { typ, buf };
}
ipcMain.handle("bild:laden", async (_e, url) => {
  try {
    const { typ, buf } = await holen(url);
    let mime = typ;
    if (!mime.startsWith("image/")) {
      const n = nativeImage.createFromBuffer(buf);
      if (n.isEmpty()) return { ok: false, grund: "Das ist kein Bild." };
      return { ok: true, src: n.toDataURL() };
    }
    if (/svg/.test(mime)) { // SVG → PNG
      const n = nativeImage.createFromBuffer(buf);
      if (!n.isEmpty()) return { ok: true, src: n.toDataURL() };
    }
    return { ok: true, src: `data:${mime};base64,${buf.toString("base64")}` };
  } catch (e) { return { ok: false, grund: e.message }; }
});

ipcMain.handle("bild:suchen", async (_e, begriff, seite = 0) => {
  const url = "https://commons.wikimedia.org/w/api.php?" + new URLSearchParams({
    action: "query", format: "json", generator: "search", gsrsearch: `${begriff} filetype:bitmap|drawing`, gsrnamespace: "6",
    gsrlimit: "30", gsroffset: String(seite * 30), prop: "imageinfo", iiprop: "url|mime|extmetadata|size", iiurlwidth: "900",
    iiextmetadatafilter: "Artist|LicenseShortName|ImageDescription", uselang: "de",
  });
  try {
    const { buf } = await holen(url, 5 * 1024 * 1024);
    const d = JSON.parse(buf.toString("utf8"));
    const seiten = Object.values((d.query && d.query.pages) || {}).sort((a, b) => (a.index || 0) - (b.index || 0));
    const text = s => String(s || "").replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();
    return {
      ok: true,
      treffer: seiten.map(p => {
        const i = (p.imageinfo || [])[0] || {};
        const m = i.extmetadata || {};
        return {
          titel: String(p.title || "").replace(/^Datei:|^File:/, "").replace(/\.[a-z0-9]+$/i, ""),
          vorschau: i.thumburl, bild: i.thumburl || i.url, seite: i.descriptionurl,
          urheber: text(m.Artist && m.Artist.value).slice(0, 80), lizenz: text(m.LicenseShortName && m.LicenseShortName.value),
        };
      }).filter(t => t.vorschau),
    };
  } catch (e) { return { ok: false, grund: e.message }; }
});

ipcMain.handle("link:oeffnen", (_e, url) => { if (/^https?:/.test(url)) shell.openExternal(url); });
ipcMain.handle("zwischenablage:text", (_e, t) => { clipboard.writeText(String(t)); return true; });

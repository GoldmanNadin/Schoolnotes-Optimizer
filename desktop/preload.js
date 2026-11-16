"use strict";
const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("notizblock", {
  desktop: true,
  entwurfLaden: () => ipcRenderer.invoke("entwurf:laden"),
  entwurfSpeichern: d => ipcRenderer.invoke("entwurf:speichern", d),
  entwurfSofortSpeichern: d => ipcRenderer.sendSync("entwurf:speichern-sofort", d),
  letzteLaden: () => ipcRenderer.invoke("letzte:laden"),
  letzteSpeichern: d => ipcRenderer.invoke("letzte:speichern", d),
  einstellungen: () => ipcRenderer.invoke("einstellungen:laden"),
  einstellungSetzen: (k, v) => ipcRenderer.invoke("einstellungen:setzen", k, v),
  eingangWaehlen: () => ipcRenderer.invoke("eingang:waehlen"),
  eingangOeffnen: () => ipcRenderer.invoke("eingang:oeffnen"),
  exportSpeichern: (name, html) => ipcRenderer.invoke("export:speichern", name, html),
  ausschnitte: daten => ipcRenderer.invoke("export:ausschnitte", daten),
  bildLaden: url => ipcRenderer.invoke("bild:laden", url),
  bildSuchen: (begriff, seite) => ipcRenderer.invoke("bild:suchen", begriff, seite),
  linkOeffnen: url => ipcRenderer.invoke("link:oeffnen", url),
  textKopieren: t => ipcRenderer.invoke("zwischenablage:text", t),
});

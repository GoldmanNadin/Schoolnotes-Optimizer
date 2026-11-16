#!/usr/bin/env node
/*
 * Wandelt eine sauber nachgezeichnete Skizze (SVG) in ein scharfes PNG für Word um.
 * Nutzt den eingebauten Browser von Windows (Edge) oder Chrome – nichts zu installieren.
 *
 *   node svg_zu_png.js <zeichnung.svg> <zeichnung.png>
 */
"use strict";
const fs = require("fs");
const os = require("os");
const path = require("path");
const { spawnSync } = require("child_process");

const [ein, aus] = process.argv.slice(2);
if (!ein || !aus) { console.error("Aufruf: node svg_zu_png.js <zeichnung.svg> <zeichnung.png>"); process.exit(1); }
const svg = fs.readFileSync(ein, "utf8");

// Größe aus viewBox bzw. width/height lesen
const vb = svg.match(/viewBox\s*=\s*"([\d.\s-]+)"/i);
let w = 800, h = 500;
if (vb) { const z = vb[1].trim().split(/\s+/).map(Number); w = z[2]; h = z[3]; }
const mw = svg.match(/<svg[^>]*\swidth\s*=\s*"([\d.]+)/i), mh = svg.match(/<svg[^>]*\sheight\s*=\s*"([\d.]+)/i);
if (mw) w = Number(mw[1]); if (mh) h = Number(mh[1]);
w = Math.ceil(w); h = Math.ceil(h);

const browser = [
  process.env["ProgramFiles(x86)"] && path.join(process.env["ProgramFiles(x86)"], "Microsoft", "Edge", "Application", "msedge.exe"),
  process.env.ProgramFiles && path.join(process.env.ProgramFiles, "Microsoft", "Edge", "Application", "msedge.exe"),
  process.env.ProgramFiles && path.join(process.env.ProgramFiles, "Google", "Chrome", "Application", "chrome.exe"),
  process.env.LOCALAPPDATA && path.join(process.env.LOCALAPPDATA, "Google", "Chrome", "Application", "chrome.exe"),
].find(p => p && fs.existsSync(p));
if (!browser) { console.error("FEHLER: Kein Edge/Chrome gefunden."); process.exit(2); }

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "svg-png-"));
const html = path.join(tmp, "seite.html");
fs.writeFileSync(html, `<!doctype html><meta charset="utf-8"><style>html,body{margin:0;background:#fff;overflow:hidden}svg{display:block;width:${w}px;height:${h}px}</style>${svg.replace(/<\?xml[^>]*>/, "")}`, "utf8");
const ziel = path.resolve(aus);
const r = spawnSync(browser, [
  "--headless=new", "--disable-gpu", "--hide-scrollbars", "--no-first-run", "--no-default-browser-check",
  `--user-data-dir=${path.join(tmp, "profil")}`, "--force-device-scale-factor=2",
  `--window-size=${w},${h}`, `--screenshot=${ziel}`, "file:///" + html.replace(/\\/g, "/"),
], { timeout: 60000, encoding: "utf8" });
fs.rmSync(tmp, { recursive: true, force: true });
if (!fs.existsSync(ziel)) { console.error("FEHLER: Bild wurde nicht erstellt.", r.stderr || ""); process.exit(3); }
console.log(JSON.stringify({ ok: true, datei: ziel, breite_px: w * 2, hoehe_px: h * 2 }));

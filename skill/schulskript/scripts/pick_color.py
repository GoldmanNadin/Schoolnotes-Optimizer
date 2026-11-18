#!/usr/bin/env python3
"""Waehlt zufaellig eine noch freie Fachfarbe.

Aufruf:
  pick_color.py --fach "Mathematik"                # nutzt faecher.json als Register
  pick_color.py --fach "Mathematik" --used 7B4FA3 B8344A   # zusaetzlich belegte Farben (Hex oder Name)
  pick_color.py --list                             # Palette + Belegung zeigen

Gibt JSON aus: {"fach":..., "name":..., "hex":..., "neu":true/false}
Hat das Fach schon eine Farbe im Register, wird diese wiederverwendet (neu=false).
"""
import argparse, json, random, sys, pathlib

HERE = pathlib.Path(__file__).resolve().parent.parent
REG = HERE / "faecher.json"

# Kraeftige, gut lesbare Farben (weisser Text auf Farbe ok), bewusst weit auseinander.
PALETTE = {
    "Gruen": "1F7A5C", "Blau": "0070C0", "Violett": "7B4FA3", "Bordeaux": "B8344A",
    "Orange": "D9731A", "Tuerkis": "00838F", "Pink": "C2478F", "Indigo": "3F51B5",
    "Braun": "8D5A3B", "Oliv": "6B8E23", "Schiefer": "546E7A", "Gold": "B8860B",
}

def norm(x):
    x = x.strip().lstrip("#").upper()
    for n, h in PALETTE.items():
        if x == n.upper(): return h
    return x

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--fach")
    ap.add_argument("--used", nargs="*", default=[])
    ap.add_argument("--registry", default=str(REG))
    ap.add_argument("--list", action="store_true")
    a = ap.parse_args()
    reg = json.load(open(a.registry, encoding="utf8")).get("faecher", {})
    used = {norm(v["hex"]) for v in reg.values()} | {norm(u) for u in a.used}
    if a.list:
        for n, h in PALETTE.items():
            print(f"{n:9} #{h}  {'belegt' if h in used else 'frei'}")
        return
    if not a.fach: sys.exit("--fach fehlt")
    for k, v in reg.items():
        if k.lower() == a.fach.lower():
            print(json.dumps({"fach": k, **v, "neu": False}, ensure_ascii=False)); return
    free = [(n, h) for n, h in PALETTE.items() if h not in used]
    if not free:  # alle vergeben -> eine zufaellige Farbe, dem Nutzer sagen
        free = list(PALETTE.items())
    n, h = random.choice(free)
    print(json.dumps({"fach": a.fach, "name": n, "hex": h, "neu": True}, ensure_ascii=False))

if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Lægger annoncer og plantegningsanalyser sammen og bygger dashboardet.

Læser  data/bolig/annoncer.json  (fra hent.py)
       docs/bolig/analyser.json  (fra analyse.py)
Skriver docs/bolig.html  (GitHub Pages; plantegningerne ligger i docs/bolig/plan/)
        dist/bolig.html  (til udgivelse som artifact, med docs/bolig/plan/ som filer)
"""

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
ANNONCER = ROOT / "data" / "bolig" / "annoncer.json"
ANALYSER = ROOT / "docs" / "bolig" / "analyser.json"
TEMPLATE = Path(__file__).with_name("template.html")
UD = [ROOT / "docs" / "bolig.html", ROOT / "dist" / "bolig.html"]

# Felter, siden ikke bruger — holdes ude for at gøre siden lettere.
UDELAD = {"beskrivelse", "plan_filer", "foto", "lat", "lon"}


def main() -> None:
    a = json.loads(ANNONCER.read_text())
    analyser = json.loads(ANALYSER.read_text()) if ANALYSER.exists() else {}

    boliger = []
    for b in a["boliger"]:
        an = analyser.get(b["id"])
        if an and an.get("_plan") != "|".join(b["plan_filer"]):
            an = None  # plantegningen er skiftet siden analysen
        if an:
            dom = an["dom"]
            an = {k: v for k, v in an.items() if not k.startswith("_")}
        else:
            dom = "uvist"
        ud = {k: v for k, v in b.items() if k not in UDELAD}
        ud["dom"] = dom
        ud["analyse"] = an
        boliger.append(ud)

    kilder = {
        "i_alt_landet": a.get("i_alt_landet"),
        "plan_boligsiden": sum(b["plan_kilde"] == "boligsiden" for b in a["boliger"]),
        "plan_maegler": sum(b["plan_kilde"] == "mægler" for b in a["boliger"]),
        "plan_mangler": sum(not b["plan_filer"] for b in a["boliger"]),
    }
    data = {"hentet": a["hentet"], "kriterier": a["kriterier"], "kilder": kilder, "boliger": boliger}
    side = TEMPLATE.read_text().replace(
        "/*__DATA__*/null", json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    )
    for sti in UD:
        sti.parent.mkdir(parents=True, exist_ok=True)
        sti.write_text(side)
        print(f"→ {sti.relative_to(ROOT)} ({len(side) // 1024} kB)")

    domme = {}
    for b in boliger:
        domme[b["dom"]] = domme.get(b["dom"], 0) + 1
    print("Domme:", domme)


if __name__ == "__main__":
    main()

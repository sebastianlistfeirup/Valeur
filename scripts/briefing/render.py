"""Sæt docs/briefing.json ind i skabelonen og skriv:
  docs/briefing.html   siden på GitHub Pages
  dist/briefing.html   samme side uden <html>-skelet (til udgivelse som artifact)
  docs/py/*.py         byggescripts, som opdateringssiden (docs/opdater.html) kører i browseren
  docs/py/roller.json  tjeklistens tekster (fra kilder.ROLLER)
"""
import json, shutil, sys
from pathlib import Path

here = Path(__file__).resolve().parent
repo = here.parents[1]
sys.path.insert(0, str(here))
import kilder

html = (here / "template.html").read_text(encoding="utf-8").replace(
    "/*DATA*/null", (repo / "docs" / "briefing.json").read_text(encoding="utf-8"))

hoved = ('<!doctype html>\n<html lang="da">\n<head>\n<meta charset="utf-8">\n'
         '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
         '</head>\n<body>\n')
(repo / "docs" / "briefing.html").write_text(hoved + html + "\n</body>\n</html>\n", encoding="utf-8")
(repo / "dist").mkdir(exist_ok=True)
(repo / "dist" / "briefing.html").write_text(html, encoding="utf-8")

py = repo / "docs" / "py"
py.mkdir(exist_ok=True)
for navn in ("kilder.py", "build_data.py", "dybde.py", "medlemmer.py"):
    shutil.copyfile(here / navn, py / navn)
(py / "roller.json").write_text(json.dumps(kilder.ROLLER, ensure_ascii=False, indent=1), encoding="utf-8")
print("skrev docs/briefing.html, dist/briefing.html og docs/py/")

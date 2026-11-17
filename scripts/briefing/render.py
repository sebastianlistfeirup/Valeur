"""Sæt docs/briefing.json ind i skabelonen og skriv docs/briefing.html (GitHub Pages) og dist/briefing.html (artifact)."""
from pathlib import Path

here = Path(__file__).resolve().parent
repo = here.parents[1]
html = (here / "template.html").read_text(encoding="utf-8").replace(
    "/*DATA*/null", (repo / "docs" / "briefing.json").read_text(encoding="utf-8"))

hoved = ('<!doctype html>\n<html lang="da">\n<head>\n<meta charset="utf-8">\n'
         '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
         '</head>\n<body>\n')
(repo / "docs" / "briefing.html").write_text(hoved + html + "\n</body>\n</html>\n", encoding="utf-8")
(repo / "dist").mkdir(exist_ok=True)
(repo / "dist" / "briefing.html").write_text(html, encoding="utf-8")
print("skrev docs/briefing.html og dist/briefing.html")

#!/usr/bin/env python3
"""Henter boligmarkedet og plantegningerne.

Kilde 1  Boligsiden (api.boligsiden.dk/search/cases) — samler sagerne fra alle
         mæglere: home, Nybolig, danbolig, EDC, LokalBolig, Estate m.fl.
Kilde 2  Mæglernes egne sagssider, når Boligsiden ikke har plantegningen med:
         Nybolig og Estate (mindworking), danbolig og LokalBolig.

Kriterier (ændres i KRITERIER nedenfor):
  - pris 6.000.000–8.000.000 kr.
  - 3 værelser
  - postnummer 1000–2300 (København K, V, Frederiksberg, Ø, Nordhavn, N, S)
  - mindst 70 m², eller under 70 m² hvis der er disponible ekstra rum
    (kælder, vægtet areal, hems/loftsrum nævnt i teksten)

Skriver data/bolig/annoncer.json og lægger plantegningerne i
data/bolig/plan/<caseID>_<n>.png (fuld opløsning, til analysen) og
docs/bolig/plan/<caseID>.webp (lille udgave, til dashboardet).
"""

import html
import io
import json
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

import requests
from PIL import Image

ROOT = Path(__file__).resolve().parents[2]
DATA = ROOT / "data" / "bolig"
PLAN_FULD = DATA / "plan"
PLAN_LILLE = ROOT / "docs" / "bolig" / "plan"
OUT = DATA / "annoncer.json"

KRITERIER = {
    "pris_min": 6_000_000,
    "pris_max": 8_000_000,
    "vaerelser": 3,
    "postnr_min": 1000,
    "postnr_max": 2300,
    "m2_min": 70,
    "m2_min_med_ekstra_rum": 60,
    "typer": {"condo", "villa apartment", "terraced house", "villa"},
}

HEADERS = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/130.0 Safari/537.36",
    "Accept-Language": "da-DK,da;q=0.9",
}
API_HEADERS = {"User-Agent": "curl/8.5.0", "Accept": "application/json"}

TYPER = {
    "condo": "Ejerlejlighed",
    "villa apartment": "Villalejlighed",
    "terraced house": "Rækkehus",
    "villa": "Villa",
}

EKSTRA_RUM = re.compile(
    r"\b(hems|kælderrum|kælder|loftsrum|pulterrum|udnyttet tag|tagetage|"
    r"ekstra rum|arbejdsværelse|kontor|alkove|walk-in|depotrum)\b",
    re.I,
)

S = requests.Session()
S.headers.update(HEADERS)


def omraade(zipcode: int) -> str:
    if 1400 <= zipcode <= 1449:
        return "Christianshavn"
    if zipcode < 1500:
        return "Indre By"
    if zipcode < 1800:
        return "Vesterbro"
    if zipcode < 2000:
        return "Frederiksberg C"
    if zipcode == 2000:
        return "Frederiksberg"
    if zipcode == 2100:
        return "Østerbro"
    if zipcode == 2150:
        return "Nordhavn"
    if zipcode == 2200:
        return "Nørrebro"
    if zipcode == 2300:
        return "Amager / København S"
    return str(zipcode)


# ── Boligsiden ──────────────────────────────────────────────────────────────

def hent_sager() -> list[dict]:
    k = KRITERIER
    url = (
        "https://api.boligsiden.dk/search/cases?per_page=100"
        f"&priceMin={k['pris_min']}&priceMax={k['pris_max']}"
        f"&numberOfRoomsMin={k['vaerelser']}&numberOfRoomsMax={k['vaerelser']}"
        "&page={side}"
    )
    sager = []
    for side in range(1, 30):
        r = requests.get(url.format(side=side), headers=API_HEADERS, timeout=60)
        r.raise_for_status()
        d = r.json()
        sager += d["cases"]
        if len(d["cases"]) < 100:
            break
    print(f"Boligsiden: {len(sager)} sager i hele landet ({d['totalHits']} ifølge API)")
    return sager


def maegler_link(c: dict) -> str | None:
    pid = c.get("providerCaseID", "")
    zipcode = c["address"]["zipCode"]
    if re.fullmatch(r"\d{17}nyb", pid):
        shop = pid[11:17]
        return f"https://www.nybolig.dk/services/redirect/case/N{shop}/N{shop}{pid[4:11]}"
    if re.fullmatch(r"\d{17}est", pid):
        shop = pid[11:17]
        return f"https://www.estate.dk/services/redirect/case/N{shop}/N{shop}{pid[4:11]}"
    if re.fullmatch(r"\d{13}dan", pid):
        return f"https://danbolig.dk/bolig/x/{zipcode}/lejlighed/{pid[:10]}-{pid[10:13]}/"
    if re.fullmatch(r"\d{13}hom", pid):
        return f"https://home.dk/sag/{pid[:10]}"
    m = re.match(r"(\d+-X\d{7})\d*lok$", pid)
    if m:
        return f"https://www.lokalbolig.dk/bolig/ejerlejlighed/x/{zipcode}/x/{m.group(1)}"
    oh = (c.get("nextOpenHouse") or {}).get("signUpUrl")
    if oh and "/sag/" in oh:
        return oh
    return None


def boligsiden_link(c: dict) -> str:
    a = c["address"]
    vej = (a.get("road") or {}).get("slug")
    if not vej:
        return f"https://www.boligsiden.dk/postnummer/{a['zipCode']}/tilsalg"
    return f"https://www.boligsiden.dk/postnummer/{a['zipCode']}/vej/{vej}/tilsalg"


def billede(sources: list[dict], bredde: int = 600) -> str | None:
    if not sources:
        return None
    best = min(sources, key=lambda s: abs(s["size"]["width"] - bredde))
    return best["url"]


# ── Plantegninger ───────────────────────────────────────────────────────────

def plan_urls_boligsiden(c: dict) -> list[str]:
    urls = []
    for im in c.get("floorPlanImages") or []:
        src = max(im["imageSources"], key=lambda s: s["size"]["width"])
        urls.append(src["url"])
    if not urls and c.get("caseUrlFloorPlan"):
        urls.append(c["caseUrlFloorPlan"])
    return urls


def plan_urls_maegler(c: dict, link: str) -> list[str]:
    """Slår plantegningen op på mæglerens egen sagsside."""
    try:
        r = S.get(link, timeout=45)
    except requests.RequestException:
        return []
    if r.status_code != 200:
        return []
    t = r.text
    pid = c.get("providerCaseID", "")
    if pid.endswith(("nyb", "est")):
        # <img class="floorplan__drawing lazy" src="https://nybolig.mindworking.eu/...">
        return [
            html.unescape(u)
            for u in re.findall(r'class="floorplan__drawing[^"]*"\s+src="([^"]+)"', t)
        ]
    if pid.endswith("dan"):
        m = re.search(r':floorplan2d="([^"]+)"', t)
        if not m:
            return []
        try:
            items = json.loads(html.unescape(m.group(1)))
        except json.JSONDecodeError:
            return []
        return [i["url"] for i in items if i.get("url")]
    if pid.endswith("lok"):
        return list(dict.fromkeys(re.findall(r'"layoutImage":"([^"]+)"', t)))
    return []


def hent_billede(url: str) -> Image.Image | None:
    try:
        r = S.get(url, timeout=60)
    except requests.RequestException:
        return None
    if r.status_code != 200 or not r.content:
        return None
    ctype = r.headers.get("content-type", "")
    data = r.content
    if "svg" in ctype or data.lstrip()[:5] in (b"<?xml", b"<svg "):
        import cairosvg

        data = cairosvg.svg2png(bytestring=data, output_width=1600, background_color="white")
    try:
        im = Image.open(io.BytesIO(data))
        im.load()
    except Exception:
        return None
    if im.mode in ("RGBA", "LA", "P"):
        bg = Image.new("RGB", im.size, "white")
        im = im.convert("RGBA")
        bg.paste(im, mask=im.split()[-1])
        im = bg
    return im.convert("RGB")


def gem_plan(case_id: str, urls: list[str]) -> list[str]:
    """Gemmer op til tre plantegninger pr. sag. Returnerer filnavnene."""
    filer = []
    for n, url in enumerate(urls[:3]):
        fuld = PLAN_FULD / f"{case_id}_{n}.png"
        if not fuld.exists():
            im = hent_billede(url)
            if im is None:
                continue
            im.thumbnail((1600, 1600))
            im.save(fuld, optimize=True)
            time.sleep(0.2)
        filer.append(fuld.name)
    if filer:
        lille = PLAN_LILLE / f"{case_id}.webp"
        if not lille.exists():
            im = Image.open(PLAN_FULD / filer[0])
            im.thumbnail((900, 900))
            im.save(lille, "WEBP", quality=72, method=6)
    return filer


# ── Normalisering ───────────────────────────────────────────────────────────

def normaliser(c: dict) -> dict:
    a = c["address"]
    tekst = (c.get("descriptionTitle") or "") + "\n" + (c.get("descriptionBody") or "")
    ekstra = sorted({m.lower() for m in EKSTRA_RUM.findall(tekst)})
    etage = a.get("floor")
    dør = a.get("door")
    adresse = f"{a['roadName']} {a['houseNumber']}"
    if etage:
        adresse += f", {'st.' if etage in ('st', '0') else etage + '.'}"
        if dør:
            adresse += f" {dør}."
    link = maegler_link(c)
    return {
        "id": c["caseID"],
        "adresse": adresse,
        "postnr": a["zipCode"],
        "by": a["cityName"],
        "omraade": omraade(a["zipCode"]),
        "lat": c["coordinates"]["lat"],
        "lon": c["coordinates"]["lon"],
        "pris": c["priceCash"],
        "m2": c["housingArea"],
        "vaegtet_m2": c.get("weightedArea"),
        "kaelder_m2": c.get("basementArea"),
        "kvmpris": c.get("perAreaPrice"),
        "vaerelser": c["numberOfRooms"],
        "type": TYPER.get(c["addressType"], c["addressType"]),
        "etage": etage,
        "byggeaar": c.get("yearBuilt"),
        "energi": (c.get("energyLabel") or "").upper() or None,
        "ejerudgift": c.get("monthlyExpense"),
        "altan": c.get("hasBalcony"),
        "terrasse": c.get("hasTerrace"),
        "elevator": c.get("hasElevator"),
        "dage": c.get("daysOnMarket"),
        "prisaendring": c.get("priceChangePercentage"),
        "maegler": c["realtor"]["name"],
        "maegler_link": link,
        "boligsiden_link": boligsiden_link(c),
        "fremvisning": (c.get("nextOpenHouse") or {}).get("date"),
        "titel": c.get("descriptionTitle"),
        "beskrivelse": (c.get("descriptionBody") or "")[:1500],
        "foto": billede((c.get("defaultImage") or {}).get("imageSources", [])),
        "ekstra_rum_i_tekst": ekstra,
        "_plan_urls": plan_urls_boligsiden(c),
        "_maegler_link": link,
        "_pid": c.get("providerCaseID", ""),
    }


def opfylder(b: dict) -> tuple[bool, str | None]:
    k = KRITERIER
    if not (k["postnr_min"] <= b["postnr"] <= k["postnr_max"]):
        return False, None
    if b["m2"] >= k["m2_min"]:
        return True, None
    ekstra = bool(b["ekstra_rum_i_tekst"] or b["kaelder_m2"] or
                  (b["vaegtet_m2"] and b["vaegtet_m2"] > b["m2"]))
    if b["m2"] >= k["m2_min_med_ekstra_rum"] and ekstra:
        return True, "Under 70 m², men med disponible ekstra rum — tjek plantegningen"
    return False, None


def main() -> None:
    PLAN_FULD.mkdir(parents=True, exist_ok=True)
    PLAN_LILLE.mkdir(parents=True, exist_ok=True)

    sager = [c for c in hent_sager() if c["addressType"] in KRITERIER["typer"]]
    boliger, fravalgt_m2 = [], 0
    for c in sager:
        b = normaliser(c)
        ok, note = opfylder(b)
        if not ok:
            if KRITERIER["postnr_min"] <= b["postnr"] <= KRITERIER["postnr_max"]:
                fravalgt_m2 += 1
            continue
        b["areal_note"] = note
        boliger.append(b)
    print(f"I postnr. {KRITERIER['postnr_min']}–{KRITERIER['postnr_max']}: "
          f"{len(boliger)} boliger ({fravalgt_m2} fravalgt på m²)")

    kilder = {"boligsiden": 0, "mægler": 0, "mangler": 0}
    for i, b in enumerate(boliger, 1):
        urls, kilde = b.pop("_plan_urls"), "boligsiden"
        link, pid = b.pop("_maegler_link"), b.pop("_pid")
        if not urls and link:
            urls, kilde = plan_urls_maegler({"providerCaseID": pid}, link), "mægler"
        filer = gem_plan(b["id"], urls) if urls else []
        b["plan_filer"] = filer
        b["plan_kilde"] = kilde if filer else None
        b["plan_thumb"] = f"bolig/plan/{b['id']}.webp" if filer else None
        kilder[kilde if filer else "mangler"] += 1
        if i % 25 == 0:
            print(f"  plantegninger: {i}/{len(boliger)}", flush=True)

    print("Plantegninger:", kilder)
    OUT.write_text(json.dumps({
        "hentet": datetime.now(timezone.utc).isoformat(timespec="minutes"),
        "kriterier": {k: v for k, v in KRITERIER.items() if k != "typer"},
        "i_alt_landet": len(sager),
        "boliger": boliger,
    }, ensure_ascii=False, indent=1))
    print(f"→ {OUT.relative_to(ROOT)}")


if __name__ == "__main__":
    sys.exit(main())

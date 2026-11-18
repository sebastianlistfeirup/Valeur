#!/usr/bin/env python3
"""Læser plantegningerne og afgør, om boligen har to reelle soveværelser.

Reglerne står i plan_prompt.md. Resultatet gemmes i docs/bolig/analyser.json,
nøglet på sagens id og plantegningernes filnavne, så en bolig kun analyseres
igen, hvis den er ny eller har fået en ny plantegning.

Tre måder at køre det på:

  python scripts/bolig/analyse.py
      Analyserer nye plantegninger med Claude (kræver ANTHROPIC_API_KEY eller
      `ant auth login`).

  python scripts/bolig/analyse.py --batches <mappe> [--str 24]
      Skriver nye, uanalyserede boliger ud i batchfiler, så de kan læses af en
      agent eller et menneske.

  python scripts/bolig/analyse.py --indlaes <mappe>
      Lægger færdige svar (<mappe>/svar_*.json, en liste af {"id": ..., ...})
      ind i analyser.json.
"""

import argparse
import base64
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
ANNONCER = ROOT / "data" / "bolig" / "annoncer.json"
PLAN = ROOT / "data" / "bolig" / "plan"
CACHE = ROOT / "docs" / "bolig" / "analyser.json"
PROMPT = Path(__file__).with_name("plan_prompt.md")

MODEL = "claude-opus-5-5"
DOMME = {"ja", "mulighed", "nej", "uvist"}


def laes_cache() -> dict:
    return json.loads(CACHE.read_text()) if CACHE.exists() else {}


def gem_cache(cache: dict) -> None:
    CACHE.parent.mkdir(parents=True, exist_ok=True)
    CACHE.write_text(json.dumps(dict(sorted(cache.items())), ensure_ascii=False, indent=1))


def noegle(b: dict) -> str:
    return "|".join(b["plan_filer"])


def mangler(boliger: list[dict], cache: dict) -> list[dict]:
    return [b for b in boliger
            if b["plan_filer"] and cache.get(b["id"], {}).get("_plan") != noegle(b)]


def kontekst(b: dict) -> str:
    linjer = [
        f"Adresse: {b['adresse']}, {b['postnr']} {b['by']}",
        f"Oplyst: {b['vaerelser']} værelser, {b['m2']} m² bolig"
        + (f", vægtet {b['vaegtet_m2']} m²" if b.get("vaegtet_m2") else "")
        + (f", kælder {b['kaelder_m2']} m²" if b.get("kaelder_m2") else ""),
        f"Overskrift: {b.get('titel') or ''}",
        f"Annoncetekst (uddrag): {(b.get('beskrivelse') or '')[:900]}",
    ]
    return "\n".join(linjer)


def valider(svar: dict) -> dict:
    if svar.get("dom") not in DOMME:
        raise ValueError(f"ukendt dom: {svar.get('dom')!r}")
    for felt in ("sovevaerelser_i_dag", "sovevaerelser_mulige"):
        svar[felt] = int(svar.get(felt) or 0)
    for felt in ("indgreb", "ekstra_rum", "opholdsflade", "begrundelse"):
        svar[felt] = (svar.get(felt) or "").strip()
    svar.setdefault("rum", [])
    svar.setdefault("sikkerhed", "middel")
    return svar


def analyser_med_claude(boliger: list[dict], cache: dict) -> None:
    import anthropic

    client = anthropic.Anthropic()
    system = PROMPT.read_text()
    for i, b in enumerate(boliger, 1):
        indhold = []
        for fil in b["plan_filer"]:
            data = base64.standard_b64encode((PLAN / fil).read_bytes()).decode()
            indhold.append({"type": "image",
                            "source": {"type": "base64", "media_type": "image/png", "data": data}})
        indhold.append({"type": "text", "text": kontekst(b) + "\n\nVurdér plantegningen."})
        try:
            with client.messages.stream(
                model=MODEL,
                max_tokens=16000,
                system=system,
                thinking={"type": "adaptive"},
                output_config={"effort": "high"},
                messages=[{"role": "user", "content": indhold}],
            ) as stream:
                msg = stream.get_final_message()
        except anthropic.APIStatusError as e:
            print(f"  {b['adresse']}: API-fejl {e.status_code}", file=sys.stderr)
            continue
        except anthropic.APIConnectionError:
            print(f"  {b['adresse']}: ingen forbindelse", file=sys.stderr)
            continue
        if msg.stop_reason == "refusal":
            print(f"  {b['adresse']}: afvist", file=sys.stderr)
            continue
        tekst = "".join(bl.text for bl in msg.content if bl.type == "text")
        try:
            svar = valider(json.loads(tekst[tekst.index("{"): tekst.rindex("}") + 1]))
        except (ValueError, json.JSONDecodeError) as e:
            print(f"  {b['adresse']}: kunne ikke læse svaret ({e})", file=sys.stderr)
            continue
        svar["_plan"] = noegle(b)
        cache[b["id"]] = svar
        gem_cache(cache)
        print(f"  {i}/{len(boliger)} {b['adresse']}: {svar['dom']}", flush=True)


def skriv_batches(boliger: list[dict], mappe: Path, stoerrelse: int) -> None:
    mappe.mkdir(parents=True, exist_ok=True)
    for n in range(0, len(boliger), stoerrelse):
        del_ = [{
            "id": b["id"],
            "kontekst": kontekst(b),
            "plantegninger": [str(PLAN / f) for f in b["plan_filer"]],
        } for b in boliger[n:n + stoerrelse]]
        sti = mappe / f"batch_{n // stoerrelse:02d}.json"
        sti.write_text(json.dumps(del_, ensure_ascii=False, indent=1))
        print(f"→ {sti} ({len(del_)} boliger)")


def indlaes(mappe: Path, boliger: list[dict], cache: dict) -> None:
    efter_id = {b["id"]: b for b in boliger}
    antal = 0
    for sti in sorted(mappe.glob("svar_*.json")):
        for svar in json.loads(sti.read_text()):
            b = efter_id.get(svar.get("id"))
            if not b:
                continue
            svar = valider({k: v for k, v in svar.items() if k != "id"})
            svar["_plan"] = noegle(b)
            cache[b["id"]] = svar
            antal += 1
    gem_cache(cache)
    print(f"Indlæst {antal} analyser → {CACHE.relative_to(ROOT)}")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--batches", type=Path)
    ap.add_argument("--str", type=int, default=24)
    ap.add_argument("--indlaes", type=Path)
    args = ap.parse_args()

    boliger = json.loads(ANNONCER.read_text())["boliger"]
    cache = laes_cache()
    if args.indlaes:
        return indlaes(args.indlaes, boliger, cache)
    nye = mangler(boliger, cache)
    print(f"{len(nye)} boliger mangler analyse")
    if args.batches:
        return skriv_batches(nye, args.batches, args.str)
    if nye:
        analyser_med_claude(nye, cache)


if __name__ == "__main__":
    main()

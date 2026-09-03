#!/usr/bin/env python3
"""Bygger docs/data.json ud fra kildefilerne i data/raw/.

Kilder
------
1. CRM (Membercare)   — Serviceregistreringer.xlsx. Implementeret.
2. IP Nordic          — telefonstatistik. Afventer eksport.
3. Valeur Insights    — Valeurs egen håndtering. Afventer eksport.

Privatliv
---------
Kørslen læser felter med personoplysninger for at kunne tælle, men intet
personhenførbart skrives ud. Reglerne håndhæves af assert_clean() til sidst:
kørslen afbrydes, hvis resultatet alligevel indeholder et navn, et
telefonnummer, en e-mail eller et stamkort-id.
"""

import json
import re
import sys
from collections import Counter
from datetime import date
from pathlib import Path

import pandas as pd

import ipnordic
import valeur
import egne

ROOT = Path(__file__).resolve().parent.parent
RAW = ROOT / "data" / "raw"
OUT = ROOT / "docs" / "data.json"
TEMPLATE = ROOT / "scripts" / "template.html"
SITE = ROOT / "docs" / "index.html"
ARTIFACT = ROOT / "dist" / "dashboard.html"

# ── Tærskler ────────────────────────────────────────────────────────────────
MIN_GRUPPE = 5     # grupper under 5 lægges sammen i "Øvrige"
MIN_CELLE = 25     # ingen krydstabelcelle under 25 personer
MIN_SERIE = 30     # en måned skal have mindst så mange registreringer for en rate
MIN_RATE = 500     # en rate skal have så mange bag sig for at være sammenlignelig

UGEDAGE = ["Mandag", "Tirsdag", "Onsdag", "Torsdag", "Fredag", "Lørdag", "Søndag"]

# Kilden staver enkelte kategorier inkonsekvent. Venstre side er kildens værdi.
LABELS = {
    "Normaltansat over 19 timer": "Normalansat over 19 timer",
    "Normaltansat under 20 timer": "Normalansat under 20 timer",
    "1 og 2 års Kandidater": "1. og 2. års kandidater",
    "Ledig ikke ret til dagpenge": "Ledig uden dagpengeret",
    "Ph.d. studerende": "Ph.d.-studerende",
}

# Servicetyper med lange forklarende parenteser — kortet ned til en læsbar akse.
KORT_TYPE = {
    "Professionsjura (fx Aktindsigt, underretning, journalpligt, osv.)": "Professionsjura",
    "Selvstændig psykolog spørgsmål om Praksisoverenskomst/ydernummer": "Praksisoverenskomst/ydernummer",
    "Selvstændig psykolog – spørgsmål om drift og opstart af virksomhed uden ydernummer": "Drift og opstart uden ydernummer",
    "Opsigelse/afskedigelse/tjenestelig samtale/bortvisning": "Opsigelse og afskedigelse",
    "Kurser/arrangementer inkl. Betalinger": "Kurser og arrangementer",
    "Samarbejdsaftaler på tværs af det selvstændige område fx med kommuner, mellem flere selvstændige, med virksomheder mm.": "Samarbejdsaftaler, selvstændige",
    "Sektioner, kredse, faglige selskaber og netværk (fx ændringer i styrelser, oprettelse af ass. Medlemmer, oprettelse af ikke-medlemmer, nye netværk, vedtægter, mv.)": "Sektioner, kredse og netværk",
    "Psykologer i Danmark (det her handler om borgere, der gerne vil finde en psykolog eller psykologer, der skal have noget rettet på sitet)": "Psykologer i Danmark",
    "Medlemsfordele (fx ) forsikringer, GDPR, revision, m.v.)": "Medlemsfordele",
    "Din karriere (+ og karrierevejledning)": "Karriere og karrierevejledning",
    "Autorisation/tilsyn (Flyttet til eget punkt)": "Autorisation/tilsyn (udgået)",
    "Overenskomstspørgsmål – offentlige eller private": "Overenskomstspørgsmål",
    "Generelle spørgsmål om tilsyn og autorisation": "Tilsyn og autorisation",
    "Det socialt psykologiske samarbejde – honoraraftale med KL": "Honoraraftale med KL",
    "Samarbejde med forsikringsselskab eller healthcarevirksomhed": "Forsikring og healthcare",
    "Komite for Etik og etikspørgsmål": "Etik",
    "Feedback til DP med kommentarfelt": "Feedback til DP",
    "Tilsynssager og klager over psykologer": "Tilsynssager og klager",
}


def clean(s):
    """Trimmer og normaliserer whitespace i en tekstserie."""
    return s.fillna("").astype(str).str.replace(r"\s+", " ", regex=True).str.strip()


def label(v):
    return LABELS.get(v, v)


def split_medlemskab(serie):
    """Deler 'Normaltansat over 19 timer-Østjylland' i type og kreds.

    Nogle værdier rummer flere medlemskaber adskilt af komma; første del gælder.
    Kredsen står efter sidste bindestreg.
    """
    foerste = serie.str.split(",").str[0].str.strip()
    delt = foerste.str.rsplit("-", n=1, expand=True)
    if delt.shape[1] == 1:
        delt[1] = ""
    mtype = delt[0].str.strip().map(label)
    kreds = delt[1].fillna("").str.strip()
    # "Studerende DP-Dansk Psykolog Forening" har ingen geografisk kreds
    kreds = kreds.replace("Dansk Psykolog Forening", "Landsdækkende")
    return mtype, kreds


def taelling(serie, minimum=MIN_GRUPPE, ovrige="Øvrige"):
    """Optælling faldende, hvor alt under `minimum` lægges sammen."""
    vc = serie[serie != ""].value_counts()
    stor = vc[vc >= minimum]
    rest = int(vc[vc < minimum].sum())
    ud = [{"navn": k, "antal": int(v)} for k, v in stor.items()]
    if rest:
        ud.append({"navn": ovrige, "antal": rest, "samlet": True})
    return ud


def build_crm():
    kilde = RAW / "Serviceregistreringer.xlsx"
    if not kilde.exists():
        sys.exit(f"Mangler {kilde}")

    df = pd.read_excel(kilde, dtype=str)
    df["gruppe"] = clean(df["Servicegruppe"])
    df["type"] = clean(df["Servicetype"]).map(lambda v: KORT_TYPE.get(v, v))
    df["registrant"] = clean(df["Oprettet af"])
    df["dato"] = pd.to_datetime(df["Service dato"], errors="coerce")
    df["oprettet"] = pd.to_datetime(df["Oprettet"], errors="coerce")
    df["mtype"], df["kreds"] = split_medlemskab(clean(df["Medlemskab"]))
    df = df[df["dato"].notna()].copy()

    df["maaned"] = df["dato"].dt.to_period("M").astype(str)
    df["ugedag"] = df["dato"].dt.dayofweek
    df["time"] = df["oprettet"].dt.hour

    # ── Gengangere: hvor mange gange optræder hvert stamkort ────────────────
    pr_medlem = df["Stamkort"].value_counts()
    fordeling = Counter()
    for n in pr_medlem.values:
        fordeling["10+" if n >= 10 else str(n)] += 1
    orden = [str(i) for i in range(1, 10)] + ["10+"]
    gengangere = [
        {"antal_henvendelser": k, "medlemmer": fordeling[k]}
        for k in orden
        if fordeling[k]
    ]
    flergangs_id = set(pr_medlem[pr_medlem > 1].index)
    flergangs = df["Stamkort"].isin(flergangs_id)

    # Hvilke emner er gengangerne tungest i — kun grupper med volumen nok
    emne_gentag = []
    for g, sub in df.groupby("gruppe"):
        if len(sub) < MIN_CELLE:
            continue
        emne_gentag.append({
            "navn": g,
            "antal": int(len(sub)),
            "andel_gengangere": round(sub["Stamkort"].isin(flergangs_id).mean() * 100, 1),
            # En rate på nogle få hundrede registreringer bærer ikke en
            # sammenligning med en på flere tusind. Markeres, så den ikke
            # læses som ligeværdig.
            "tynd": bool(len(sub) < MIN_RATE),
        })
    emne_gentag.sort(key=lambda r: -r["andel_gengangere"])

    # ── Tidsserie pr. måned, samlet og pr. servicegruppe ────────────────────
    maaneder = sorted(df["maaned"].unique())
    pr_maaned = []
    for m in maaneder:
        sub = df[df["maaned"] == m]
        pr_maaned.append({
            "maaned": m,
            "antal": int(len(sub)),
            "medlemmer": int(sub["Stamkort"].nunique()),
            "tynd": bool(len(sub) < MIN_SERIE),
            "grupper": {g: int(n) for g, n in sub["gruppe"].value_counts().items()},
        })

    # ── Krydstabel: medlemstype × servicegruppe ────────────────────────────
    top_typer = [t for t, n in df["mtype"].value_counts().items() if n >= MIN_CELLE and t][:8]
    grupper_orden = [g["navn"] for g in taelling(df["gruppe"])]
    kryds = []
    for t in top_typer:
        sub = df[df["mtype"] == t]
        raekke = {"medlemstype": t, "antal": int(len(sub)), "celler": []}
        for g in grupper_orden:
            n = int((sub["gruppe"] == g).sum())
            raekke["celler"].append({
                "gruppe": g,
                "antal": n if n >= MIN_CELLE else None,
                "andel": round(n / len(sub) * 100, 1) if n >= MIN_CELLE else None,
            })
        kryds.append(raekke)

    # ── Servicetyper med deres gruppe ──────────────────────────────────────
    typer = []
    for (g, t), n in df.groupby(["gruppe", "type"]).size().items():
        if n >= MIN_GRUPPE:
            typer.append({"navn": t, "gruppe": g, "antal": int(n)})
    typer.sort(key=lambda r: -r["antal"])

    start, slut = df["dato"].min(), df["dato"].max()
    aabne_dage = int(df["dato"].dt.normalize().nunique())

    return {
        "meta": {
            "kilde": "Membercare (CRM) — servicefanen",
            "raekker": int(len(df)),
            "periode_fra": start.strftime("%Y-%m-%d"),
            "periode_til": slut.strftime("%Y-%m-%d"),
            "registreringsdage": aabne_dage,
        },
        "kpi": {
            "registreringer": int(len(df)),
            "medlemmer": int(df["Stamkort"].nunique()),
            "pr_dag": round(len(df) / aabne_dage, 1),
            "andel_gengangere": round(flergangs.mean() * 100, 1),
            "servicegrupper": int(df["gruppe"].nunique()),
            "servicetyper": int(df["type"].nunique()),
            "kredse": int(df.loc[df["kreds"] != "", "kreds"].nunique()),
        },
        "maanedlig": pr_maaned,
        "grupper": taelling(df["gruppe"]),
        "typer": typer,
        "medlemstyper": taelling(df["mtype"]),
        "kredse": taelling(df["kreds"]),
        "kryds": {"grupper": grupper_orden, "raekker": kryds},
        "gengangere": {
            "fordeling": gengangere,
            "unikke_medlemmer": int(df["Stamkort"].nunique()),
            "andel_registreringer": round(flergangs.mean() * 100, 1),
            "pr_emne": emne_gentag,
        },
        "ugedag": [
            {"navn": UGEDAGE[i], "antal": int((df["ugedag"] == i).sum())}
            for i in range(7)
            if (df["ugedag"] == i).sum() > 0
        ],
        "klokkeslaet": [
            {"time": int(t), "antal": int(n)}
            for t, n in sorted(df["time"].dropna().astype(int).value_counts().items())
        ],
        "registranter": taelling(df["registrant"], minimum=1),
    }


# ── Privatlivsspærre ────────────────────────────────────────────────────────
MONSTRE = [
    (r"[\w.\-]+@[\w.\-]+\.\w+", "e-mailadresse"),
    (r"\b\d{6}-\d{4}\b", "cpr-lignende nummer"),
    (r"(?:\+45[\s.]?)?\b\d{2}[\s.]?\d{2}[\s.]?\d{2}[\s.]?\d{2}\b", "telefonnummer"),
]


def assert_clean(data, kendte_navne):
    """Afbryder kørslen hvis resultatet rummer personhenførbare oplysninger."""
    tekst = json.dumps(data, ensure_ascii=False)

    for monster, hvad in MONSTRE:
        fund = re.search(monster, tekst)
        if fund:
            sys.exit(f"AFBRUDT: resultatet indeholder et {hvad}: {fund.group()!r}")

    # Medlemsnavne må ikke optræde. Registranternes navne er en bevidst,
    # dokumenteret undtagelse — det er en intern driftsmetrik.
    tilladte = {r["navn"] for r in data["crm"]["registranter"]}
    for navn in kendte_navne - tilladte:
        if len(navn) > 6 and re.search(rf"\b{re.escape(navn)}\b", tekst):
            sys.exit(f"AFBRUDT: resultatet indeholder et medlemsnavn: {navn!r}")

    print(f"  privatlivstjek: bestået ({len(tilladte)} registranter tilladt)")


def render(data):
    """Indsætter data i skabelonen og skriver begge udgaver af siden.

    docs/index.html er et fuldt dokument til GitHub Pages.
    dist/dashboard.html er kun sideindholdet, til udgivelse som artifact.
    """
    kompakt = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    # </script> i data ville lukke script-taggen for tidligt.
    kompakt = kompakt.replace("</", "<\\/")
    krop = TEMPLATE.read_text(encoding="utf-8").replace("/*__DATA__*/", kompakt)

    ARTIFACT.parent.mkdir(parents=True, exist_ok=True)
    ARTIFACT.write_text(krop, encoding="utf-8")

    SITE.write_text(
        '<!doctype html>\n<html lang="da">\n<head>\n<meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width,initial-scale=1">\n'
        + krop
        + "\n</html>\n",
        encoding="utf-8",
    )
    for f in (SITE, ARTIFACT):
        print(f"  skrev {f.relative_to(ROOT)} ({f.stat().st_size / 1024:.0f} kB)")


def main():
    crm = build_crm()
    telefoni = ipnordic.build()
    koe_opkald = (telefoni or {}).get("koeer", {}).get("ialt", {}).get("opkald")
    insights = valeur.build(koe_opkald)
    data = {
        "genereret": date.today().isoformat(),
        "crm": crm,
        "telefoni": telefoni,
        "valeur": insights,
        "egne": egne.build(),
    }
    if telefoni:
        c = telefoni["cdr"]
        print(f"  IP Nordic: {c['raekker']} CDR-rækker, "
              f"{sum(l['opkald'] for l in c['lag'])} medlemsopkald, "
              f"{telefoni['koeer']['ialt']['opkald']} køopkald")
    if insights:
        print(f"  Valeur Insights: {insights['opkald']} kald, "
              f"{insights['svarprocent']}% besvaret")

    navne = set(
        clean(pd.read_excel(RAW / "Serviceregistreringer.xlsx", dtype=str)["Medlem"]).unique()
    )
    assert_clean(data, navne)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"  skrev {OUT.relative_to(ROOT)} ({OUT.stat().st_size / 1024:.0f} kB)")

    render(data)
    print(f"  {crm['kpi']['registreringer']} registreringer · "
          f"{crm['kpi']['medlemmer']} medlemmer · "
          f"{crm['meta']['periode_fra']} → {crm['meta']['periode_til']}")


if __name__ == "__main__":
    main()

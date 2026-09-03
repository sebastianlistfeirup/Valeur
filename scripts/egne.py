"""Læser DP's egen opgørelse over opkald til sekretariatet, 2025 mod 2026.

Filen er sammenstillet i hånden fra IP Nordic-udtræk og dækker januar til
marts. Den er værdifuld netop dér, hvor de øvrige kilder ikke rækker:
januar-marts 2025 ligger før Valeur gik i luften 5. maj 2025, så det er den
eneste egentlige før-måling der findes. Tallene er DP's egne og er ikke
efterregnet her.
"""

import datetime as dt
from pathlib import Path

import pandas as pd

RAW = Path(__file__).resolve().parent.parent / "data" / "raw" / "egne"
MAANEDER = {1: "Januar", 2: "Februar", 3: "Marts"}


def build():
    sti = RAW / "oversigt_2025_2026.xlsx"
    if not sti.exists():
        return None
    df = pd.read_excel(sti, header=0)
    df = df.rename(columns={df.columns[0]: "afdeling"})
    # Arket rummer to tabeller under hinanden. Kun den første holder 2025 op
    # mod 2026 — stop ved den tomme række der skiller dem.
    tom = df["afdeling"].isna()
    if tom.any():
        df = df.iloc[: int(tom.idxmax())]

    # Kolonnerne er par af (2026, 2025) pr. måned, hver efterfulgt af en delta.
    par = []
    for i, c in enumerate(df.columns[1:], start=1):
        if isinstance(c, (dt.datetime, dt.date, pd.Timestamp)):
            par.append((i, c.year, c.month))
    maaneder = sorted({m for _, _, m in par})

    raekker = []
    for _, r in df.iterrows():
        navn = str(r["afdeling"]).strip()
        if not navn or navn == "nan" or "pr. dag" in navn.lower():
            continue
        pr_maaned = {}
        for i, aar, m in par:
            v = pd.to_numeric(r.iloc[i], errors="coerce")
            if pd.notna(v):
                pr_maaned.setdefault(m, {})[aar] = int(v)
        komplet = {m: d for m, d in pr_maaned.items() if len(d) == 2}
        if not komplet:
            continue
        foer = sum(d[min(d)] for d in komplet.values())
        efter = sum(d[max(d)] for d in komplet.values())
        raekker.append({
            "navn": navn.replace(" (I alt)", "").replace(" (i alt)", "").strip(),
            "foer": foer, "efter": efter,
            "aendring": round((efter - foer) / foer * 100, 1) if foer else None,
            "i_alt_raekke": navn.strip().lower() == "i alt",
        })

    total = next((r for r in raekker if r["i_alt_raekke"]), None)
    raekker = [r for r in raekker if not r["i_alt_raekke"]]
    raekker.sort(key=lambda r: -(r["foer"] + r["efter"]))

    aar = sorted({a for _, a, _ in par})
    return {
        "kilde": "DP's egen opgørelse, sammenstillet fra IP Nordic",
        "foer_aar": min(aar), "efter_aar": max(aar),
        "maaneder": [MAANEDER.get(m, str(m)) for m in maaneder],
        "raekker": raekker,
        "i_alt": total or {
            "foer": sum(r["foer"] for r in raekker),
            "efter": sum(r["efter"] for r in raekker),
        },
    }

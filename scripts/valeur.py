"""Læser eksporterne fra Valeur Insights.

Insights er Valeurs eget BI-værktøj, og hver eksport er ét diagram. De to
brugbare til historik er dagsopgørelsen (kald i alt og besvarede pr. dag) og
de to timefordelinger — den ene over alle kald, den anden over dem der blev
afbrudt inden nogen nåede at tage den.

Læses med header=2, fordi række 1 er en filterbeskrivelse og række 2 er tom.
"""

from pathlib import Path

import pandas as pd

RAW = Path(__file__).resolve().parent.parent / "data" / "raw" / "valeur"
UGEDAGE_EN = {"Monday": 0, "Tuesday": 1, "Wednesday": 2, "Thursday": 3,
              "Friday": 4, "Saturday": 5, "Sunday": 6}
UGEDAGE = ["Mandag", "Tirsdag", "Onsdag", "Torsdag", "Fredag", "Lørdag", "Søndag"]


def _tabel(navn, kolonner):
    sti = RAW / navn
    if not sti.exists():
        return None
    df = pd.read_excel(sti, header=2)
    df.columns = kolonner[:len(df.columns)]
    return df


def build(koe_opkald=None):
    """koe_opkald: antal opkald sekretariatets køer modtog i samme periode.
    Bruges til at anslå, hvor stor en del Valeur klarer selv."""
    d = _tabel("opkald_per_dag.xlsx", ["dato", "besvarede", "ialt"])
    if d is None:
        return None
    d["dato"] = pd.to_datetime(d["dato"], errors="coerce")
    for k in ("besvarede", "ialt"):
        d[k] = pd.to_numeric(d[k], errors="coerce").fillna(0)
    d = d[d["dato"].notna() & (d["ialt"] > 0)]

    ialt, besvarede = int(d["ialt"].sum()), int(d["besvarede"].sum())

    maanedlig = []
    for m, s in d.groupby(d["dato"].dt.to_period("M")):
        o, b = int(s["ialt"].sum()), int(s["besvarede"].sum())
        maanedlig.append({
            "maaned": str(m), "opkald": o, "besvarede": b,
            "svarprocent": round(b / o * 100, 1) if o else None,
            # En måned med en håndfuld kald bærer ikke en rate.
            "tynd": bool(o < 50),
        })

    ugedag = [
        {"navn": UGEDAGE[i],
         "opkald": int(s["ialt"].sum()), "besvarede": int(s["besvarede"].sum())}
        for i, s in d.groupby(d["dato"].dt.dayofweek)
    ]

    def _fordeling(fil, kolonner):
        t = _tabel(fil, kolonner)
        if t is None:
            return None
        t[kolonner[0]] = pd.to_numeric(t[kolonner[0]], errors="coerce")
        t[kolonner[1]] = pd.to_numeric(t[kolonner[1]], errors="coerce")
        t = t.dropna().sort_values(kolonner[0])
        return [{"time": int(r[0]), "andel": round(float(r[1]) * 100, 1)}
                for r in t.itertuples(index=False)]

    # Hvor meget Valeur klarer uden at sende videre. Et skøn, ikke en måling:
    # køtallet kommer fra IP Nordic og dækker en lidt anden periode.
    selv = None
    if koe_opkald and besvarede:
        rest = besvarede - koe_opkald
        if rest > 0:
            selv = {
                "besvarede": besvarede,
                "videre": int(koe_opkald),
                "selv": int(rest),
                "andel": round(rest / besvarede * 100, 1),
                "maal": 70,
            }

    return {
        "kilde": "Valeur Insights",
        "periode_fra": d["dato"].min().strftime("%Y-%m-%d"),
        "periode_til": d["dato"].max().strftime("%Y-%m-%d"),
        "dage": int(len(d)),
        "opkald": ialt,
        "besvarede": besvarede,
        "svarprocent": round(besvarede / ialt * 100, 1) if ialt else None,
        "maanedlig": maanedlig,
        "ugedag": ugedag,
        "timefordeling": _fordeling("timefordeling.xlsx", ["time", "andel"]),
        "afbrudt_pr_time": _fordeling("timefordeling_afbrudt.xlsx", ["time", "andel"]),
        "haandtering": selv,
    }

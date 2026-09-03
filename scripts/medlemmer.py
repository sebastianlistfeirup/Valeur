"""Læser de anonymiserede medlemslister og holder dem op mod henvendelserne.

To øjebliksbilleder: maj 2025, da samarbejdet startede, og september 2026.
Listerne har ingen navne, men stamkortnummeret er den samme nøgle som i CRM.
Det gør det muligt at svare på briefingens spørgsmål — ringer de "rigtige"
medlemmer? — ved at holde hver gruppes andel af henvendelserne op mod dens
andel af bestanden.

Privatliv
---------
Stamkortnummeret bruges udelukkende til at slå en registrering op i bestanden
og skrives aldrig ud. Alder, kreds og kommune kan tilsammen udpege en person,
så alt opgøres i grupper, og grupper under tærsklen slås sammen.
"""

from pathlib import Path

import pandas as pd

RAW = Path(__file__).resolve().parent.parent / "data" / "raw" / "medlemmer"
MIN_GRUPPE = 25          # ingen gruppe under 25 personer vises
ALDERSBAAND = [(0, 25, "under 26"), (26, 30, "26–30"), (31, 35, "31–35"),
               (36, 40, "36–40"), (41, 45, "41–45"), (46, 50, "46–50"),
               (51, 55, "51–55"), (56, 60, "56–60"), (61, 65, "61–65"),
               (66, 70, "66–70"), (71, 200, "over 70")]

LABELS = {
    "Normaltansat over 19 timer": "Normalansat over 19 timer",
    "Normaltansat under 20 timer": "Normalansat under 20 timer",
    "1 og 2 års Kandidater": "1. og 2. års kandidater",
}


def _baand(alder):
    if pd.isna(alder):
        return None
    for lav, hoej, navn in ALDERSBAAND:
        if lav <= alder <= hoej:
            return navn
    return None


def _aktive(df, paa):
    """Aktiv på et givet tidspunkt: ingen udmeldelsesdato, eller en der ligger
    efter. Listerne rummer udmeldelser med varsel, så en dato i fremtiden
    betyder stadig medlem på opgørelsestidspunktet."""
    return df[df["udm"].isna() | (df["udm"] > pd.Timestamp(paa))]


def _laes(sti):
    df = pd.read_excel(sti, dtype=str)
    df["stamkort"] = df["Stamkort"].astype(str).str.strip()
    df["mtype"] = (df["Medlemstype"].fillna("").str.strip()
                   .str.split(" - ").str[0].map(lambda v: LABELS.get(v, v)))
    df["kreds"] = df["Kreds"].fillna("").str.strip()
    df["koen"] = df["Køn"].fillna("").str.strip()
    df["alder"] = pd.to_numeric(df["Alder"], errors="coerce")
    df["baand"] = df["alder"].map(_baand)
    df["udm"] = pd.to_datetime(df["Udm.dato"], errors="coerce")
    df["ydernummer"] = df["Har ydernummer"].fillna("").str.strip() != ""
    return df.drop_duplicates(subset="stamkort")


def _andele(serie, minimum=MIN_GRUPPE):
    vc = serie[serie.astype(str).str.len() > 0].value_counts()
    stor = vc[vc >= minimum]
    rest = int(vc[vc < minimum].sum())
    ud = [{"navn": k, "antal": int(v)} for k, v in stor.items()]
    if rest >= minimum:
        ud.append({"navn": "Øvrige", "antal": rest})
    i_alt = sum(r["antal"] for r in ud)
    for r in ud:
        r["andel"] = round(r["antal"] / i_alt * 100, 1) if i_alt else 0
    return ud


def _repraesentation(bestand, crm_pr_gruppe, felt, aktive):
    """Holder hver gruppes andel af henvendelserne op mod dens andel af bestanden."""
    b = aktive[felt]
    b = b[b.astype(str).str.len() > 0].value_counts()
    b_ialt = int(b.sum())
    h_ialt = sum(crm_pr_gruppe.values())
    ud = []
    for navn, antal in b.items():
        if antal < MIN_GRUPPE:
            continue
        henv = int(crm_pr_gruppe.get(navn, 0))
        if henv < MIN_GRUPPE:
            continue
        b_pct = antal / b_ialt * 100
        h_pct = henv / h_ialt * 100
        ud.append({
            "navn": navn,
            "medlemmer": int(antal), "bestand_andel": round(b_pct, 1),
            "henvendelser": henv, "henv_andel": round(h_pct, 1),
            "faktor": round(h_pct / b_pct, 2) if b_pct else None,
        })
    ud.sort(key=lambda r: -(r["faktor"] or 0))
    return ud


def build(crm_df=None):
    filer = sorted(RAW.glob("bestand_*.xlsx"))
    if not filer:
        return None
    snapshots = {}
    for f in filer:
        naar = f.stem.replace("bestand_", "")
        snapshots[naar] = _laes(f)

    tider = sorted(snapshots)
    dato = {t: f"{t}-01" for t in tider}
    foer, sen = snapshots[tider[0]], snapshots[tider[-1]]
    aktive = _aktive(sen, dato[tider[-1]])

    ud = {
        "snapshots": [
            {"naar": t, "medlemmer": int(len(snapshots[t])),
             "aktive": int(len(_aktive(snapshots[t], dato[t])))}
            for t in tider
        ],
        "bestand": {
            "naar": tider[-1],
            "medlemmer": int(len(aktive)),
            "medlemstyper": _andele(aktive["mtype"]),
            "kredse": _andele(aktive["kreds"]),
            "alder": [
                {"navn": n, "antal": int((aktive["baand"] == n).sum())}
                for _, _, n in ALDERSBAAND
                if (aktive["baand"] == n).sum() >= MIN_GRUPPE
            ],
            "koen": _andele(aktive["koen"]),
            "median_alder": float(aktive["alder"].median()),
            "ydernummer": int(aktive["ydernummer"].sum()),
        },
        "vaekst": {
            "fra": tider[0], "til": tider[-1],
            "foer": int(len(_aktive(foer, dato[tider[0]]))),
            "efter": int(len(aktive)),
        },
    }
    for a in ud["bestand"]["alder"]:
        i_alt = sum(x["antal"] for x in ud["bestand"]["alder"])
        a["andel"] = round(a["antal"] / i_alt * 100, 1) if i_alt else 0

    # ── Kobling til henvendelserne ─────────────────────────────────────────
    if crm_df is not None and "Stamkort" in crm_df:
        crm = crm_df.copy()
        crm["stamkort"] = crm["Stamkort"].astype(str).str.strip()
        crm["dato"] = pd.to_datetime(crm["Service dato"], errors="coerce")
        skel = pd.Timestamp(dato[tider[-1]]) - (
            pd.Timestamp(dato[tider[-1]]) - pd.Timestamp(dato[tider[0]])) / 2
        kolonner = ["stamkort", "mtype", "kreds", "baand", "koen", "alder"]
        dele = []
        for maerke, snap in ((tider[0], foer), (tider[-1], sen)):
            valg = crm["dato"] < skel if maerke == tider[0] else crm["dato"] >= skel
            d = crm[valg].merge(snap[kolonner], on="stamkort", how="left")
            d["bestand"] = maerke
            dele.append(d)
        slaaet_op = pd.concat(dele, ignore_index=True)
        traf = slaaet_op["mtype"].notna()
        ud["kobling"] = {
            "registreringer": int(len(crm)),
            "fundet": int(traf.sum()),
            "andel": round(traf.mean() * 100, 1),
            "skel": skel.strftime("%Y-%m-%d"),
        }
        fundet = slaaet_op[traf]
        # Unikke medlemmer, ikke registreringer — ellers vejer en genganger tungt.
        # Kun den seneste halvdel: dér er bestanden samtidig med henvendelsen.
        nyeste = fundet[fundet["bestand"] == tider[-1]]
        pr_person = nyeste.drop_duplicates(subset="stamkort")
        for felt, noegle in [("mtype", "medlemstype"), ("baand", "alder"),
                             ("kreds", "kreds"), ("koen", "koen")]:
            tael = pr_person[felt].value_counts().to_dict()
            ud.setdefault("repraesentation", {})[noegle] = _repraesentation(
                sen, tael, felt, aktive)
        ud["kobling"]["personer"] = int(len(pr_person))
        ud["kobling"]["median_alder_ringer"] = float(pr_person["alder"].median())
    return ud

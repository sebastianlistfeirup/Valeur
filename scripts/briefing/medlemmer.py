"""
Økonomi og medlemmer: pris pr. opkald (Valeur Insights) og medlemslisterne koblet til Membercare.

Medlemslisterne indeholder kontaktoplysninger. Scriptet læser KUN kolonnerne
kreds, medlemstype, stamkort, alder og DP Indm./Udm.dato (efter position, da
overskrifterne varierer mellem filerne). Stamkort bruges kun til koblingen og
forlader aldrig scriptet. Ingen gruppe under MIN_GRUPPE personer vises.
"""
import glob, re
from pathlib import Path
import numpy as np
import pandas as pd

START = pd.Timestamp("2025-05-05")
MIN_GRUPPE = 25
KOLONNER = {0: "kreds", 1: "type", 2: "stamkort", 8: "alder", 20: "dp_ind", 21: "dp_ud"}
KERNE = ["Normaltansat over 19 timer", "Selvstændig"]
KREDS_NAVN = {"Kreds Trekantsområdet": "Trekantsområdet", "MIDTVESTjylland": "Midtvestjylland",
              "Viborg \"Limfjord\"": "Viborg-Limfjord", "Dansk Psykolog Forening": "Uden kreds"}
ALDER = [0, 25, 30, 35, 40, 45, 50, 55, 60, 65, 70, 200]
ALDER_NAVN = ["Under 26", "26–30", "31–35", "36–40", "41–45", "46–50", "51–55", "56–60", "61–65", "66–70", "Over 70"]


# ------------------------------------------------------------------ økonomi
def oekonomi(root):
    fil = root / "Nye Valeur træk" / "Opkaldsdata (5).xlsx"
    if not fil.exists():
        return None
    t = pd.read_excel(fil, header=None)
    hoved = t.index[t[0].astype(str).str.strip().eq("Kundenavn")][0]
    r = t.iloc[hoved + 1]
    besv, tale, efter, kr = float(r[2]), float(r[3]), float(r[4]), float(r[5])
    return {"besvarede": int(besv), "tale_min": round(tale), "efter_min": round(efter), "kr": round(kr),
            "kr_pr_opkald": round(kr / besv, 2), "kr_pr_min": round(kr / (tale + efter), 2),
            "tale_pr_opkald": round(tale / besv * 60), "efter_pr_opkald": round(efter / besv * 60),
            "timer": round((tale + efter) / 60)}


# ------------------------------------------------------------------ medlemslister
def _laes(root):
    ud = []
    for f in sorted(glob.glob(str(root / "Medlemslister" / "*.xlsx"))):
        m = re.findall(r"periode_(\d\d_\d\d_\d\d)", f)
        if not m:
            continue
        d = pd.read_excel(f, header=None, skiprows=1, usecols=list(KOLONNER), dtype={2: str})
        d.columns = [KOLONNER[c] for c in d.columns]
        d["snap"] = pd.to_datetime(m[0], format="%d_%m_%y")
        for c in ("dp_ind", "dp_ud"):
            d[c] = pd.to_datetime(d[c], errors="coerce")
        d["stamkort"] = d["stamkort"].astype("string").str.strip()
        d["kreds"] = d["kreds"].astype("string").str.strip().replace(KREDS_NAVN)
        ud.append(d)
    return pd.concat(ud) if ud else None


def _aktive(p, dato):
    x = p[(p["snap"] == dato) & ~(p["dp_ud"] <= dato)]
    return x.drop_duplicates("stamkort") if x["stamkort"].notna().any() else x   # 18/11-26-listen har ingen stamkort


def _andele(df, felt, vaerdi="k"):
    g = df.groupby(felt, observed=True)[vaerdi].agg(["mean", "size"])
    g = g[g["size"] >= MIN_GRUPPE]
    return [{"g": str(k), "a": round(float(r["mean"]), 3), "n": int(r["size"])} for k, r in g.iterrows()]


def _diff(a, b):
    """Forskel i andel b-a med 95 %-interval (normaltilnærmelse)."""
    p1, n1, p2, n2 = a.mean(), len(a), b.mean(), len(b)
    se = np.sqrt(p1 * (1 - p1) / n1 + p2 * (1 - p2) / n2)
    return {"uden": round(p1, 4), "n_uden": int(n1), "med": round(p2, 4), "n_med": int(n2),
            "diff": round(p2 - p1, 4), "lav": round(p2 - p1 - 1.96 * se, 4), "hoej": round(p2 - p1 + 1.96 * se, 4)}


def medlemmer(root):
    p = _laes(root)
    if p is None:
        return None
    snaps = sorted(p["snap"].unique())
    bestand = [{"d": pd.Timestamp(s).strftime("%Y-%m-%d"), "n": int(len(_aktive(p, s)))} for s in snaps]
    p = p[p["stamkort"].notna()]
    koblet = sorted(p["snap"].unique())
    sidste = pd.Timestamp(koblet[-1])
    seneste = p.sort_values("snap").drop_duplicates("stamkort", keep="last").set_index("stamkort")
    paa_sidste = set(p[p["snap"] == sidste]["stamkort"])

    mc = pd.read_excel(root / "Membercare/Registreringer i servicefanen.xlsx", dtype=str,
                       usecols=["Stamkort", "Servicetype", "Service dato"])
    mc["dato"] = pd.to_datetime(mc["Service dato"], errors="coerce")
    mc["Stamkort"] = mc["Stamkort"].astype("string").str.strip()
    mc = mc[mc["dato"] >= START]
    mc["kont"] = mc["Servicetype"].fillna("").str.strip().eq("Medlemskab og kontingent")
    koblingsgrad = round(mc["Stamkort"].isin(set(p["stamkort"])).mean(), 3)

    # --- dækning: medlemmer på seneste koblede liste, kontakt de seneste 12 måneder
    m = _aktive(p, sidste).copy()
    k12 = set(mc[(mc["dato"] > sidste - pd.DateOffset(months=12)) & (mc["dato"] <= sidste)]["Stamkort"])
    m["k"] = m["stamkort"].isin(k12)
    m["alder_b"] = pd.cut(m["alder"], ALDER, labels=ALDER_NAVN)
    m["kreds"] = m["kreds"].fillna("Uden kreds")
    pr_type = sorted(_andele(m, "type"), key=lambda x: -x["a"])
    pr_kreds = sorted(_andele(m, "kreds"), key=lambda x: -x["a"])
    daekning = {"dato": sidste.strftime("%Y-%m-%d"), "medlemmer": int(len(m)), "i_kontakt": int(m["k"].sum()),
                "andel": round(m["k"].mean(), 3), "pr_type": pr_type,
                "pr_alder": _andele(m, "alder_b"), "pr_kreds": pr_kreds}

    # --- udmeldelser 1: naiv sammenligning (kohorte 30/6-25, kontakt før udmeldelse)
    def udmeldt(df):
        return (~df["stamkort"].isin(paa_sidste)) | df["stamkort"].map(seneste["dp_ud"]).notna()

    k0 = pd.Timestamp("2025-06-30")
    b = _aktive(p, k0).copy()
    b["udmeldt"] = udmeldt(b)
    slut = pd.Series(b["stamkort"].values, index=b["stamkort"].values).map(seneste["dp_ud"]).fillna(sidste).clip(upper=sidste)
    v = mc[mc["Stamkort"].isin(slut.index)]
    v = v[v["dato"] <= v["Stamkort"].map(slut)]
    kun_kont = v.groupby("Stamkort")["kont"].all()
    antal = v.groupby("Stamkort").size()
    b["faglig"] = b["stamkort"].isin(kun_kont[~kun_kont].index)
    b["kontakt"] = b["stamkort"].isin(kun_kont.index)
    b["n"] = b["stamkort"].map(antal).fillna(0)
    bk = b[b["type"].isin(KERNE)]
    naiv = _diff(bk[~bk["kontakt"]]["udmeldt"], bk[bk["faglig"]]["udmeldt"])
    trappe = []
    for navn, lo, hi in (("Ingen", 0, 0), ("1", 1, 1), ("2–3", 2, 3), ("4 eller flere", 4, 10**6)):
        x = bk[(bk["n"] >= lo) & (bk["n"] <= hi) & (bk["faglig"] | ~bk["kontakt"])]
        if len(x) >= MIN_GRUPPE:
            trappe.append({"g": navn, "a": round(x["udmeldt"].mean(), 4), "n": int(len(x))})

    # --- udmeldelser 2: fast tidsramme (kontakt maj-dec 2025, udmeldelse derefter)
    L = pd.Timestamp("2025-12-31")
    c = _aktive(p, L).copy()
    c["udmeldt"] = udmeldt(c)
    w = mc[mc["dato"] <= L]
    kk = w.groupby("Stamkort")["kont"].all()
    c["kontakt"] = c["stamkort"].isin(kk.index)
    c["faglig"] = c["stamkort"].isin(kk[~kk].index)
    ck = c[c["type"].isin(KERNE)]
    fast = _diff(ck[~ck["kontakt"]]["udmeldt"], ck[ck["faglig"]]["udmeldt"])
    fast_type = []
    for t in KERNE:
        x = c[c["type"] == t]
        fast_type.append({"g": t, **_diff(x[~x["kontakt"]]["udmeldt"], x[x["faglig"]]["udmeldt"])})

    # --- nye medlemmer
    ny = p[p["dp_ind"] >= START].drop_duplicates("stamkort").copy()
    foerste = mc.groupby("Stamkort")["dato"].min()
    ny["foerste"] = ny["stamkort"].map(foerste)
    ny["dage"] = (ny["foerste"] - ny["dp_ind"]).dt.days
    ny["k"] = ny["foerste"].notna()
    ny["k90"] = ny["dage"].between(-30, 90)
    nye = {"antal": int(len(ny)), "andel_kontakt": round(ny["k"].mean(), 3), "andel_90": round(ny["k90"].mean(), 3),
           "pr_type": sorted(_andele(ny.assign(k=ny["k"]), "type"), key=lambda x: -x["n"])}

    return {"bestand": bestand, "koblingsgrad": koblingsgrad, "daekning": daekning,
            "udmeldelser": {"naiv": naiv, "trappe": trappe, "fast": fast, "fast_type": fast_type,
                            "kohorte_naiv": k0.strftime("%Y-%m-%d"), "kohorte_fast": L.strftime("%Y-%m-%d"),
                            "slut": sidste.strftime("%Y-%m-%d")},
            "nye": nye}


def build(root):
    root = Path(root)
    return {"oekonomi": oekonomi(root), "medlemmer": medlemmer(root)}

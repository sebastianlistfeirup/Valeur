"""
Dybere spadestik: seks vinkler ud over briefingen.

Bemærkningsfeltet i Membercare læses KUN for at klassificere hver registrering
(besked/viderestilling ja/nej, nævner en rykker ja/nej). Teksten forlader aldrig
scriptet; outputtet er udelukkende optællinger.
"""
import glob, re, datetime as dt
from pathlib import Path
import pandas as pd

START = pd.Timestamp("2025-05-05")
SKABELON_START = pd.Timestamp("2025-07-01")   # Valeurs faste beskedskabelon tages i brug
MIN_N = 100                                    # emner med færre registreringer vises ikke

BESKED = re.compile(r"Har ringet i dag|ringes op|ring(?:e)? (?:tilbage|op)|kontakt(?:es)?|videresend|"
                    r"stillet (?:videre|om)|stillede (?:videre|om)|omstill|viderestill|lagt besked|lagde besked|"
                    r"\b15(?:10|15|20|25|122|123)\b", re.I)
RYKKER = re.compile(r"rykker|rykkerbrev|betalingspåmindelse|påmindelse om betaling", re.I)
KREDS_NAVN = {"Kreds Trekantsområdet": "Trekantsområdet", "MIDTVESTjylland": "Midtvestjylland",
              "Viborg \"Limfjord\"": "Viborg-Limfjord", "Dansk Psykolog Forening": "Uden kreds"}
UGEDAGE = ["Mandag", "Tirsdag", "Onsdag", "Torsdag", "Fredag"]


def _kort(t):
    t = re.sub(r"^Selvstændig psykolog[ –-]*spørgsmål om ", "", t)
    t = re.sub(r"\s*\(.*?\)", "", t)
    return t.strip()


def _membercare(mc_sti):
    df = pd.read_excel(mc_sti, dtype=str)
    df["dato"] = pd.to_datetime(df["Service dato"], errors="coerce")
    df = df[df["dato"] >= START].copy()
    df["type"] = df["Servicetype"].fillna("Ikke angivet").str.strip().map(_kort)
    df["gruppe"] = df["Servicegruppe"].fillna("Ikke angivet").str.strip()
    tekst = df["Bemærkning"].fillna("")
    df["besked"] = tekst.str.contains(BESKED)
    df["rykker"] = (df["Overskrift"].fillna("") + " " + tekst).str.contains(RYKKER)
    df["tom_note"] = tekst.str.strip() == ""
    df.drop(columns=["Bemærkning", "Overskrift", "Medlem"], inplace=True)   # fritekst og navne ud
    df["md"] = df["dato"].dt.to_period("M").astype(str)
    return df


# 1 ------------------------------------------------------------------ løser Valeur selv?
def loeser_selv(df, min_n=MIN_N):
    d = df[df["dato"] >= SKABELON_START]
    if not len(d):
        return None
    pr_md = d.groupby("md").agg(n=("besked", "size"), besked=("besked", "sum")).reset_index()
    pr_gruppe = d.groupby("gruppe").agg(n=("besked", "size"), andel=("besked", "mean")).reset_index()
    pr_type = (d.groupby("type").agg(n=("besked", "size"), andel=("besked", "mean")).reset_index()
               .query("n >= @min_n").sort_values(["andel", "type"], ascending=[False, True], kind="mergesort"))
    return {
        "fra": SKABELON_START.strftime("%Y-%m-%d"), "n": int(len(d)),
        "andel_besked": round(d["besked"].mean(), 3), "andel_tom": round(d["tom_note"].mean(), 3),
        "pr_md": [{"md": r.md, "n": int(r.n), "b": int(r.besked)} for r in pr_md.itertuples()],
        "pr_gruppe": [{"g": r.gruppe, "n": int(r.n), "a": round(r.andel, 3)} for r in pr_gruppe.itertuples() if r.n >= min_n],
        "pr_type": [{"t": r.type, "n": int(r.n), "a": round(r.andel, 3)} for r in pr_type.itertuples()],
    }


# 2 ------------------------------------------------------------------ egne udsendelser
def udsendelser(df):
    d = df.copy()
    iso = d["dato"].dt.isocalendar()
    d["uge"] = iso["year"].astype(str) + "-U" + iso["week"].astype(str).str.zfill(2)
    d["kontingent"] = d["type"] == "Medlemskab og kontingent"
    d["kursus"] = d["type"].str.startswith("Kurser/arrangementer")
    uger = d.groupby("uge").agg(n=("type", "size"), kont=("kontingent", "sum"), kursus=("kursus", "sum"),
                                ryk=("rykker", "sum"), start=("dato", "min")).reset_index()
    md = d.groupby("md").agg(n=("type", "size"), kont=("kontingent", "sum"), ryk=("rykker", "sum")).reset_index()
    return {
        "uger": [{"uge": r.uge, "start": r.start.strftime("%Y-%m-%d"), "n": int(r.n), "kont": int(r.kont),
                  "kursus": int(r.kursus), "ryk": int(r.ryk)} for r in uger.itertuples()],
        "md": [{"md": r.md, "n": int(r.n), "kont": int(r.kont), "ryk": int(r.ryk)} for r in md.itertuples()],
        "ryk_i_alt": int(d["rykker"].sum()),
        "ryk_andel_kontingent": round(d.loc[d["rykker"], "kontingent"].mean(), 3) if d["rykker"].any() else None,
    }


# 3 ------------------------------------------------------------------ årshjulet
def aarshjul(df):
    d = df.copy()
    d["m"] = d["dato"].dt.month
    top = d["type"].value_counts().head(14).index
    tot_m = d.groupby("m").size()
    andel = d[d["type"].isin(top)].groupby(["type", "m"]).size().unstack(fill_value=0).div(tot_m, axis=1)
    # indeks: emnets andel i måneden / emnets gennemsnitlige andel. 100 = normalt.
    idx = andel.div(andel.mean(axis=1), axis=0) * 100
    return {"maaneder": [int(m) for m in idx.columns],
            "emner": [{"t": t, "n": int((d["type"] == t).sum()),
                       "idx": [round(float(idx.loc[t, m])) for m in idx.columns],
                       "top": int(idx.loc[t].idxmax())} for t in top],
            "md_n": {int(m): int(n) for m, n in tot_m.items()}}


# 4 ------------------------------------------------------------------ sager der kommer igen
def gentagne(df, min_n=MIN_N):
    d = df.sort_values(["Stamkort", "dato"]).copy()
    d["naeste"] = d.groupby("Stamkort")["dato"].shift(-1)
    d["dage"] = (d["naeste"] - d["dato"]).dt.days
    d["igen14"] = d["dage"].le(14)
    d = d[d["dato"] <= d["dato"].max() - pd.Timedelta(days=14)]   # sidste 14 dage kan ikke nå at gentage sig
    pr_type = (d.groupby("type").agg(n=("igen14", "size"), a=("igen14", "mean")).reset_index()
               .query("n >= @min_n").sort_values(["a", "type"], ascending=[False, True], kind="mergesort"))
    samme = d[d["igen14"]].copy()
    naeste_type = d.groupby("Stamkort")["type"].shift(-1)
    d["samme_emne"] = d["igen14"] & (naeste_type == d["type"])
    return {"andel": round(d["igen14"].mean(), 3), "n": int(len(d)),
            "samme_emne": round(d.loc[d["igen14"], "samme_emne"].mean(), 3),
            "median_dage": float(samme["dage"].median()) if len(samme) else None,
            "fordeling_dage": {k: int(v) for k, v in pd.cut(samme["dage"], [-1, 0, 1, 3, 7, 14],
                               labels=["Samme dag", "Dagen efter", "2–3 dage", "4–7 dage", "8–14 dage"]).value_counts().sort_index().items()},
            "pr_type": [{"t": r.type, "n": int(r.n), "a": round(r.a, 3)} for r in pr_type.itertuples()]}


# 5 ------------------------------------------------------------------ geografi
def geografi(df, min_n=25):
    felt = df["Medlemskab"].fillna("").str.split(",").str[0]
    kreds = felt.str.split("-", n=1).str[1].fillna("").str.strip()
    kreds = kreds.replace(KREDS_NAVN).replace("", pd.NA)
    seg = felt.str.split("-").str[0].str.strip()
    d = pd.DataFrame({"kreds": kreds, "selv": seg.eq("Selvstændig"), "gruppe": df["gruppe"]}).dropna(subset=["kreds"])
    agg = (d.groupby("kreds").agg(n=("selv", "size"), selv=("selv", "mean")).reset_index()
           .sort_values(["n", "kreds"], ascending=[False, True], kind="mergesort"))
    radg = d.assign(r=d["gruppe"].eq("Rådgivning/forhandling")).groupby("kreds")["r"].mean()
    return [{"k": r.kreds, "n": int(r.n), "selv": round(r.selv, 3), "raad": round(float(radg[r.kreds]), 3)}
            for r in agg.itertuples() if r.n >= min_n]


# 6 ------------------------------------------------------------------ køen og prognosen
def koe_opgivelse(kald_stier):
    rows = []
    for f in kald_stier:
        t = pd.read_excel(f, dtype=str, usecols=["Kønummer", "Starttidspunkt", "Besvarede", "Afbrudte",
                                                  "Ventetid for afbrudte", "Ventetid"])
        rows.append(t)
    k = pd.concat(rows)
    k["ts"] = pd.to_datetime(k["Starttidspunkt"], format="%d-%m-%Y %H:%M:%S")
    k = k[k["ts"] >= START]
    def s(v):
        if not isinstance(v, str) or ":" not in v:
            return None
        h, m, sec = (int(x) for x in v.split(":"))
        return h * 3600 + m * 60 + sec
    k["afb"] = k["Afbrudte"].eq("1")
    k["afb_s"] = k["Ventetid for afbrudte"].map(s)
    k["time"] = k["ts"].dt.hour
    k = k[k["time"].between(8, 16)]
    pr_time = k.groupby("time").agg(n=("afb", "size"), afb=("afb", "sum"),
                                    med=("afb_s", "median")).reset_index()
    a = k[k["afb"] & k["afb_s"].notna()]
    spand = pd.cut(a["afb_s"], [-1, 30, 60, 120, 300, 10**6],
                   labels=["Under 30 sek.", "30–60 sek.", "1–2 min.", "2–5 min.", "Over 5 min."]).value_counts().sort_index()
    return {"pr_time": [{"t": int(r.time), "n": int(r.n), "afb": int(r.afb), "med": None if pd.isna(r.med) else round(r.med)}
                        for r in pr_time.itertuples()],
            "spand": {str(k_): int(v) for k_, v in spand.items()},
            "afb_i_alt": int(k["afb"].sum()), "kald_i_alt": int(len(k))}


def prognose(valeur_dage, i_dag):
    d = pd.DataFrame(valeur_dage)
    d["d"] = pd.to_datetime(d["d"])
    d = d[d["d"].dt.dayofweek.isin([0, 1, 3, 4])]
    d["m"], d["wd"], d["y"] = d["d"].dt.month, d["d"].dt.dayofweek, d["d"].dt.year
    # vækst: de seneste tre hele måneder mod samme måneder året før
    sidste = pd.Timestamp(i_dag).to_period("M") - 1
    mdr = [sidste - i for i in range(3)]
    nu = sum(d[d["d"].dt.to_period("M") == m]["i"].sum() for m in mdr)
    foer = sum(d[d["d"].dt.to_period("M") == (m - 12)]["i"].sum() for m in mdr)
    vaekst = nu / foer if foer else 1.0
    profil = d.groupby(["m", "wd"])["i"].mean()
    wd_snit = d.groupby("wd")["i"].mean()
    ud, dag = [], pd.Timestamp(i_dag) + pd.Timedelta(days=1)
    while len(ud) < 20:
        if dag.dayofweek in (0, 1, 3, 4) and not (dag.month == 12 and dag.day >= 22):
            base = profil.get((dag.month, dag.dayofweek), wd_snit[dag.dayofweek])
            ud.append({"d": dag.strftime("%Y-%m-%d"), "wd": UGEDAGE[dag.dayofweek], "f": round(float(base) * vaekst)})
        dag += pd.Timedelta(days=1)
    mat = d.groupby(["m", "wd"])["i"].mean().unstack()
    return {"vaekst": round(vaekst, 3), "vaekst_mdr": [str(m) for m in reversed(mdr)], "dage": ud,
            "profil": {int(m): {UGEDAGE[int(w)]: round(float(v)) for w, v in r.items() if pd.notna(v)} for m, r in mat.iterrows()}}


def mc_dele(df, min_n=MIN_N):
    """De dele, der kan beregnes på et udsnit af Membercare (en afdeling)."""
    if not len(df):
        return None
    return {"loeser": loeser_selv(df, min_n), "udsendelser": udsendelser(df), "aarshjul": aarshjul(df),
            "gentagne": gentagne(df, min_n), "geografi": geografi(df, max(10, min_n // 4))}


def build(mc_sti, kald_stier, valeur_dage, i_dag, df=None):
    df = _membercare(mc_sti) if df is None else df
    return {"loeser": loeser_selv(df), "udsendelser": udsendelser(df), "aarshjul": aarshjul(df),
            "gentagne": gentagne(df), "geografi": geografi(df), "koe": koe_opgivelse(kald_stier),
            "prognose": prognose(valeur_dage, i_dag)}

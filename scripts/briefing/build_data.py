"""
Byg datagrundlaget til Valeur-dashboardet direkte fra mappen "Thomas data".

Kør:  python -X utf8 scripts/briefing/build_data.py [kildemappe]
      kildemappe er standard data/raw/thomas (mappen "Thomas data", uden for versionsstyring)
Skriver: docs/briefing.json  (kun aggregerede tal - ingen navne, numre eller fritekst)

Kilder (alle relative til kildemappen):
  Valeur Insights/                       Valeurs frontlinje (opkald pr. dag, time, ugedag, medlem)
  Alle opkald opdelt pr. måned (linjer)/  Sekretariatets køer pr. måned (IP Nordic)
  Alle opkald opdelt pr indkomne opkald/  Hvert enkelt kald i sekretariatets køer (IP Nordic)
  Membercare/Registreringer i servicefanen.xlsx   Hvad og hvem (CRM)
  Forskellige datagennemgange/            Benchmark, mails, Valeurs månedsopgørelse
  Gamle Valeur rapporter/                 Efterbehandling pr. måned m.m.
"""
import glob, json, re, sys, warnings, datetime as dt
from pathlib import Path
from collections import Counter, defaultdict
import pandas as pd
import openpyxl

warnings.filterwarnings("ignore")
REPO = Path(__file__).resolve().parents[2]
ROOT = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else REPO / "data" / "raw" / "thomas"
OUT = REPO / "docs" / "briefing.json"
START = pd.Timestamp("2025-05-05")          # Valeur gik i luften
MIN_CELLE = 5                               # ingen tal under 5 personer i krydstabeller

KOE_NAVN = {"1510": "Ansættelse (rådgivning)", "1515": "Selvstændig", "15122": "Kursus/arrangement",
            "15123": "Medlemsservice", "1520": "zExpense", "1525": "Debitor", "1500": "Velkommen/Øvrige"}
EN_MD = {m: i for i, m in enumerate(["January", "February", "March", "April", "May", "June", "July",
                                     "August", "September", "October", "November", "December"], 1)}


def sek(v):
    if v is None or (isinstance(v, float) and pd.isna(v)) or v in ("", "-"):
        return None
    if isinstance(v, dt.timedelta):
        return v.total_seconds()
    if isinstance(v, dt.time):
        return v.hour * 3600 + v.minute * 60 + v.second
    p = str(v).strip().split(":")
    try:
        p = [float(x) for x in p]
    except ValueError:
        return None
    while len(p) < 3:
        p.insert(0, 0)
    return p[0] * 3600 + p[1] * 60 + p[2]


def num(v):
    if v is None or v == "" or (isinstance(v, float) and pd.isna(v)):
        return None
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        return None


def r0(v):
    return None if v is None or pd.isna(v) else round(float(v))


def mkey(ts):
    return f"{ts.year}-{ts.month:02d}"


# ---------------------------------------------------------------- 1. Valeur frontlinje
def valeur():
    d = pd.read_excel(ROOT / "Valeur Insights/Opkald per dag - ud fra valgt filter på kunderapport (1).xlsx",
                      header=None)
    d.columns = ["dato", "besvarede", "ialt"]
    d = d[d["dato"].map(lambda x: isinstance(x, (dt.datetime, pd.Timestamp)))]
    d["dato"] = pd.to_datetime(d["dato"])
    d["besvarede"] = pd.to_numeric(d["besvarede"], errors="coerce")
    d["ialt"] = pd.to_numeric(d["ialt"], errors="coerce")
    d = d[d["dato"] >= START].fillna(0)
    d = d[d["ialt"] > 0]
    d["uge"] = d["dato"].dt.isocalendar().week
    d["aar"] = d["dato"].dt.isocalendar().year
    d["md"] = d["dato"].dt.to_period("M").astype(str)

    uger = d.groupby(["aar", "uge"]).agg(start=("dato", "min"), ialt=("ialt", "sum"),
                                         besvarede=("besvarede", "sum"), dage=("dato", "nunique")).reset_index()
    md = d.groupby("md").agg(ialt=("ialt", "sum"), besvarede=("besvarede", "sum"),
                             dage=("dato", "nunique")).reset_index()

    def fordeling(fil, n):
        t = pd.read_excel(ROOT / "Valeur Insights" / fil, header=None)
        t.columns = ["k", "v"]
        t = t[pd.to_numeric(t["v"], errors="coerce").notna()]
        return {str(k): float(v) for k, v in zip(t["k"], t["v"])}

    return {
        "dage": [{"d": r.dato.strftime("%Y-%m-%d"), "i": int(r.ialt), "b": int(r.besvarede)} for r in d.itertuples()],
        "uger": [{"uge": f"{r.aar}-U{int(r.uge):02d}", "start": r.start.strftime("%Y-%m-%d"), "i": int(r.ialt),
                  "b": int(r.besvarede), "dage": int(r.dage)} for r in uger.itertuples()],
        "maaneder": [{"md": r.md, "i": int(r.ialt), "b": int(r.besvarede), "dage": int(r.dage)}
                     for r in md.itertuples()],
        "total": int(d["ialt"].sum()), "besvarede": int(d["besvarede"].sum()), "aabningsdage": int(d["dato"].nunique()),
        "foerste": d["dato"].min().strftime("%Y-%m-%d"), "sidste": d["dato"].max().strftime("%Y-%m-%d"),
        "time": fordeling("Timefordelingen - ud fra valgt filter på kunderapport (1).xlsx", 1),
        "time_afbrudt": fordeling("Timefordelingen for kald afbrudt inden kø (1).xlsx", 1),
        "ugedag": fordeling("Ugedages andel af kald - ud fra valgt filter på kunderapport (1).xlsx", 1),
        "medlem": fordeling("Medlem_Ikke_medlem.xlsx", 1),
    }


# ---------------------------------------------------------------- 2. Sekretariatets køer pr. måned
def koeer_maaned():
    ud = {}
    for f in glob.glob(str(ROOT / "Alle opkald opdelt pr. måned (linjer)" / "*.xlsx")):
        t = pd.read_excel(f, dtype=str)
        for r in t.itertuples(index=False):
            navn, aar = str(r[0]).split()
            koe = str(r[1])
            key = f"{aar}-{EN_MD[navn]:02d}"
            ud.setdefault(koe, {})[key] = {
                "opkald": int(num(r[3]) or 0), "besvarede": int(num(r[5]) or 0),
                "afbrudte": int(num(r[11]) or 0),
                "samtale": sek(r[6]), "vent": sek(r[7]), "maxvent": sek(r[8]),
            }
    return ud


# ---------------------------------------------------------------- 3. Hvert kald i sekretariatets køer
def koeer_kald():
    rows = []
    for f in glob.glob(str(ROOT / "Alle opkald opdelt pr indkomne opkald (linjer)" / "*.xlsx")):
        t = pd.read_excel(f, dtype=str)
        t = t[["Nummer", "Kønummer", "Starttidspunkt", "Besvarede", "Samtaletid", "Ventetid", "Afbrudte",
               "Ventetid for afbrudte"]]
        rows.append(t)
    df = pd.concat(rows)
    df["ts"] = pd.to_datetime(df["Starttidspunkt"], format="%d-%m-%Y %H:%M:%S")
    df["besv"] = df["Besvarede"].fillna("0").astype(int)
    df["vent_s"] = df["Ventetid"].map(sek)
    df["afb_s"] = df["Ventetid for afbrudte"].map(sek)
    df["tale_s"] = df["Samtaletid"].map(sek)
    df["efter"] = df["ts"] >= START
    df["time"] = df["ts"].dt.hour
    df["ugedag"] = df["ts"].dt.dayofweek

    # sammenlign samme køer i 8 mdr. før (sep 24 - apr 25) og de seneste 12 hele mdr. (nov 25 - okt 26)
    foer = df[(df["ts"] >= "2024-09-01") & (df["ts"] < "2025-05-01")]
    efter = df[(df["ts"] >= "2025-11-01") & (df["ts"] < "2026-11-01")]

    def maal(x):
        dage = x["ts"].dt.normalize().nunique()
        b = x[x["besv"] == 1]
        return {"kald": len(x), "pr_dag": round(len(x) / max(dage, 1), 1), "besvarede": int(len(b)),
                "besv_pr_dag": round(len(b) / max(dage, 1), 1),
                "svarpct": round(len(b) / max(len(x), 1), 3),
                "vent": r0(b["vent_s"].mean()), "tale": r0(b["tale_s"].mean()),
                "vent_afbrudt": r0(x["afb_s"].mean()),
                "andel_vent_over_2": round((b["vent_s"] > 120).mean(), 3) if len(b) else None,
                "dage": int(dage)}

    foer_efter = {}
    for koe in sorted(df["Kønummer"].unique()):
        foer_efter[koe] = {"foer": maal(foer[foer["Kønummer"] == koe]), "efter": maal(efter[efter["Kønummer"] == koe])}
    foer_efter["alle"] = {"foer": maal(foer), "efter": maal(efter)}

    # gentagne opkald i sekretariatets køer efter Valeur (numre bruges kun til at tælle, kasseres her)
    e = df[df["efter"]]
    pr_nummer = e.groupby("Nummer").size()
    gentag = {"numre": int(len(pr_nummer)), "kald": int(len(e)),
              "fordeling": {str(k): int(v) for k, v in pr_nummer.clip(upper=6).value_counts().sort_index().items()}}

    # time- og ugedagsprofil efter Valeur, besvaret vs. ikke besvaret
    tid = e.groupby(["time"]).agg(kald=("besv", "size"), besv=("besv", "sum")).reset_index()
    ugd = e.groupby(["ugedag"]).agg(kald=("besv", "size"), besv=("besv", "sum")).reset_index()

    # månedlig svarprocent pr. kø fra kaldsdata
    df["md"] = df["ts"].dt.to_period("M").astype(str)
    mdr = df.groupby(["Kønummer", "md"]).agg(kald=("besv", "size"), besv=("besv", "sum"),
                                              vent=("vent_s", "mean")).reset_index()
    pr_md = defaultdict(dict)
    for r in mdr.itertuples():
        pr_md[r.Kønummer][r.md] = {"kald": int(r.kald), "besv": int(r.besv), "vent": r0(r.vent)}

    return {"foer_efter": foer_efter, "gentag": gentag,
            "time": [{"t": int(r.time), "k": int(r.kald), "b": int(r.besv)} for r in tid.itertuples()],
            "ugedag": [{"d": int(r.ugedag), "k": int(r.kald), "b": int(r.besv)} for r in ugd.itertuples()],
            "pr_md": pr_md, "periode": [df["ts"].min().strftime("%Y-%m-%d"), df["ts"].max().strftime("%Y-%m-%d")]}


# ---------------------------------------------------------------- 4. Membercare / servicefanen
def membercare():
    df = pd.read_excel(ROOT / "Membercare/Registreringer i servicefanen.xlsx", dtype=str)
    df["dato"] = pd.to_datetime(df["Service dato"], errors="coerce")
    df["oprettet"] = pd.to_datetime(df["Oprettet"], errors="coerce")
    df = df[df["dato"].notna()]
    # Valeur-agenter: dem, der skriver "Valeur" i deres noter
    bem = df["Bemærkning"].fillna("")
    valeur_hits = df.assign(v=bem.str.contains("Valeur", case=False)).groupby("Oprettet af")["v"].agg(["sum", "size"])
    valeur_agenter = set(valeur_hits[(valeur_hits["sum"] >= 3)].index)
    df["af_valeur"] = df["Oprettet af"].isin(valeur_agenter)

    efter = df[df["dato"] >= START].copy()
    gruppe = efter["Servicegruppe"].fillna("Ikke angivet").str.strip()
    typ = efter["Servicetype"].fillna("Ikke angivet").str.strip()
    efter["gruppe"], efter["type"] = gruppe, typ

    # medlemstype = første led af medlemskabsfeltet. Feltet viser medlemskabet på
    # udtrækstidspunktet, så udmeldte og medlemskaber under oprettelse står tomme.
    TOM = "Medlemskab ikke angivet"

    def seg(s):
        if not isinstance(s, str) or not s.strip():
            return TOM
        s = s.strip().split(",")[0]
        s = s.split("-")[0].strip()
        s = re.sub(r"\s+", " ", s)
        return s.rstrip(".") or TOM
    efter["seg"] = efter["Medlemskab"].map(seg)

    tom = efter[efter["seg"] == TOM]
    tomt_felt = {"registreringer": int(len(tom)), "personer": int(tom["Stamkort"].nunique()),
                 "andel_kontingent": round((tom["type"] == "Medlemskab og kontingent").mean(), 3),
                 "andel_kontingent_alle": round((efter["type"] == "Medlemskab og kontingent").mean(), 3)}
    grp = efter["gruppe"].value_counts()
    typer = (efter.groupby(["gruppe", "type"]).size().reset_index(name="n")
             .sort_values("n", ascending=False))
    typer = typer[typer["n"] >= MIN_CELLE]
    segs = efter["seg"].value_counts()
    segs_ud = {k: int(v) for k, v in segs.items() if v >= MIN_CELLE}
    segs_ud["Øvrige grupper"] = int(segs[segs < MIN_CELLE].sum())

    top_seg = list(segs.index[:6])
    kryds = efter[efter["seg"].isin(top_seg)].groupby(["seg", "gruppe"]).size().unstack(fill_value=0)

    efter["md"] = efter["dato"].dt.to_period("M").astype(str)
    md = efter.groupby(["md", "gruppe"]).size().unstack(fill_value=0)

    # gentagne henvendelser pr. stamkort (bruges kun til at tælle)
    pr_pers = efter.groupby("Stamkort").size()
    gentag = {str(k): int(v) for k, v in pr_pers.clip(upper=6).value_counts().sort_index().items()}

    pr_md_valeur = efter.groupby("md")["af_valeur"].agg(["sum", "size"]).reset_index()

    return {
        "periode": [efter["dato"].min().strftime("%Y-%m-%d"), efter["dato"].max().strftime("%Y-%m-%d")],
        "total_alle": int(len(df)), "total_efter": int(len(efter)),
        "foer_valeur": {"foerste": df["dato"].min().strftime("%Y-%m-%d"),
                        "antal": int((df["dato"] < START).sum())},
        "unikke": int(efter["Stamkort"].nunique()),
        "tomt_felt": tomt_felt,
        "gentag": gentag,
        "andel_valeur": round(efter["af_valeur"].mean(), 3),
        "antal_valeur_agenter": len(valeur_agenter),
        "grupper": {k: int(v) for k, v in grp.items()},
        "typer": [{"g": r.gruppe, "t": r.type, "n": int(r.n)} for r in typer.itertuples()],
        "segmenter": segs_ud,
        "kryds": {s: {g: int(n) for g, n in kryds.loc[s].items()} for s in kryds.index},
        "pr_md": {m: {g: int(n) for g, n in md.loc[m].items()} for m in md.index},
        "valeur_pr_md": [{"md": r.md, "valeur": int(r.sum), "alle": int(r.size)} for r in pr_md_valeur.itertuples()],
        "_agenter": sorted(valeur_agenter),   # fjernes før udskrivning
    }


# ---------------------------------------------------------------- 5. Valeurs benchmark og månedsopgørelse
def tid_mmss(v):
    """Benchmark-arket gemmer mm:ss skiftevis som (t,m) og (0,m,s)."""
    if isinstance(v, dt.time):
        if v.second == 0:
            return v.hour * 60 + v.minute
        return (v.hour * 60 + v.minute) * 60 + v.second
    return None


def benchmark():
    wb = openpyxl.load_workbook(ROOT / "Forskellige datagennemgange/Benchmark DP 2026 (1) (002).xlsx", data_only=True)
    ws = wb["Benchmark"]
    perioder, cur = [], None
    for r in ws.iter_rows(values_only=True):
        r = list(r) + [None] * 6
        if r[0] == "Benchmarks":
            continue
        if isinstance(r[0], str) and r[0].strip():
            cur = {"periode": r[0].strip()}
            perioder.append(cur)
        if cur is None or not isinstance(r[1], str):
            continue
        m = r[1].strip()
        v = r[3]
        if isinstance(v, str):
            continue
        if m in ("SLA svarrate", "FCR", "Registrering MC"):
            if isinstance(v, (int, float)):
                cur[m] = float(v)
        elif m in ("Gns Ventetid", "Max taletid", "Efterbehandling", "AHT"):
            s = tid_mmss(v)
            if s is not None:
                cur[m] = s
        elif m == "Antal opkald" and isinstance(v, (int, float)):
            cur[m] = int(v)
    perioder = [p for p in perioder if len(p) > 1]
    maal = {"SLA svarrate": 0.9, "Gns Ventetid": 90, "Max taletid": 720, "FCR": 0.75,
            "Efterbehandling": 90, "Registrering MC": 0.9}
    return {"perioder": perioder, "maal": maal}


def valeur_maaned():
    wb = openpyxl.load_workbook(ROOT / "Forskellige datagennemgange/Datagennemgang_21_09_26.xlsx", data_only=True)
    ws = wb["Opkald - månedsopdelt"]
    ud = []
    for r in list(ws.iter_rows(values_only=True))[1:]:
        if not isinstance(r[0], dt.datetime):
            continue
        ud.append({"md": r[0].strftime("%Y-%m"), "tilbudt": r[1], "besvaret": r[2],
                   "vent": sek(r[3]), "tale": sek(r[4]), "efter": sek(r[5]),
                   "tale_timer": round(sek(r[6]) / 3600), "efter_timer": round(sek(r[7]) / 3600)})
    # efterbehandling maj 25 - (Valeur_statistik_11_06)
    wb = openpyxl.load_workbook(ROOT / "Gamle Valeur rapporter/Valeur_statistik_11_06.xlsx", data_only=True)
    eb = []
    for r in list(wb["Efterbehandling"].iter_rows(values_only=True))[2:]:
        if isinstance(r[0], dt.datetime) and isinstance(r[1], dt.time):
            eb.append({"md": r[0].strftime("%Y-%m"), "s": r[1].hour * 60 + r[1].minute,
                       "note": r[2] if len(r) > 2 and r[2] else None})
    return {"maaneder": ud, "efterbehandling": eb}


# ---------------------------------------------------------------- 6. Mails pr. uge
def mails():
    wb = openpyxl.load_workbook(ROOT / "Forskellige datagennemgange/2026  Samlet Oversigt mails Dansk Psykologforening ugeopdelt (002).xlsx", data_only=True)
    ud = {}
    for ws, aar in ((wb["2025"], 2025), (wb["2026"], 2026)):
        uger = []
        for r in ws.iter_rows(values_only=True):
            celler = [c for c in r if c is not None and not (isinstance(c, str) and not c.strip())]
            hoved = [int(re.findall(r"\d+", c)[0]) for c in celler if isinstance(c, str) and re.match(r"^\s*Uge\s*\d+", c)]
            if hoved:
                uger = hoved
                continue
            if celler and isinstance(celler[0], str) and celler[0].strip() == "Total" and uger:
                tal = [c for c in celler if isinstance(c, (int, float))]
                for u, v in zip(uger, tal):
                    if v:
                        ud[f"{aar}-U{u:02d}"] = int(v)
                uger = []
    return dict(sorted(ud.items()))


def koe_foer_efter(md):
    """Samme køer, 8 mdr. før (sep 24-apr 25) mod de seneste 12 hele mdr. (nov 25-okt 26). Kilde: kø-rapporterne."""
    foer = [f"2024-{m:02d}" for m in range(9, 13)] + [f"2025-{m:02d}" for m in range(1, 5)]
    efter = [f"2025-{m:02d}" for m in (11, 12)] + [f"2026-{m:02d}" for m in range(1, 11)]
    ud = {}
    for koe, rk in md.items():
        if koe == "1500":
            continue
        def agg(mdr):
            r = [rk[m] for m in mdr if m in rk]
            o = sum(x["opkald"] for x in r); b = sum(x["besvarede"] for x in r)
            vw = sum((x["vent"] or 0) * x["besvarede"] for x in r)
            tw = sum((x["samtale"] or 0) * x["besvarede"] for x in r)
            return {"opkald_md": round(o / len(mdr), 1), "besv_md": round(b / len(mdr), 1),
                    "svarpct": round(b / o, 3) if o else None,
                    "vent": round(vw / b) if b else None, "tale": round(tw / b) if b else None,
                    "tale_timer_md": round(tw / 3600 / len(mdr), 1)}
        ud[koe] = {"foer": agg(foer), "efter": agg(efter)}
    return ud


if __name__ == "__main__":
    data = {"genereret": dt.date.today().isoformat(), "valeur_start": START.strftime("%Y-%m-%d")}
    data["valeur"] = valeur()
    data["koeer_md"] = koeer_maaned()
    data["koeer"] = koeer_kald()
    data["mc"] = membercare()
    data["benchmark"] = benchmark()
    data["valeur_md"] = valeur_maaned()
    data["mails"] = mails()
    data["koe_navn"] = KOE_NAVN
    data["koe_fe"] = koe_foer_efter(data["koeer_md"])
    data["koeer"].pop("gentag")

    agenter = data["mc"].pop("_agenter")
    txt = json.dumps(data, ensure_ascii=False, default=str)
    # sikkerhedsnet: ingen telefonnumre, mails eller cpr-lignende tal i output
    assert not re.search(r"\+45\d{8}|@[a-z0-9-]+\.[a-z]{2,}|\b\d{6}-\d{4}\b", txt), "Personhenførbart indhold fundet!"
    OUT.write_text(txt, encoding="utf-8")
    print("Valeur-agenter fundet:", len(agenter))
    print("skrev", OUT, len(txt), "bytes")

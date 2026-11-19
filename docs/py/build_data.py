"""
Byg datagrundlaget til briefing-dashboardet.

Kør lokalt:   python -X utf8 scripts/briefing/build_data.py "<mappe med eksporter>"
              (uden mappe: data/raw/thomas i repoet, eller mappen over scriptet)
I browseren:  docs/opdater.html kører den samme byg() via Pyodide.

Filerne findes ud fra deres indhold af kilder.py, så filnavne og mappestruktur er
ligegyldige. Mangler en kilde, genbruges de tal, der sidst blev publiceret (hvis
de findes), og det står i data["kilder"].

Output er kun aggregerede tal. Før noget skrives, tjekkes outputtet for
telefonnumre, mailadresser, cpr-lignende numre og navne fra Membercare.
"""
import json, re, sys, warnings, datetime as dt
from pathlib import Path
from collections import defaultdict
import pandas as pd
import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
import kilder as K

warnings.filterwarnings("ignore")
START = pd.Timestamp("2025-05-05")          # Valeur gik i luften
FOER = ("2024-09", "2025-04")               # sammenligningsperioden før Valeur (samme køer som i dag)
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


def maaneder(fra, til):
    """Liste af 'ÅÅÅÅ-MM' fra og med fra til og med til."""
    return [str(p) for p in pd.period_range(fra, til, freq="M")]


def efter_periode(sidste_dato):
    """De seneste 12 hele måneder før den måned, data slutter i."""
    sidste = pd.Timestamp(sidste_dato)
    slut = (sidste + pd.Timedelta(days=1)).to_period("M") - 1     # hel måned, hvis data går til månedens sidste dag
    return (str(slut - 11), str(slut))


# ---------------------------------------------------------------- 1. Valeur frontlinje
def _fordeling(sti):
    t = pd.read_excel(sti, header=None)
    t = t.iloc[:, :2]
    t.columns = ["k", "v"]
    t = t[pd.to_numeric(t["v"], errors="coerce").notna()]
    return {str(k): float(v) for k, v in zip(t["k"], t["v"])}


def valeur(sti_dag, fordelinger, tidligere):
    d = pd.read_excel(sti_dag, header=None).iloc[:, :3]
    d.columns = ["dato", "besvarede", "ialt"]
    d = d[d["dato"].map(lambda x: isinstance(x, (dt.datetime, pd.Timestamp)))]
    d["dato"] = pd.to_datetime(d["dato"])
    d["besvarede"] = pd.to_numeric(d["besvarede"], errors="coerce")
    d["ialt"] = pd.to_numeric(d["ialt"], errors="coerce")
    d = d[d["dato"] >= START].fillna(0)
    d = d[d["ialt"] > 0]
    iso = d["dato"].dt.isocalendar()
    d["uge"], d["aar"] = iso.week, iso.year
    d["md"] = d["dato"].dt.to_period("M").astype(str)
    uger = d.groupby(["aar", "uge"]).agg(start=("dato", "min"), ialt=("ialt", "sum"),
                                         besvarede=("besvarede", "sum"), dage=("dato", "nunique")).reset_index()
    md = d.groupby("md").agg(ialt=("ialt", "sum"), besvarede=("besvarede", "sum"),
                             dage=("dato", "nunique")).reset_index()
    ud = {
        "dage": [{"d": r.dato.strftime("%Y-%m-%d"), "i": int(r.ialt), "b": int(r.besvarede)} for r in d.itertuples()],
        "uger": [{"uge": f"{r.aar}-U{int(r.uge):02d}", "start": r.start.strftime("%Y-%m-%d"), "i": int(r.ialt),
                  "b": int(r.besvarede), "dage": int(r.dage)} for r in uger.itertuples()],
        "maaneder": [{"md": r.md, "i": int(r.ialt), "b": int(r.besvarede), "dage": int(r.dage)} for r in md.itertuples()],
        "total": int(d["ialt"].sum()), "besvarede": int(d["besvarede"].sum()), "aabningsdage": int(d["dato"].nunique()),
        "foerste": d["dato"].min().strftime("%Y-%m-%d"), "sidste": d["dato"].max().strftime("%Y-%m-%d"),
    }
    for felt, rolle in (("time", "valeur_time"), ("time_afbrudt", "valeur_time_afbrudt"),
                        ("ugedag", "valeur_ugedag"), ("medlem", "valeur_medlem")):
        if fordelinger.get(rolle):
            ud[felt] = _fordeling(fordelinger[rolle])
        else:
            ud[felt] = (tidligere or {}).get("valeur", {}).get(felt, {})
    return ud


# ---------------------------------------------------------------- 2. Sekretariatets køer pr. måned
def koeer_maaned(stier):
    ud = {}
    for f in stier.values():
        t = pd.read_excel(f, dtype=str)
        for r in t.itertuples(index=False):
            dele = str(r[0]).split()
            if len(dele) != 2 or dele[0] not in EN_MD:
                continue
            key = f"{dele[1]}-{EN_MD[dele[0]]:02d}"
            ud.setdefault(str(r[1]), {})[key] = {
                "opkald": int(num(r[3]) or 0), "besvarede": int(num(r[5]) or 0), "afbrudte": int(num(r[11]) or 0),
                "samtale": sek(r[6]), "vent": sek(r[7]), "maxvent": sek(r[8]),
            }
    return ud


# ---------------------------------------------------------------- 3. Hvert kald i sekretariatets køer
def koeer_kald(stier, efter):
    rows = []
    for f in stier.values():
        t = pd.read_excel(f, dtype=str)
        rows.append(t[["Kønummer", "Starttidspunkt", "Besvarede", "Samtaletid", "Ventetid", "Afbrudte",
                       "Ventetid for afbrudte"]])
    df = pd.concat(rows)
    df["ts"] = pd.to_datetime(df["Starttidspunkt"], format="%d-%m-%Y %H:%M:%S", errors="coerce")
    df = df[df["ts"].notna()]
    df["besv"] = df["Besvarede"].fillna("0").astype(int)
    df["vent_s"] = df["Ventetid"].map(sek)
    df["afb_s"] = df["Ventetid for afbrudte"].map(sek)
    df["tale_s"] = df["Samtaletid"].map(sek)
    df["time"] = df["ts"].dt.hour
    df["ugedag"] = df["ts"].dt.dayofweek
    df["md"] = df["ts"].dt.to_period("M").astype(str)

    e = df[df["ts"] >= START]
    tid = e.groupby(["time"]).agg(kald=("besv", "size"), besv=("besv", "sum")).reset_index()
    ugd = e.groupby(["ugedag"]).agg(kald=("besv", "size"), besv=("besv", "sum")).reset_index()
    mdr = df.groupby(["Kønummer", "md"]).agg(kald=("besv", "size"), besv=("besv", "sum"),
                                              vent=("vent_s", "mean")).reset_index()
    pr_md = defaultdict(dict)
    for r in mdr.itertuples():
        pr_md[r.Kønummer][r.md] = {"kald": int(r.kald), "besv": int(r.besv), "vent": r0(r.vent)}
    return {"time": [{"t": int(r.time), "k": int(r.kald), "b": int(r.besv)} for r in tid.itertuples()],
            "ugedag": [{"d": int(r.ugedag), "k": int(r.kald), "b": int(r.besv)} for r in ugd.itertuples()],
            "pr_md": pr_md, "periode": [df["ts"].min().strftime("%Y-%m-%d"), df["ts"].max().strftime("%Y-%m-%d")]}


def koe_foer_efter(md, efter):
    """Samme køer: FOER mod de seneste 12 hele måneder. Kilde: kø-rapporterne pr. måned."""
    foer, eft = maaneder(*FOER), maaneder(*efter)
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
        ud[koe] = {"foer": agg(foer), "efter": agg(eft)}
    return ud


# ---------------------------------------------------------------- 4. Membercare / servicefanen
def membercare(sti):
    df = pd.read_excel(sti, dtype=str)
    df["dato"] = pd.to_datetime(df["Service dato"], errors="coerce")
    df = df[df["dato"].notna()]
    bem = df["Bemærkning"].fillna("")
    hits = df.assign(v=bem.str.contains("Valeur", case=False)).groupby("Oprettet af")["v"].agg(["sum", "size"])
    agenter = set(hits[hits["sum"] >= 3].index)
    df["af_valeur"] = df["Oprettet af"].isin(agenter)

    efter = df[df["dato"] >= START].copy()
    efter["gruppe"] = efter["Servicegruppe"].fillna("Ikke angivet").str.strip()
    efter["type"] = efter["Servicetype"].fillna("Ikke angivet").str.strip()
    TOM = "Medlemskab ikke angivet"

    def seg(s):
        if not isinstance(s, str) or not s.strip():
            return TOM
        s = s.strip().split(",")[0].split("-")[0].strip()
        return re.sub(r"\s+", " ", s).rstrip(".") or TOM
    efter["seg"] = efter["Medlemskab"].map(seg)

    tom = efter[efter["seg"] == TOM]
    grp = efter["gruppe"].value_counts()
    typer = efter.groupby(["gruppe", "type"]).size().reset_index(name="n").sort_values("n", ascending=False)
    typer = typer[typer["n"] >= MIN_CELLE]
    segs = efter["seg"].value_counts()
    segs_ud = {k: int(v) for k, v in segs.items() if v >= MIN_CELLE}
    segs_ud["Øvrige grupper"] = int(segs[segs < MIN_CELLE].sum())
    kryds = efter[efter["seg"].isin(list(segs.index[:6]))].groupby(["seg", "gruppe"]).size().unstack(fill_value=0)
    efter["md"] = efter["dato"].dt.to_period("M").astype(str)
    md = efter.groupby(["md", "gruppe"]).size().unstack(fill_value=0)
    pr_pers = efter.groupby("Stamkort").size()
    pr_md_valeur = efter.groupby("md")["af_valeur"].agg(["sum", "size"]).reset_index()

    navne = set()
    for kol in ("Medlem", "Oprettet af", "Ændret af"):
        if kol in df:
            navne |= {n.strip() for n in df[kol].dropna() if len(n.strip()) > 5 and len(n.strip().split()) > 1}
    return {
        "periode": [efter["dato"].min().strftime("%Y-%m-%d"), efter["dato"].max().strftime("%Y-%m-%d")],
        "total_alle": int(len(df)), "total_efter": int(len(efter)),
        "foer_valeur": {"foerste": df["dato"].min().strftime("%Y-%m-%d"), "antal": int((df["dato"] < START).sum())},
        "unikke": int(efter["Stamkort"].nunique()),
        "tomt_felt": {"registreringer": int(len(tom)), "personer": int(tom["Stamkort"].nunique()),
                      "andel_kontingent": round((tom["type"] == "Medlemskab og kontingent").mean(), 3),
                      "andel_kontingent_alle": round((efter["type"] == "Medlemskab og kontingent").mean(), 3)},
        "gentag": {str(k): int(v) for k, v in pr_pers.clip(upper=6).value_counts().sort_index().items()},
        "andel_valeur": round(efter["af_valeur"].mean(), 3),
        "antal_valeur_agenter": len(agenter),
        "grupper": {k: int(v) for k, v in grp.items()},
        "typer": [{"g": r.gruppe, "t": r.type, "n": int(r.n)} for r in typer.itertuples()],
        "segmenter": segs_ud,
        "kryds": {s: {g: int(n) for g, n in kryds.loc[s].items()} for s in kryds.index},
        "pr_md": {m: {g: int(n) for g, n in md.loc[m].items()} for m in md.index},
        "valeur_pr_md": [{"md": r.md, "valeur": int(r.sum), "alle": int(r.size)} for r in pr_md_valeur.itertuples()],
    }, navne


# ---------------------------------------------------------------- 5. Valeurs egne ark
def tid_mmss(v):
    """Benchmark-arket gemmer mm:ss skiftevis som (t,m) og (0,m,s)."""
    if isinstance(v, dt.time):
        if v.second == 0:
            return v.hour * 60 + v.minute
        return (v.hour * 60 + v.minute) * 60 + v.second
    return None


def benchmark(sti):
    ws = openpyxl.load_workbook(sti, data_only=True)["Benchmark"]
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
        m, v = r[1].strip(), r[3]
        if isinstance(v, str):
            continue
        if m in ("SLA svarrate", "FCR", "Registrering MC") and isinstance(v, (int, float)):
            cur[m] = float(v)
        elif m in ("Gns Ventetid", "Max taletid", "Efterbehandling", "AHT"):
            s = tid_mmss(v)
            if s is not None:
                cur[m] = s
        elif m == "Antal opkald" and isinstance(v, (int, float)):
            cur[m] = int(v)
    return {"perioder": [p for p in perioder if len(p) > 1],
            "maal": {"SLA svarrate": 0.9, "Gns Ventetid": 90, "Max taletid": 720, "FCR": 0.75,
                     "Efterbehandling": 90, "Registrering MC": 0.9}}


def valeur_maaned(sti):
    ws = openpyxl.load_workbook(sti, data_only=True)["Opkald - månedsopdelt"]
    ud = []
    for r in list(ws.iter_rows(values_only=True))[1:]:
        if not isinstance(r[0], dt.datetime):
            continue
        ud.append({"md": r[0].strftime("%Y-%m"), "tilbudt": r[1], "besvaret": r[2],
                   "vent": sek(r[3]), "tale": sek(r[4]), "efter": sek(r[5]),
                   "tale_timer": round((sek(r[6]) or 0) / 3600), "efter_timer": round((sek(r[7]) or 0) / 3600)})
    return ud


def efterbehandling(sti):
    ws = openpyxl.load_workbook(sti, data_only=True)["Efterbehandling"]
    eb = []
    for r in list(ws.iter_rows(values_only=True))[2:]:
        if isinstance(r[0], dt.datetime) and isinstance(r[1], dt.time):
            eb.append({"md": r[0].strftime("%Y-%m"), "s": r[1].hour * 60 + r[1].minute,
                       "note": r[2] if len(r) > 2 and r[2] else None})
    return eb


def mails(sti):
    wb = openpyxl.load_workbook(sti, data_only=True)
    ud = {}
    for ws in wb.worksheets:
        if not re.fullmatch(r"20\d\d", ws.title.strip()):
            continue
        aar, uger = int(ws.title.strip()), []
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


# ---------------------------------------------------------------- samling
MONSTRE = [(r"\+45\s?\d{8}", "telefonnummer"), (r"[\w.+-]+@[\w-]+\.[a-z]{2,}", "mailadresse"),
           (r"\b\d{6}-\d{4}\b", "cpr-lignende nummer")]


def tjek_privatliv(txt, navne):
    fund = []
    for m, hvad in MONSTRE:
        if re.search(m, txt):
            fund.append(hvad)
    navnefund = [n for n in navne if n in txt]
    if navnefund:
        fund.append(f"{len(navnefund)} navne fra Membercare")
    return fund


def byg(rod, tidligere=None, mtimes=None, log=print):
    """Find kilderne under rod og byg data. Returnér (data, oversigt over filer)."""
    tidligere = tidligere or {}
    log("Finder kildefiler …")
    valgt, oversigt = K.find(rod, mtimes=mtimes, log=log)
    navn = lambda s: Path(s).name
    meta = {}

    def ny(rolle, fra=None, til=None, note=None, filer=None):
        v = valgt.get(rolle)
        fl = filer or ([navn(x) for x in v.values()] if isinstance(v, dict) else [navn(v)] if v else [])
        meta[rolle] = {"status": "ny", "filer": fl, "fra": fra, "til": til, "note": note}

    def genbrug(rolle, *nøgler):
        gl = tidligere.get("kilder", {}).get(rolle)
        meta[rolle] = {**(gl or {}), "status": "genbrugt" if gl or nøgler and _hent(tidligere, nøgler) is not None
                       else "mangler"}
        return _hent(tidligere, nøgler) if nøgler else None

    data = {"genereret": dt.date.today().isoformat(), "valeur_start": START.strftime("%Y-%m-%d"),
            "koe_navn": KOE_NAVN}
    navne = set()

    # Valeur frontlinje
    if valgt.get("valeur_dag"):
        log("Læser Valeur Insights …")
        fd = {r: valgt.get(r) for r in ("valeur_time", "valeur_time_afbrudt", "valeur_ugedag", "valeur_medlem")}
        data["valeur"] = valeur(valgt["valeur_dag"], fd, tidligere)
        ny("valeur_dag", data["valeur"]["foerste"], data["valeur"]["sidste"])
        for r, sti in fd.items():
            if sti:
                filt = next((o.get("filter") for o in oversigt if o["rolle"] == r and o["brugt"]), "")
                d = re.findall(r"\d{2}-\d{2}-\d{4}", filt or "")
                ny(r, *(pd.to_datetime(x, dayfirst=True).strftime("%Y-%m-%d") for x in d[:2]) if len(d) >= 2 else (None, None),
                   note=None if d else "Hele perioden (intet datofilter)")
            else:
                genbrug(r)
    else:
        data["valeur"] = genbrug("valeur_dag", "valeur")
        for r in ("valeur_time", "valeur_time_afbrudt", "valeur_ugedag", "valeur_medlem"):
            if r not in meta:
                genbrug(r)
    if not data["valeur"]:
        raise ValueError("Opkald pr. dag fra Valeur Insights mangler, og der er ingen tidligere tal at bruge.")
    efter = efter_periode(data["valeur"]["sidste"])
    data["perioder"] = {"foer": list(FOER), "efter": list(efter)}

    # Sekretariatets køer
    if valgt.get("koe_maaned"):
        log("Læser køerne pr. måned …")
        data["koeer_md"] = koeer_maaned(valgt["koe_maaned"])
        alle = sorted({m for k in data["koeer_md"].values() for m in k})
        ny("koe_maaned", alle[0], alle[-1], note=f"{len(valgt['koe_maaned'])} køer")
    else:
        data["koeer_md"] = genbrug("koe_maaned", "koeer_md")
    data["koe_fe"] = koe_foer_efter(data["koeer_md"], efter)

    if valgt.get("koe_kald"):
        log("Læser hvert kald i køerne …")
        data["koeer"] = koeer_kald(valgt["koe_kald"], efter)
        ny("koe_kald", *data["koeer"]["periode"], note=f"{len(valgt['koe_kald'])} køer")
    else:
        data["koeer"] = genbrug("koe_kald", "koeer")

    # Membercare
    if valgt.get("membercare"):
        log("Læser Membercare …")
        data["mc"], navne = membercare(valgt["membercare"])
        ny("membercare", *data["mc"]["periode"], note=f"{data['mc']['total_efter']:,} registreringer".replace(",", "."))
    else:
        data["mc"] = genbrug("membercare", "mc")

    # Valeurs egne ark
    for rolle, felt, fn in (("benchmark", ("benchmark",), benchmark), ("mails", ("mails",), mails)):
        if valgt.get(rolle):
            log(f"Læser {K.ROLLE[rolle]['navn'].lower()} …")
            data[felt[0]] = fn(valgt[rolle])
            if rolle == "benchmark":
                pp = data["benchmark"]["perioder"]
                ny(rolle, pp[0]["periode"] if pp else None, pp[-1]["periode"] if pp else None)
            else:
                kk = list(data["mails"])
                ny(rolle, kk[0].replace("-U", " uge ") if kk else None, kk[-1].replace("-U", " uge ") if kk else None)
        else:
            data[felt[0]] = genbrug(rolle, *felt)
    data["valeur_md"] = {}
    if valgt.get("valeur_md"):
        data["valeur_md"]["maaneder"] = valeur_maaned(valgt["valeur_md"])
        mm = data["valeur_md"]["maaneder"]
        ny("valeur_md", mm[0]["md"] if mm else None, mm[-1]["md"] if mm else None)
    else:
        data["valeur_md"]["maaneder"] = genbrug("valeur_md", "valeur_md", "maaneder") or []
    if valgt.get("efterbehandling"):
        data["valeur_md"]["efterbehandling"] = efterbehandling(valgt["efterbehandling"])
        eb = data["valeur_md"]["efterbehandling"]
        ny("efterbehandling", eb[0]["md"] if eb else None, eb[-1]["md"] if eb else None)
    else:
        data["valeur_md"]["efterbehandling"] = genbrug("efterbehandling", "valeur_md", "efterbehandling") or []

    # Dybere spadestik og medlemmer
    import dybde, medlemmer
    if valgt.get("membercare") and valgt.get("koe_kald") and valgt.get("valeur_dag"):
        log("Beregner de dybere spadestik …")
        data["dybde"] = dybde.build(valgt["membercare"], list(valgt["koe_kald"].values()),
                                    data["valeur"]["dage"], data["valeur"]["sidste"])
    else:
        data["dybde"] = tidligere.get("dybde")
    om = tidligere.get("om") or {}
    if valgt.get("valeur_pris"):
        om = {**om, "oekonomi": medlemmer.oekonomi(valgt["valeur_pris"])}
        ny("valeur_pris", data["valeur"]["foerste"], data["valeur"]["sidste"], note="Totaler for hele perioden")
    else:
        genbrug("valeur_pris")
    lister = valgt.get("medlemsliste")
    if lister and valgt.get("membercare"):
        log("Kobler medlemslister og Membercare …")
        om = {**om, "medlemmer": medlemmer.medlemmer(lister, valgt["membercare"])}
        ny("medlemsliste", min(lister), max(lister), note=f"{len(lister)} lister")
    else:
        genbrug("medlemsliste")
    data["om"] = om or None

    data["kilder"] = {r["id"]: meta.get(r["id"], {"status": "mangler"}) for r in K.ROLLER}
    for r in data["kilder"].values():
        if r.get("status") == "ny":
            r["opdateret"] = data["genereret"]

    log("Tjekker for persondata …")
    txt = json.dumps(data, ensure_ascii=False, default=str)
    fund = tjek_privatliv(txt, navne)
    if fund:
        raise ValueError("Outputtet indeholder persondata (" + ", ".join(fund) + "). Intet er gemt.")
    log("Færdig.")
    return data, oversigt


def _hent(d, nøgler):
    for n in nøgler:
        if not isinstance(d, dict) or n not in d:
            return None
        d = d[n]
    return d


if __name__ == "__main__":
    her = Path(__file__).resolve().parent
    repo = her.parents[1] if (her.parents[1] / "docs").exists() else None
    rod = Path(sys.argv[1]) if len(sys.argv) > 1 else (repo / "data" / "raw" / "thomas" if repo else her.parent)
    ud = (repo / "docs" / "briefing.json") if repo else her / "data.json"
    tidligere = json.loads(ud.read_text(encoding="utf-8")) if ud.exists() else None
    data, oversigt = byg(rod, tidligere)
    ud.write_text(json.dumps(data, ensure_ascii=False, default=str), encoding="utf-8")
    for r in K.ROLLER:
        m = data["kilder"][r["id"]]
        print(f"  {m['status']:9s} {r['navn']:42s} {m.get('fra') or '':>10} – {m.get('til') or '':<10} {', '.join(m.get('filer') or [])[:60]}")
    print("skrev", ud)

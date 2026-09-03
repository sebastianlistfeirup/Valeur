"""Læser Valeurs ugerapporter, benchmark-ark og mailoversigt.

Tre ting, de øvrige kilder ikke kan svare på, ligger her:

  Straksbesvarelse   Agentperformance-arket har "Omstillet/videresendt" pr.
                     agent pr. måned. Besvarede minus omstillede er de opkald,
                     Valeur klarede fra ende til anden — den KPI, dashboardet
                     hidtil kun kunne anslå.
  Medlemsstatus      Ugerapporterne registrerer, om den der ringer er medlem,
                     ikke-medlem eller associeret.
  Tid pr. emne       Samme rapporter giver samtaletid og efterbehandling for
                     hvert enkelt emne — altså hvad en henvendelsestype koster.

Ugerapporterne findes både som .xlsx og som .mhtml (en SSRS-udskrift med
base64-kodet HTML). Begge parses til samme form.
"""

import base64
import email
import glob
import io
import re
from pathlib import Path

import pandas as pd

RAW = Path(__file__).resolve().parent.parent / "data" / "raw" / "valeur"
KOLONNER = ["kategori", "koe", "emne", "antal", "pct_kategori", "pct_total",
            "samtale", "efterbehandling"]

# Kategorinavnene er lange og bærer et præfiks, der er ens for alle.
KATEGORI = {
    "henvendelsesårsager": "emne",
    "medlemsgrupper": "medlemsgruppe",
    "medlemskab": "medlemskab",
    "identificering af ledelse": "rolle",
}


def _sek(v):
    if not isinstance(v, str) or not re.match(r"^\d+:\d{2}:\d{2}$", v.strip()):
        return None
    t, m, s = (int(x) for x in v.strip().split(":"))
    return t * 3600 + m * 60 + s


def _mmss(sek):
    if sek is None or pd.isna(sek):
        return None
    s = int(round(sek))
    return f"{s // 60}:{s % 60:02d}"


def _kategori(raa):
    """'Psykologforeningen | Henvendelsesårsager (opkald)' → 'emne'."""
    t = str(raa).split("|")[-1].strip().lower()
    for noegle, kort in KATEGORI.items():
        if t.startswith(noegle):
            return kort
    return None


def _periode(tekst):
    d = re.findall(r"(\d{2})/(\d{2})/(\d{4})", str(tekst))
    if len(d) >= 2:
        return (f"{d[0][2]}-{d[0][1]}-{d[0][0]}", f"{d[1][2]}-{d[1][1]}-{d[1][0]}")
    return (None, None)


def _fra_mhtml(sti):
    """Pakker SSRS-udskriften ud og finder rapportens tabel."""
    msg = email.message_from_bytes(sti.read_bytes())
    for del_ in msg.walk():
        if del_.get_content_type() != "text/html":
            continue
        raa = del_.get_payload(decode=False)
        try:
            html = base64.b64decode(raa).decode("utf-8", errors="replace")
        except Exception:
            html = raa
        tabeller = pd.read_html(io.StringIO(html))
        for t in tabeller:
            if t.shape[1] >= 8 and t.shape[0] > 10:
                return t
    return None


def _laes_uge(sti):
    if sti.suffix == ".mhtml":
        raa = _fra_mhtml(sti)
        if raa is None:
            return None, (None, None)
        df = raa.iloc[:, :8].copy()
        df.columns = KOLONNER
        df = df[df["emne"].notna()]
        periode = (None, None)
    else:
        df = pd.read_excel(sti, sheet_name=0, header=0)
        df = df.iloc[:, :8]
        df.columns = KOLONNER
        try:
            p = pd.read_excel(sti, sheet_name="Report parameters", header=None)
            periode = _periode(" ".join(str(v) for v in p.values.ravel()))
        except Exception:
            periode = (None, None)

    if periode == (None, None):
        m = re.search(r"(\d{4})(\d{2})(\d{2})_(\d{4})(\d{2})(\d{2})", sti.stem)
        if m:
            a = m.groups()
            periode = (f"{a[0]}-{a[1]}-{a[2]}", f"{a[3]}-{a[4]}-{a[5]}")

    # Rækker med koe == "Total" gentager alt fra den navngivne kø. Behold én.
    df = df[df["koe"].astype(str).str.strip().str.lower() != "total"]
    df = df[df["emne"].astype(str).str.strip().str.lower() != "total"]
    df["felt"] = df["kategori"].map(_kategori)
    df = df[df["felt"].notna()].copy()
    df["antal"] = pd.to_numeric(df["antal"], errors="coerce").fillna(0).astype(int)
    df["samtale_sek"] = df["samtale"].map(_sek)
    df["efter_sek"] = df["efterbehandling"].map(_sek)
    df["emne"] = df["emne"].astype(str).str.replace(r"^[•\t\s]+", "", regex=True).str.strip()
    return df[df["antal"] > 0], periode


def _vejet(df, kol):
    w = df["antal"].where(df[kol].notna())
    return float((df[kol] * w).sum() / w.sum()) if w.sum() else None


def _uger():
    filer = sorted(glob.glob(str(RAW / "uger" / "*.xlsx")) +
                   glob.glob(str(RAW / "uger" / "*.mhtml")))
    dele, perioder = [], []
    for f in filer:
        try:
            df, p = _laes_uge(Path(f))
        except Exception:
            continue
        if df is None or df.empty:
            continue
        df["fil"] = Path(f).stem
        dele.append(df)
        if p[0]:
            perioder.append(p)
    if not dele:
        return None
    alle = pd.concat(dele, ignore_index=True)

    def opgoer(felt, minimum=1):
        s = alle[alle["felt"] == felt]
        ud = []
        for navn, g in s.groupby("emne"):
            n = int(g["antal"].sum())
            if n < minimum:
                continue
            ud.append({
                "navn": navn, "antal": n,
                "samtale_sek": round(_vejet(g, "samtale_sek") or 0),
                "efter_sek": round(_vejet(g, "efter_sek") or 0),
                "samtale": _mmss(_vejet(g, "samtale_sek")),
                "efterbehandling": _mmss(_vejet(g, "efter_sek")),
            })
        ud.sort(key=lambda r: -r["antal"])
        return ud

    medlemskab = opgoer("medlemskab")
    for r in medlemskab:
        n = r["navn"].lower()
        r["type"] = ("medlem" if n.startswith("medlem")
                     else "associeret" if "associeret" in n
                     else "ikke-medlem")

    i_alt = sum(r["antal"] for r in medlemskab)
    medlem = sum(r["antal"] for r in medlemskab if r["type"] == "medlem")
    return {
        "uger": int(alle["fil"].nunique()),
        "periode_fra": min(p[0] for p in perioder) if perioder else None,
        "periode_til": max(p[1] for p in perioder) if perioder else None,
        "emner": opgoer("emne", minimum=10),
        "medlemsgrupper": opgoer("medlemsgruppe"),
        "medlemskab": medlemskab,
        "roller": opgoer("rolle"),
        "medlemsandel": round(medlem / i_alt * 100, 1) if i_alt else None,
        "registreringer": i_alt,
    }


def _agentark():
    """Agentopgørelsen ligger i to udgaver af benchmark-filen under hvert sit
    arknavn. Tag den, der dækker flest måneder."""
    bedst = None
    for fil in ("benchmark.xlsx", "benchmark_team.xlsx"):
        sti = RAW / fil
        if not sti.exists():
            continue
        for ark in ("Agentperformance", "Team Performance"):
            try:
                df = pd.read_excel(sti, sheet_name=ark, header=None)
            except Exception:
                continue
            if bedst is None or len(df) > len(bedst):
                bedst = df
    return bedst


def _agenter():
    df = _agentark()
    if df is None:
        return None
    maaneder, aktuel = [], None
    for _, r in df.iterrows():
        v = [("" if pd.isna(x) else str(x)).strip() for x in r]
        if v[0] and not v[1] and v[0].replace(" ", "").isalpha():
            aktuel = {"maaned": v[0].capitalize(), "besvarede": 0, "omstillet": 0}
            maaneder.append(aktuel)
        elif aktuel and v[0].isdigit() and len(v) > 5:
            b = pd.to_numeric(v[2], errors="coerce")
            o = pd.to_numeric(v[5], errors="coerce")
            if pd.notna(b):
                aktuel["besvarede"] += int(b)
            if pd.notna(o):
                aktuel["omstillet"] += int(o)
    maaneder = [m for m in maaneder if m["besvarede"] > 0]
    for m in maaneder:
        m["selv"] = m["besvarede"] - m["omstillet"]
        m["andel"] = round(m["selv"] / m["besvarede"] * 100, 1)
    if not maaneder:
        return None
    b = sum(m["besvarede"] for m in maaneder)
    o = sum(m["omstillet"] for m in maaneder)
    return {
        "maaneder": maaneder,
        "besvarede": b, "omstillet": o, "selv": b - o,
        "andel": round((b - o) / b * 100, 1),
        "maal": 70,
    }


def _mails():
    sti = RAW / "mails.xlsx"
    if not sti.exists():
        return None
    ud = []
    for ark in ("2025", "2026"):
        try:
            df = pd.read_excel(sti, sheet_name=ark, header=None)
        except Exception:
            continue
        uge = None
        for _, r in df.iterrows():
            for i in range(0, df.shape[1] - 2, 4):
                a = ("" if pd.isna(r.iloc[i]) else str(r.iloc[i])).strip()
                v = pd.to_numeric(r.iloc[i + 2], errors="coerce") if i + 2 < len(r) else None
                m = re.match(r"^Uge\s*(\d+)$", a, re.I)
                if m:
                    uge = int(m.group(1))
                elif a.lower() == "total" and uge and pd.notna(v):
                    ud.append({"aar": int(ark), "uge": uge, "mails": int(v)})
                    uge = None
    if not ud:
        return None
    df = pd.DataFrame(ud).drop_duplicates(subset=["aar", "uge"])
    pr_aar = [{"aar": int(a), "uger": int(len(g)), "mails": int(g["mails"].sum()),
               "pr_uge": round(g["mails"].mean(), 1)}
              for a, g in df.groupby("aar")]
    return {
        "pr_aar": pr_aar,
        "i_alt": int(df["mails"].sum()),
        "raekker": df.sort_values(["aar", "uge"]).to_dict("records"),
    }


# Kun de mål, der står som rene procenter. Benchmark-arkets tidsceller er
# tvetydige — nogle er mm:ss, andre hh:mm:ss ("00:26:53" er 26 minutter,
# "00:47:00" er 47 sekunder) — så tiderne hentes i stedet fra IP Nordic og
# fra ugerapporterne, hvor formatet er entydigt.
MAAL_NAVN = {
    "sla svarrate": "Svarrate",
    "fcr": "Løst ved første opkald",
}


def _benchmark():
    """Valeurs egne servicemål pr. uge: målsætning mod opnået."""
    sti = RAW / "benchmark.xlsx"
    if not sti.exists():
        return None
    df = pd.read_excel(sti, sheet_name="Benchmark", header=None)
    maal = {}
    for _, r in df.iterrows():
        v = [("" if pd.isna(x) else str(x)).strip() for x in r]
        if len(v) < 4:
            continue
        navn = next((vist for n, vist in MAAL_NAVN.items()
                     if v[1].lower().strip().startswith(n)), None)
        if not navn:
            continue
        m = pd.to_numeric(v[2], errors="coerce")
        o = pd.to_numeric(v[3], errors="coerce")
        # Enkelte celler er tastet som "18.14" eller "36:38%". En andel uden
        # for 0-1 er en indtastningsfejl og tælles ikke med.
        if pd.isna(o) or not (0 < o <= 1):
            continue
        d = maal.setdefault(navn, {"navn": navn, "maal": None, "vaerdier": []})
        if d["maal"] is None and pd.notna(m):
            d["maal"] = float(m)
        d["vaerdier"].append(float(o))

    ud = []
    for d in maal.values():
        if not d["vaerdier"]:
            continue
        snit = sum(d["vaerdier"]) / len(d["vaerdier"])
        ud.append({
            "navn": d["navn"], "uger": len(d["vaerdier"]),
            "maal": round(d["maal"] * 100, 1) if d["maal"] else None,
            "snit": round(snit * 100, 1),
            "bedste": round(max(d["vaerdier"]) * 100, 1),
            "vaerste": round(min(d["vaerdier"]) * 100, 1),
            "opfyldt": bool(d["maal"] and snit >= d["maal"]),
        })
    return ud or None


def build():
    uger, agenter, mails = _uger(), _agenter(), _mails()
    bench = _benchmark()
    if not any((uger, agenter, mails, bench)):
        return None
    return {"uger": uger, "agenter": agenter, "mails": mails, "benchmark": bench}

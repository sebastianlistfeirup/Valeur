"""Læser IP Nordics eksporter: CDR-rådata og de daglige kø-rapporter.

Opkaldsrejsen, som den ser ud i CDR'en
--------------------------------------
Et medlem ringer til hovednummeret og lander ét af tre steder, alt efter
hvornår. Numrene er ikke navngivet i eksporten — rollerne herunder er udledt
af mønstrene og skal bekræftes af DP:

  3824      Velkomstbesked (tryksystemet). Median 13 sekunder, ingen samtale.
            Tungest om onsdagen og uden for 10-13 — altså lukkebeskeden.
  773345xx  Valeur. Starter 5. maj 2025, dagen samarbejdet gik i luften.
            Aldrig onsdag, næsten kun 10-13, median samtale 5½ minut.
  309060xx  Telefonsvarer. Døgnet rundt, alle ugedage, median 24 sekunder.

Kan Valeur ikke svare, viderestilles opkaldet til en kø hos sekretariatet
(1510 Ansættelse, 1515 Selvstændig, 15122 Kursus, 15123 Øvrige, 1520
Zexpense, 1525 Debitor). Køerne har deres egne rapporter med ventetid og
svarprocent.

Privatliv
---------
CDR'en rummer medlemmernes telefonnumre i klartekst og — i "Fra abonnent" —
navn og fuld adresse fra nummeroplysningen. Ingen af delene forlader denne
fil: numre bruges kun til at tælle gentagne opkald og kasseres derefter,
abonnentfeltet læses aldrig.
"""

import glob
from pathlib import Path

import pandas as pd

RAW = Path(__file__).resolve().parent.parent / "data" / "raw" / "ipnordic"

# Udledte roller. Bekræftes af DP — se modulets docstring.
VELKOMST = "3824"
VALEUR = "773345xx"
SVARER = "309060xx"

LAG = {
    VELKOMST: ("velkomst", "Velkomstbesked"),
    VALEUR: ("valeur", "Valeur"),
    SVARER: ("svarer", "Telefonsvarer"),
}

UGEDAGE = ["Mandag", "Tirsdag", "Onsdag", "Torsdag", "Fredag", "Lørdag", "Søndag"]
AABEN_FRA, AABEN_TIL = 10, 13          # DP's telefonåbningstid
LUKKEDAG = 2                            # onsdag


def _sek(serie):
    """'00:04:26' → 266.0. Tomme felter bliver til NaN, ikke nul."""
    s = serie.fillna("").astype(str).str.strip()
    delt = s.str.split(":", expand=True)
    if delt.shape[1] < 3:
        return pd.Series([float("nan")] * len(s), index=s.index)
    tal = delt.apply(pd.to_numeric, errors="coerce")
    return tal[0] * 3600 + tal[1] * 60 + tal[2]


def _mmss(sekunder):
    """266.0 → '4:26'. None hvis der ikke er noget at vise."""
    if sekunder is None or pd.isna(sekunder):
        return None
    s = int(round(sekunder))
    return f"{s // 60}:{s % 60:02d}"


def _vejet(df, kolonne, vaegt):
    """Vejet gennemsnit — et gennemsnit af daglige gennemsnit vægter en dag
    med tre opkald som en dag med tres."""
    v = df[kolonne]
    w = df[vaegt].where(v.notna())
    if w.sum() == 0:
        return None
    return float((v * w).sum() / w.sum())


# ── CDR ─────────────────────────────────────────────────────────────────────
def laes_cdr():
    df = pd.read_excel(RAW / "cdr.xlsx", dtype=str)
    df["ts"] = pd.to_datetime(df["Starttidspunkt"], format="%d-%m-%Y %H:%M:%S",
                              errors="coerce")
    df = df[df["ts"].notna()].copy()
    df["sek"] = _sek(df["Varighed"])
    df["til"] = df["Til"].fillna("").str.strip()
    df["fra"] = df["Fra"].fillna("").str.strip()
    return df


def build_cdr(df):
    ialt = len(df)
    lag_data = []
    ugedag_lag, time_lag, maaned_lag = {}, {}, {}

    for nr, (noegle, navn) in LAG.items():
        s = df[df["til"] == nr]
        if s.empty:
            continue
        uden_for = s[(s["ts"].dt.hour < AABEN_FRA) | (s["ts"].dt.hour >= AABEN_TIL)]
        lag_data.append({
            "noegle": noegle,
            "navn": navn,
            "nummer": nr,
            "opkald": int(len(s)),
            "median_sek": float(s["sek"].median()),
            "onsdag": int((s["ts"].dt.dayofweek == LUKKEDAG).sum()),
            "uden_for_aabningstid": int(len(uden_for)),
            "foerste": s["ts"].min().strftime("%Y-%m-%d"),
        })
        ugedag_lag[noegle] = [
            int((s["ts"].dt.dayofweek == i).sum()) for i in range(7)
        ]
        time_lag[noegle] = {
            int(t): int(n) for t, n in s["ts"].dt.hour.value_counts().items()
        }
        maaned_lag[noegle] = {
            str(m): int(n)
            for m, n in s["ts"].dt.to_period("M").value_counts().sort_index().items()
        }

    maaneder = sorted({m for d in maaned_lag.values() for m in d})

    # Gentagne opkald fra samme nummer — tælles, nummeret gemmes ikke.
    medlemsopkald = df[df["til"].isin([VALEUR, SVARER, VELKOMST])]
    pr_nummer = medlemsopkald.loc[
        medlemsopkald["fra"].str.match(r"^\d{8}$", na=False), "fra"
    ].value_counts()
    gentagne = {}
    for n in pr_nummer.values:
        gentagne["10+" if n >= 10 else str(n)] = gentagne.get(
            "10+" if n >= 10 else str(n), 0) + 1

    return {
        "raekker": ialt,
        "periode_fra": df["ts"].min().strftime("%Y-%m-%d"),
        "periode_til": df["ts"].max().strftime("%Y-%m-%d"),
        "lag": lag_data,
        "ugedage": UGEDAGE,
        "ugedag_lag": ugedag_lag,
        "time_lag": time_lag,
        "maaneder": maaneder,
        "maaned_lag": maaned_lag,
        "gentagne": {
            "numre": int(len(pr_nummer)),
            "opkald": int(pr_nummer.sum()),
            "fordeling": [
                {"opkald": k, "numre": gentagne[k]}
                for k in [str(i) for i in range(1, 10)] + ["10+"]
                if k in gentagne
            ],
        },
    }


# ── Kø-rapporter ────────────────────────────────────────────────────────────
def laes_koeer():
    filer = sorted(glob.glob(str(RAW / "queue_*.xlsx")))
    if not filer:
        return None
    df = pd.concat([pd.read_excel(f, dtype=str) for f in filer], ignore_index=True)
    df = df.drop_duplicates(subset=["Dato", "Kønummer"])
    df["dato"] = pd.to_datetime(df["Dato"], format="%d-%m-%Y", errors="coerce")
    for kol, ny in [("Opkald", "opkald"), ("Besvarede", "besvarede"),
                    ("Afbrudte", "afbrudte")]:
        df[ny] = pd.to_numeric(df[kol], errors="coerce").fillna(0)
    for kol, ny in [("Gns. samtaletid", "samtale"), ("Gns. ventetid", "vente"),
                    ("Maks. ventetid", "maks_vente"),
                    ("Gns. ventetid for afbrudte", "vente_afbrudt")]:
        df[ny] = _sek(df[kol])
    df["maaned"] = df["dato"].dt.to_period("M").astype(str)
    return df[df["dato"].notna()]


def build_koeer(df):
    koeer = []
    for (nr, navn), s in df.groupby(["Kønummer", "Kønavn"]):
        opkald, besvarede = int(s["opkald"].sum()), int(s["besvarede"].sum())
        koeer.append({
            "nummer": str(nr),
            "navn": str(navn).strip(),
            "opkald": opkald,
            "besvarede": besvarede,
            "svarprocent": round(besvarede / opkald * 100, 1) if opkald else None,
            "dage": int(len(s)),
            "gns_samtale": _mmss(_vejet(s, "samtale", "besvarede")),
            "gns_vente": _mmss(_vejet(s, "vente", "besvarede")),
            "maks_vente": _mmss(s["maks_vente"].max()),
        })
    koeer.sort(key=lambda r: -r["opkald"])

    maaneder = sorted(df["maaned"].unique())
    pr_maaned = []
    for m in maaneder:
        s = df[df["maaned"] == m]
        opkald, besvarede = int(s["opkald"].sum()), int(s["besvarede"].sum())
        pr_maaned.append({
            "maaned": m,
            "opkald": opkald,
            "besvarede": besvarede,
            "svarprocent": round(besvarede / opkald * 100, 1) if opkald else None,
            "gns_vente_sek": round(_vejet(s, "vente", "besvarede") or 0, 1),
            "gns_samtale_sek": round(_vejet(s, "samtale", "besvarede") or 0, 1),
            "pr_koe": {str(k).strip(): int(v) for k, v in
                       s.groupby("Kønavn")["opkald"].sum().items()},
        })

    # År til år: kun de måneder der findes i begge år, så en halv sæson ikke
    # sammenlignes med en hel.
    df = df.copy()
    df["aar"] = df["dato"].dt.year
    df["mnr"] = df["dato"].dt.month
    aar = sorted(df["aar"].unique())
    aar_til_aar = None
    if len(aar) >= 2:
        tidl, sen = aar[0], aar[-1]
        # Klip begge år til det kalendervindue, de har til fælles. Uden det
        # ville en hel september 2025 blive holdt op mod tre dage af 2026 —
        # og januar 2026 mod et 2025, hvor målingen slet ikke var begyndt.
        md = df["dato"].dt.strftime("%m-%d")
        fra = max(md[df.aar == tidl].min(), md[df.aar == sen].min())
        til = min(md[df.aar == tidl].max(), md[df.aar == sen].max())
        klip = df[md.between(fra, til)]
        faelles = sorted(set(klip.loc[klip.aar == tidl, "mnr"]) &
                         set(klip.loc[klip.aar == sen, "mnr"]))
        if faelles:
            raekker = []
            for navn, s in klip.groupby("Kønavn"):
                a = int(s.loc[s.aar == tidl, "opkald"].sum())
                b = int(s.loc[s.aar == sen, "opkald"].sum())
                raekker.append({
                    "navn": str(navn).strip(), "foer": a, "efter": b,
                    "aendring": round((b - a) / a * 100, 1) if a else None,
                })
            raekker.sort(key=lambda r: -(r["foer"] + r["efter"]))
            i_alt_a = sum(r["foer"] for r in raekker)
            i_alt_b = sum(r["efter"] for r in raekker)
            aar_til_aar = {
                "foer_aar": int(tidl), "efter_aar": int(sen),
                "maaneder": [int(m) for m in faelles],
                "fra": fra, "til": til,
                "raekker": raekker,
                "i_alt": {
                    "foer": i_alt_a, "efter": i_alt_b,
                    "aendring": round((i_alt_b - i_alt_a) / i_alt_a * 100, 1) if i_alt_a else None,
                },
            }

    opkald = int(df["opkald"].sum())
    besvarede = int(df["besvarede"].sum())
    return {
        "koeer": koeer,
        "maanedlig": pr_maaned,
        "aar_til_aar": aar_til_aar,
        "ialt": {
            "opkald": opkald,
            "besvarede": besvarede,
            "svarprocent": round(besvarede / opkald * 100, 1) if opkald else None,
            "gns_vente": _mmss(_vejet(df, "vente", "besvarede")),
            "gns_samtale": _mmss(_vejet(df, "samtale", "besvarede")),
            "periode_fra": df["dato"].min().strftime("%Y-%m-%d"),
            "periode_til": df["dato"].max().strftime("%Y-%m-%d"),
        },
    }


def build():
    if not (RAW / "cdr.xlsx").exists():
        return None
    cdr = laes_cdr()
    koe_df = laes_koeer()
    return {
        "kilde": "IP Nordic — CDR og daglige kø-rapporter",
        "cdr": build_cdr(cdr),
        "koeer": build_koeer(koe_df) if koe_df is not None else None,
        "roller_udledt": {
            "velkomst": VELKOMST, "valeur": VALEUR, "svarer": SVARER,
        },
    }

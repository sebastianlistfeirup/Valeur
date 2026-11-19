"""
Find kildefilerne ud fra deres INDHOLD, ikke deres navn.

Hver .xlsx-fil åbnes, og de første rækker og fanenavne afgør, hvilken kilde den er.
Findes flere versioner af samme kilde, bruges den nyeste (ændringsdato).
Kun medlemslister skal have en dato i filnavnet, fordi listen ikke selv angiver,
hvilken dag den gælder for: fx "Medlemsliste periode_30_09_26.xlsx".

Rollerne beskriver også, hvad der skal eksporteres. Den samme tekst vises i
tjeklisten på opdateringssiden (docs/opdater.html), via ROLLER.
"""
import os, re, datetime as dt
from pathlib import Path
import openpyxl

ROLLER = [
    # nøgle, navn, påkrævet, flere filer, system, hvad der skal eksporteres, hyppighed
    {"id": "valeur_dag", "navn": "Opkald pr. dag", "system": "Valeur Insights", "paakraevet": True, "flere": False,
     "eksport": "Kunderapporten \"Opkald per dag\" uden filter. Kolonnerne Date, Besvarede kald og Kald i alt.",
     "kendetegn": "Kolonnerne Date · Besvarede kald · Kald i alt"},
    {"id": "valeur_time", "navn": "Timefordeling", "system": "Valeur Insights", "paakraevet": True, "flere": False,
     "eksport": "\"Timefordelingen\" fra kunderapporten uden filter.",
     "kendetegn": "Kolonnerne Time · %GT Antal af unique_call_id"},
    {"id": "valeur_time_afbrudt", "navn": "Timefordeling for kald afbrudt inden kø", "system": "Valeur Insights",
     "paakraevet": True, "flere": False, "eksport": "\"Timefordelingen for kald afbrudt inden kø\" uden filter.",
     "kendetegn": "Kolonnerne Time · Andel kald afbrudt inden kø"},
    {"id": "valeur_ugedag", "navn": "Ugedagenes andel af kald", "system": "Valeur Insights", "paakraevet": True,
     "flere": False, "eksport": "\"Ugedages andel af kald\" fra kunderapporten uden filter.",
     "kendetegn": "Kolonnerne Ugedag · %GT Samlet antal opkald"},
    {"id": "valeur_medlem", "navn": "Medlem eller ikke medlem", "system": "Valeur Insights", "paakraevet": True,
     "flere": False, "eksport": "\"Medlem_Ikke_medlem\" uden filter.", "kendetegn": "Kolonnerne Årsag · Andel"},
    {"id": "valeur_pris", "navn": "Pris og tid i alt", "system": "Valeur Insights", "paakraevet": True, "flere": False,
     "eksport": "\"Opkaldsdata\" med filteret activityType_name indeholder ikke 'opstart'. Hele perioden fra pilotstart.",
     "kendetegn": "Kolonnerne Kundenavn · Besvarede kald · SamtaleTid i minutter · Samarbejde (kr.)"},
    {"id": "koe_maaned", "navn": "Køerne pr. måned", "system": "IP Nordic", "paakraevet": True, "flere": True,
     "eksport": "Kø-rapporten opdelt pr. måned, én fil pr. kø: 1510 Ansættelse, 1515 Selvstændig, 15122 Kursus, "
                "15123 Øvrige/Medlemsservice, 1520 Zexpense og 1525 Debitor. Fra februar 2024 til i dag.",
     "kendetegn": "Kolonnerne Dato · Kønummer · Kønavn · Opkald, med måneder som \"January 2026\""},
    {"id": "koe_kald", "navn": "Hvert kald i køerne", "system": "IP Nordic", "paakraevet": True, "flere": True,
     "eksport": "\"Alle opkald opdelt pr. indkomne opkald\", én fil pr. kø (samme seks køer). Fra februar 2024 til i dag.",
     "kendetegn": "Kolonnerne Nummer · Kønummer · Kønavn · Starttidspunkt"},
    {"id": "membercare", "navn": "Registreringer i servicefanen", "system": "Membercare", "paakraevet": True,
     "flere": False, "eksport": "Alle registreringer i servicefanen fra 5. maj 2025 til i dag.",
     "kendetegn": "Kolonnerne Stamkort · Medlem · Servicegruppe · Servicetype · Service dato"},
    {"id": "medlemsliste", "navn": "Medlemslister", "system": "Membercare", "paakraevet": True, "flere": True,
     "eksport": "Medlemslisten pr. kvartalsskifte. Den skal have kolonnerne Medlemstype, Stamkort, Kreds, Alder, "
                "DP Indm.dato og DP Udm.dato; alle andre kolonner (adresse, e-mail, telefon, arbejdsgiver) kan og bør "
                "slettes før upload. Datoen SKAL stå i filnavnet, "
                "fx \"Medlemsliste periode_31_12_26.xlsx\". Upload alle kvartaler fra 31. marts 2025, ikke kun det nyeste.",
     "kendetegn": "Kolonnerne Medlemstype · Stamkort · DP Indm.dato · DP Udm.dato"},
    {"id": "benchmark", "navn": "Benchmark-ark", "system": "Valeur (manuelt ark)", "paakraevet": False, "flere": False,
     "eksport": "Valeurs benchmark-ark med fanen \"Benchmark\" (mål og opnået pr. uge/måned).",
     "kendetegn": "Fanen Benchmark med Målsætning · Opnået"},
    {"id": "valeur_md", "navn": "Valeurs månedsopgørelse", "system": "Valeur (manuelt ark)", "paakraevet": False,
     "flere": False, "eksport": "Ark med fanen \"Opkald - månedsopdelt\" (tilbudte og besvarede kald pr. måned).",
     "kendetegn": "Fanen Opkald - månedsopdelt"},
    {"id": "efterbehandling", "navn": "Efterbehandling pr. måned", "system": "Valeur (manuelt ark)",
     "paakraevet": False, "flere": False, "eksport": "Ark med fanen \"Efterbehandling\" (efterbehandlingstid pr. måned).",
     "kendetegn": "Fanen Efterbehandling"},
    {"id": "mails", "navn": "Mails pr. uge", "system": "Manuel optælling", "paakraevet": False, "flere": False,
     "eksport": "Oversigten over mails, ugeopdelt, med én fane pr. år (fx \"2025\", \"2026\").",
     "kendetegn": "Første celle \"Mails til Dansk Psykolog Forening\""},
]
ROLLE = {r["id"]: r for r in ROLLER}


def _txt(v):
    return str(v).strip() if v is not None else ""


def _start(row, *navne):
    row = [_txt(c) for c in row]
    return len(row) >= len(navne) and all(row[i] == n for i, n in enumerate(navne))


def dato_fra_navn(navn):
    """Gyldighedsdato for en medlemsliste ud fra filnavnet."""
    m = re.search(r"periode[ _-]*(\d{1,2})[._-](\d{1,2})[._-](\d{2,4})", navn, re.I)
    if not m:
        kandidater = re.findall(r"(?<!\d)(\d{1,2})[._-](\d{1,2})[._-](\d{2,4})(?!\d)", navn)
        m = kandidater[-1] if kandidater else None
        g = m if m else None
    else:
        g = m.groups()
    if not g:
        return None
    d, mth, y = (int(x) for x in g)
    y = y + 2000 if y < 100 else y
    try:
        return dt.date(y, mth, d)
    except ValueError:
        return None


def klassificer(sti):
    """Returnér liste af (rolle, nøgle, ekstra) for en fil. Tom liste = ukendt fil."""
    try:
        wb = openpyxl.load_workbook(sti, read_only=True, data_only=True)
    except Exception as e:  # låst, beskadiget eller ikke en Excel-fil
        return [("fejl", None, {"fejl": type(e).__name__})]
    try:
        fund = []
        ark = wb.sheetnames
        ws = wb.worksheets[0]
        ws.reset_dimensions()          # Valeur Insights-eksporter angiver forkert arkstørrelse (A1:A1)
        rows = [list(r) for r in ws.iter_rows(max_row=4, values_only=True)]
        rows += [[]] * (4 - len(rows))
        r0, r1 = rows[0], rows[1]
        foerste = _txt(r0[0]) if r0 else ""

        for r in rows[:3]:
            if _start(r, "Date", "Besvarede kald", "Kald i alt"):
                fund.append(("valeur_dag", None, {}))
            elif _start(r, "Time", "%GT Antal af unique_call_id"):
                fund.append(("valeur_time", None, {}))
            elif _start(r, "Time", "Andel kald afbrudt inden kø"):
                fund.append(("valeur_time_afbrudt", None, {}))
            elif _start(r, "Ugedag", "%GT Samlet antal opkald"):
                fund.append(("valeur_ugedag", None, {}))
            elif _start(r, "Årsag", "Andel"):
                fund.append(("valeur_medlem", None, {}))
            elif "Kundenavn" in [_txt(c) for c in r] and "Samarbejde (kr.)" in [_txt(c) for c in r]:
                fund.append(("valeur_pris", None, {}))
        if fund:
            fund = [(a, b, {"filter": foerste.replace("\n", " ")}) for a, b, _ in fund]
        if _start(r0, "Dato", "Kønummer", "Kønavn", "Opkald") and re.match(r"^[A-Z][a-z]+ \d{4}$", _txt(r1[0] if r1 else "")):
            fund.append(("koe_maaned", _txt(r1[1]), {}))
        if _start(r0, "Nummer", "Kønummer", "Kønavn", "Starttidspunkt") and len(r1) > 1:
            fund.append(("koe_kald", _txt(r1[1]), {}))
        h0 = [_txt(c) for c in r0]
        if {"Stamkort", "Servicegruppe", "Servicetype", "Service dato"} <= set(h0):
            fund.append(("membercare", None, {}))
        if {"Medlemstype", "DP Indm.dato", "DP Udm.dato"} <= set(h0):
            d = dato_fra_navn(Path(sti).name)
            sk = h0.index("Stamkort") if "Stamkort" in h0 else (2 if len(h0) > 2 and h0[2] == "" else None)
            har_stamkort = sk is not None and len(r1) > sk and r1[sk] not in (None, "")
            fund.append(("medlemsliste", d.isoformat() if d else None, {"stamkort": har_stamkort}))
        if "Benchmark" in ark:
            fund.append(("benchmark", None, {}))
        if "Opkald - månedsopdelt" in ark:
            fund.append(("valeur_md", None, {}))
        if "Efterbehandling" in ark:
            fund.append(("efterbehandling", None, {}))
        if foerste.lower().startswith("mails") and any(re.fullmatch(r"20\d\d", a.strip()) for a in ark):
            fund.append(("mails", None, {}))
        return fund
    finally:
        wb.close()


def find(rod, mtimes=None, log=print):
    """Gennemgå alle .xlsx under rod. Returnér (valgt, oversigt).

    valgt:    {rolle: sti}  eller for flere-filer-roller {rolle: {nøgle: sti}}
    oversigt: liste over alle filer med rolle og om de blev brugt (til tjeklisten)
    """
    rod = Path(rod)
    mtimes = mtimes or {}
    filer = sorted(p for p in rod.rglob("*.xlsx") if not p.name.startswith("~$"))
    kandidater, oversigt = {}, []
    for i, p in enumerate(filer, 1):
        rel = str(p.relative_to(rod)).replace("\\", "/")
        mt = mtimes.get(rel) or p.stat().st_mtime * 1000
        stoerrelse = p.stat().st_size
        fund = klassificer(p)
        if i % 10 == 0:
            log(f"Gennemgået {i} af {len(filer)} filer")
        if not fund:
            oversigt.append({"fil": rel, "rolle": None, "brugt": False})
            continue
        for rolle, noegle, ekstra in fund:
            post = {"fil": rel, "rolle": rolle, "noegle": noegle, "mtime": mt, "str": stoerrelse, **ekstra, "brugt": False}
            oversigt.append(post)
            if rolle == "fejl":
                continue
            if rolle == "medlemsliste" and (not noegle or not ekstra.get("stamkort")):
                post["advarsel"] = ("Ingen dato i filnavnet" if not noegle else
                                    "Listen har ingen stamkortnumre og kan kun bruges til at tælle medlemmer")
                if not noegle:
                    continue
            k = (rolle, noegle)
            # nyeste version vinder; ved (næsten) samme tidspunkt den største fil
            if k not in kandidater or (round(mt / 60000), stoerrelse) > (round(kandidater[k]["mtime"] / 60000), kandidater[k]["str"]):
                kandidater[k] = post
    valgt = {}
    for (rolle, noegle), post in kandidater.items():
        post["brugt"] = True
        sti = rod / post["fil"]
        if ROLLE[rolle]["flere"]:
            valgt.setdefault(rolle, {})[noegle] = sti
        else:
            valgt[rolle] = sti
    return valgt, oversigt

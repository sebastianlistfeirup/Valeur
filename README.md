# Medlemmernes indgang

Dashboard over henvendelserne til Dansk Psykolog Forenings hovednummer —
Valeursamarbejdet.

Bygget efter samme model som udsendelsesdashboardet: et script normaliserer
kilderne til én datafil, og siden læser den. Forskellen er, at data her ikke
kommer fra ét API, men fra tre kilder, der skal lægges oven på hinanden.

## De tre kilder

| Kilde | Hvad den ved | Omfang |
|---|---|---|
| **Membercare (CRM)** — servicefanen | Hvem der ringer, hvad de ringer om, hvornår | 9.548 registreringer |
| **IP Nordic** — CDR og kø-rapporter | Hvert opkaldsforsøg: hvor det landede, ventetid, svarprocent | 66.910 poster, 7 køer, 2023–2026 |
| **Valeur Insights** | Hvor mange kald Valeur selv når at tage | 15.080 kald, 257 dage |
| **Valeurs ugerapporter** | Medlemsstatus, samtaletid og efterbehandling pr. emne | 20 rapporter, 5.491 registreringer |
| **Medlemslisten** | Bestanden, så hver gruppes andel af opkaldene kan holdes op mod dens andel af medlemmerne | Maj 2025 og september 2026 |
| **DP's egen opgørelse** | Opkald til sekretariatet januar–marts, 2025 mod 2026 | Uafhængig bekræftelse |

De tre første tæller ikke det samme og må ikke lægges sammen. CRM ser kun de
henvendelser, nogen talte med et medlem om; telefonidataene ser hvert forsøg,
også dem der ramte en telefonsvarer. Derfor står der 9.548 registreringer over
for 33.882 opkald.

### Opkaldsrejsen

Numrene er ikke navngivet i IP Nordics eksport. Rollerne er udledt af
mønstrene og **bør bekræftes af DP**:

| Nummer | Rolle | Hvorfor | Opkald |
|---|---|---|---|
| `773345xx` | Valeur | Går i luften 5. maj 2025, dagen samarbejdet startede. Aldrig onsdag, næsten kun 10–13. Median samtale 5:25 | 15.529 |
| `309060xx` | Telefonsvarer | Døgnet rundt, alle ugedage. Median 24 sekunder | 12.382 |
| `3824` | Velkomstbesked | Median 13 sekunder, tungest om onsdagen og uden for åbningstid | 5.971 |

Kan Valeur ikke svare, sendes opkaldet til en kø hos sekretariatet — 1510
Ansættelse, 1515 Selvstændig, 15122 Kursus, 15123 Øvrige, 1520 Zexpense,
1525 Debitor. Kø 1500 (Velkommen/Øvrige) blev nedlagt, da Valeur overtog
rollen.

### Før og efter

Kø-rapporterne rækker tilbage til 2023, så samarbejdet kan måles mod tiden
før. To regler holder sammenligningen ærlig: kun køer der findes i **begge**
perioder tælles med (ellers ville den nedlagte kø 1500 alene gøre "før"
dobbelt så stor), og alt opgøres **pr. åbningsdag**, fordi perioderne ikke er
lige lange.

## Kør det

```sh
pip install pandas openpyxl
python3 scripts/build_data.py
```

Modulerne er én pr. kilde: `ipnordic.py`, `valeur.py`, `egne.py`. Hver
returnerer `None`, hvis dens filer mangler, så siden bygger videre uden dem.

Scriptet læser `data/raw/`, skriver `docs/data.json`, og renderer to udgaver af
siden: `docs/index.html` (til GitHub Pages) og `dist/dashboard.html` (til
udgivelse som artifact).

## Privatliv

Kildefilerne indeholder personoplysninger — navne, stamkortnumre, og fritekst
med telefonnumre og mailadresser. Derfor:

- `data/raw/` er holdt uden for versionsstyring.
- Bemærkningsfeltet i CRM bruges ikke. Godt hver fjerde bemærkning rummer
  personhenførbar fritekst.
- Telefonnumrene i CDR-udtrækket bruges kun til at tælle gentagne opkald og
  kasseres derefter. Nummeroplysningsfeltet — navn og fuld adresse — læses aldrig.
- Ingen krydstabelcelle under 25 registreringer; grupper under fem lægges sammen.
- Kørslen afbrydes automatisk, hvis resultatet alligevel matcher et medlemsnavn,
  et telefonnummer, en e-mail eller et cpr-lignende nummer.

Registranternes fornavne er en bevidst undtagelse — det er en intern driftsmetrik,
og undtagelsen er navngivet i `assert_clean()`.

## Filer

```
scripts/build_data.py   samler kilderne → data.json → siden
scripts/ipnordic.py     CDR og kø-rapporter
scripts/valeur.py       Valeur Insights
scripts/valeur_rapporter.py  ugerapporter, benchmark og agentopgørelse
scripts/medlemmer.py    medlemslisterne, koblet til CRM på stamkort
scripts/egne.py         DP's egen 2025/2026-opgørelse
scripts/template.html   dashboardet: markup, stil og tegnekode
docs/                   det, GitHub Pages udstiller
data/raw/               kildefiler (uden for versionsstyring)
```

## Briefing-dashboardet (`docs/briefing.html`)

En selvstændig side, bygget efter medarbejderbriefingen fra september 2025
(*Valeur_Statistik_Medarbejderbriefing_17_09_25*), med data frem til november 2026.
Den læser eksporterne direkte fra mappen "Thomas data", som den ligger på
skrivebordet, og kræver hverken CDR-udtrækket eller medlemslisterne.

Sektioner: status, opkald og svar, aflastning af sekretariatet (før og efter, kø
for kø), hvornår der ringes, hvad og hvem (Membercare), Valeur målt mod de
aftalte mål, mails pr. uge og en liste over datahuller.

### Opdatér med nye data

**Den nemme vej:** gå til [`opdater.html`](https://sebastianlistfeirup.github.io/Valeur/opdater.html)
(linket står også nederst på dashboardet).

1. Siden viser en tjekliste over alle kilder: hvad der skal eksporteres, og hvilken
   periode dashboardet bygger på lige nu.
2. Vælg mappen med de nye eksporter. Filerne genkendes ud fra indholdet, så navne og
   undermapper er ligegyldige, og den nyeste version af hver kilde bliver brugt. Kun
   medlemslister skal have datoen i filnavnet, fx `Medlemsliste periode_31_12_26.xlsx`.
3. Tryk **Byg**. Browseren kører de samme Python-scripts som nedenfor via Pyodide,
   lokalt i en baggrundstråd. Ingen filer forlader computeren. Det tager 2–5 minutter.
4. Se forhåndsvisningen, log ind med en admin-nøgle og tryk **Publicér**. Kun
   `docs/briefing.json` og `docs/briefing.html` bliver skrevet til repoet.

Admin-nøglen er en *fine-grained personal access token* med adgang til kun dette repo
og *Contents: Read and write*. Den gemmes kun i browserfanen (sessionStorage).

Mangler en kilde i en upload, genbruges de tal, der sidst blev publiceret, og
tjeklisten markerer det.

**Fra kommandolinjen:**

```sh
python -X utf8 scripts/briefing/build_data.py "<mappe med eksporter>"   # → docs/briefing.json
python -X utf8 scripts/briefing/render.py                               # → docs/briefing.html, dist/, docs/py/
```

Uden argument læser scriptet `data/raw/thomas/`. `render.py` kopierer også
byggescripts og tjeklistens tekster til `docs/py/`, som opdateringssiden henter.
Kør den altid efter ændringer i `scripts/briefing/`.

`scripts/briefing/kilder.py` genkender kildefilerne ud fra deres indhold og
indeholder tjeklistens tekster (`ROLLER`).

`scripts/briefing/dybde.py` står for de seks dybere spadestik: hvor meget Valeur
løser selv, udsendelsernes effekt, årshjulet, sager der kommer igen, geografi
samt køopgivelse og prognose. Modulet læser bemærkningsfeltet i Membercare,
men kun for at klassificere hver registrering (besked/viderestilling ja/nej,
nævner en rykker ja/nej). Fritekst, navne og overskrifter fjernes, før noget
regnes ud, og outputtet er udelukkende optællinger.

`scripts/briefing/medlemmer.py` står for sektionen Økonomi og medlemmer. Den
læser Valeurs samlede pris og tid og de kvartalsvise medlemslister. Fra
medlemslisterne læses kun kreds, medlemstype, stamkort, alder og ind- og
udmeldelsesdatoer, fundet ud fra kolonneoverskrifterne, så andre kolonner kan og
bør slettes før upload. Stamkort bruges kun til at koble til Membercare, og ingen
gruppe under 25 personer vises.

Samme privatlivsregler gælder overalt: kun aggregerede tal, ingen krydstabelcelle
under 5, og kørslen stopper, hvis outputtet indeholder telefonnumre, mailadresser,
cpr-lignende numre eller navne fra Membercare.

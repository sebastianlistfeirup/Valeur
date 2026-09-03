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
| **IP Nordic** — CDR og kø-rapporter | Hvert opkaldsforsøg: hvor det landede, ventetid, svarprocent | 66.910 poster, 6 køer |
| **Valeur Insights** | Hvor mange kald Valeur selv når at tage | 15.080 kald, 257 dage |
| **DP's egen opgørelse** | Opkald til sekretariatet januar–marts, 2025 mod 2026 | Den eneste før-måling |

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
1525 Debitor.

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

Registranternes navne er en bevidst undtagelse — det er en intern driftsmetrik,
og undtagelsen er navngivet i `assert_clean()`.

## Filer

```
scripts/build_data.py   samler kilderne → data.json → siden
scripts/ipnordic.py     CDR og kø-rapporter
scripts/valeur.py       Valeur Insights
scripts/egne.py         DP's egen 2025/2026-opgørelse
scripts/template.html   dashboardet: markup, stil og tegnekode
docs/                   det, GitHub Pages udstiller
data/raw/               kildefiler (uden for versionsstyring)
```

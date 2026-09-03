# Medlemmernes indgang

Dashboard over henvendelserne til Dansk Psykolog Forenings hovednummer —
Valeursamarbejdet.

Bygget efter samme model som udsendelsesdashboardet: et script normaliserer
kilderne til én datafil, og siden læser den. Forskellen er, at data her ikke
kommer fra ét API, men fra tre kilder, der skal lægges oven på hinanden.

## De tre kilder

| Kilde | Hvad den ved | Status |
|---|---|---|
| **Membercare (CRM)** — servicefanen | Hvem der ringer, hvad de ringer om, hvornår | I drift |
| **IP Nordic** — statistikmodulet | Alle opkald: svarprocent, ventetid, samtaletid, kø | Afventer eksport |
| **Valeur Insights** | Valeurs egen håndtering: straksbesvaret mod viderestillet | Afventer eksport |

CRM rummer kun de henvendelser, der **blev** registreret. Et opkald, der aldrig
blev besvaret, findes ikke i CRM — derfor kan svarprocent, ventetid og
straksbesvarelsesrate ikke udledes af denne kilde alene.

## Kør det

```sh
pip install pandas openpyxl
python3 scripts/build_data.py
```

Scriptet læser `data/raw/`, skriver `docs/data.json`, og renderer to udgaver af
siden: `docs/index.html` (til GitHub Pages) og `dist/dashboard.html` (til
udgivelse som artifact).

## Privatliv

Kildefilerne indeholder personoplysninger — navne, stamkortnumre, og fritekst
med telefonnumre og mailadresser. Derfor:

- `data/raw/` er holdt uden for versionsstyring.
- Bemærkningsfeltet bruges ikke. Godt hver fjerde bemærkning rummer
  personhenførbar fritekst.
- Ingen krydstabelcelle under 25 registreringer; grupper under fem lægges sammen.
- Kørslen afbrydes automatisk, hvis resultatet alligevel matcher et medlemsnavn,
  et telefonnummer, en e-mail eller et cpr-lignende nummer.

Registranternes navne er en bevidst undtagelse — det er en intern driftsmetrik,
og undtagelsen er navngivet i `assert_clean()`.

## Filer

```
scripts/build_data.py   normaliserer kilder → data.json → siden
scripts/template.html   dashboardet: markup, stil og tegnekode
docs/                   det, GitHub Pages udstiller
data/raw/               kildefiler (uden for versionsstyring)
```

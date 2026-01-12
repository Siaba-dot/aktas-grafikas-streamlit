# Streamlit aplikacija: Akto atnaujinimas iš grafiko

Ši aplikacija leidžia įkelti **grafiką** (.ods/.xlsx/.xls) ir **aktą** (.xls/.xlsx),
bei sugeneruoja atnaujintą akto `.xlsx` failą pagal pasirinktą mėnesį.

## Funkcijos
- Palaiko `.ods` (LibreOffice), `.xlsx` ir `.xls` grafikus.
- Lietuviški mėnesių pavadinimų atpažinimas (pvz., SAUSIS/SAUSIO).
- Fuzzy atitikimas tarp grafiko ir akto eilučių pavadinimų.
- Perkelia `X` bei tekstines žymas (pvz., `1 kartas per mėn.`) į akto **Periodiškumą**.
- Jei žymoje yra skaičius, bando perrašyti **Plotas kv m./kiekis/val**.
- (Pasirinktinai) Perskaičiuoja **Kainą** = Kiekis × Įkainis, jei randa abu stulpelius.

## Naudojimas (Streamlit Cloud)
1. Įkelkite projektą į GitHub: `app.py`, `requirements.txt`, `README.md`.
2. Streamlit Cloud pasirinkite **New app** → nurodykite repo ir branch.
3. Aplikacijai užsikrovus, įkelkite `grafikas.ods` ir akto `.xls/.xlsx`, pasirinkite mėnesį ir spauskite **Atnaujinti**.

## Reikalavimai
- `streamlit`, `pandas`, `openpyxl`, `xlrd<2` (senam `.xls`), `odfpy` (`.ods`).

## Pastabos
- Jei grafike stulpeliai su mėnesiais turi sudėtinę antraštę (multi-row), aplikacija bando panaudoti pirmą eilutę kaip header.
- Jei pavadinimai labai skiriasi, gali prireikti sutvarkyti terminiją.

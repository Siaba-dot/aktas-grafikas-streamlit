
# app.py
# Streamlit app: Akto atnaujinimas iš grafiko (.ods/.xlsx/.xls)
# Author: M365 Copilot for Sigita Abasovienė

import streamlit as st
import pandas as pd
import re
import io
import tempfile
from datetime import datetime
import unicodedata

st.set_page_config(page_title="Akto atnaujinimas iš grafiko", layout="wide")

# -------- Pagalbiniai --------
LT_MONTH_ALIASES = {
    1: ['SAUSIS','SAUSIO','SAU','SAU.'],
    2: ['VASARIS','VASARIO','VAS','VAS.'],
    3: ['KOVAS','KOVO','KOV','KOV.'],
    4: ['BALANDIS','BALANDŽIO','BAL','BAL.'],
    5: ['GEGUŽĖ','GEGUŽĖS','GEG','GEG.'],
    6: ['BIRŽELIS','BIRŽELIO','BIR','BIR.'],
    7: ['LIEPA','LIEPOS','LIE','LIE.'],
    8: ['RUGPJŪTIS','RUGPJŪČIO','RGP','RGP.','RUGPJ'],
    9: ['RUGSĖJIS','RUGSĖJO','RGŠ','RGŠ.','RUGS'],
    10:['SPALIS','SPALIO','SPA','SPA.'],
    11:['LAPKritis','LAPKRIČIO','LAP','LAP.','LAPK'],
    12:['GRUODIS','GRUODŽIO','GRU','GRU.'],
}

FREQUENCY_PATTERNS = [
    (re.compile(r'\b(\d+)\s*kart(?:as|ai)\s*per\s*mėn', re.I), 'times_per_month'),
    (re.compile(r'\b(\d+[.,]?\d*)\s*kv\.m\.?', re.I), 'sqm'),
    (re.compile(r'\b(\d+[.,]?\d*)\s*vnt\.?', re.I), 'units'),
]

WEEKDAY_ALIASES = ['pirmadienis','antradienis','trečiadienis','ketvirtadienis','penktadienis','šeštadienis','sekmadienis']

def _norm(s: str) -> str:
    if s is None:
        return ''
    s = str(s)
    # Normalizuojame diakritinius brūkšnius ir kietuosius tarpus
    s = s.replace('\u2013','-').replace('\u2014','-').replace('\u00a0',' ')
    s = ' '.join(s.split())  # suvienodiname tarpus
    return s.lower()

def normalize_text(x):
    if pd.isna(x):
        return ''
    if isinstance(x, (int, float)):
        return str(x).strip()
    return str(x).strip()

def save_to_temp(uploaded_file, suffix: str):
    # Kai kurie pandas engine reikalauja kelio, todėl rašome į temp
    t = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    t.write(uploaded_file.read())
    t.flush()
    t.close()
    return t.name

# -------- Grafiko nuskaitymas --------
def parse_schedule(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith('.ods'):
        path = save_to_temp(uploaded_file, '.ods')
        # paprasčiausias būdas: read_excel su engine='odf' paima pirmą lapą
        df = pd.read_excel(path, engine='odf', header=0)
        return df, None  # sheet_name not available via simple read
    elif name.endswith('.xlsx'):
        path = save_to_temp(uploaded_file, '.xlsx')
        xls = pd.ExcelFile(path, engine='openpyxl')
        sheet = st.selectbox("Pasirinkite grafiko lapą", options=xls.sheet_names)
        df = xls.parse(sheet, header=0)
        return df, sheet
    elif name.endswith('.xls'):
        path = save_to_temp(uploaded_file, '.xls')
        xls = pd.ExcelFile(path, engine='xlrd')
        sheet = st.selectbox("Pasirinkite grafiko lapą", options=xls.sheet_names)
        df = xls.parse(sheet, header=0)
        return df, sheet
    else:
        raise ValueError('Nepalaikomas grafiko formatas. Naudokite .ods, .xlsx arba .xls')

# -------- Akto nuskaitymas --------
def load_act(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith('.xlsx'):
        path = save_to_temp(uploaded_file, '.xlsx')
        xls = pd.ExcelFile(path, engine='openpyxl')
        sheet = st.selectbox("Pasirinkite akto lapą", options=xls.sheet_names)
        df = xls.parse(sheet, header=0)
        return df, sheet
    elif name.endswith('.xls'):
        path = save_to_temp(uploaded_file, '.xls')
        xls = pd.ExcelFile(path, engine='xlrd')
        sheet = st.selectbox("Pasirinkite akto lapą", options=xls.sheet_names)
        df = xls.parse(sheet, header=0)
        return df, sheet
    else:
        raise ValueError('Nepalaikomas akto formatas. Naudokite .xls arba .xlsx')

# -------- Mėnesio (grupės) stulpelio paieška --------
def month_aliases(month:int):
    return [a.lower() for a in LT_MONTH_ALIASES.get(month, [])]

def find_month_column(df: pd.DataFrame, month: int):
    aliases = month_aliases(month)
    # 1) tiesiog per columns
    for c in df.columns:
        sc = _norm(c)
        if any(a in sc for a in aliases):
            return c
    # 2) bandymas: pirmą eilutę laikyti header (jei daugiaeilės antraštės)
    if not df.empty:
        header_candidate = [_norm(v) for v in list(df.iloc[0].values)]
        for i, hc in enumerate(header_candidate):
            if any(a in hc for a in aliases):
                return df.columns[i]
    # Jei nerasta, grąžinam None (naudotojas pasirinks rankiniu būdu)
    return None

def is_weekday_column_name(c: str) -> bool:
    sc = _norm(c)
    return any(w in sc for w in WEEKDAY_ALIASES)

# -------- Žymos sujungimas per kelis stulpelius --------
def collect_marker_from_columns(row: pd.Series, selected_cols: list):
    """
    Iš kelių stulpelių (pvz. Pirmadienis... Penktadienis) paima žymas:
    - Jei bet kur yra 'X' (nepriklausomai nuo formato), grąžina 'X'
    - Jei yra tekstinė dažnio frazė (pvz., '1 kartas per mėn.'), grąžina ją
    - Jei yra skaičiai, bandome naudoti skaičių (pvz., kv.m., vnt. ir pan.)
    Pirmenybė: X > dažnis > skaičius > tuščia
    """
    texts = []
    numbers = []
    x_flag = False

    for c in selected_cols:
        val = normalize_text(row.get(c, ''))
        if not val:
            continue
        if 'x' in val.lower():
            x_flag = True
        # match frequency text
        for rgx, kind in FREQUENCY_PATTERNS:
            m = rgx.search(val)
            if m:
                texts.append((val, kind, m.group(1)))
        # numeric marker
        try:
            n = float(val.replace(',', '.'))
            numbers.append(n)
        except:
            pass

    if x_flag:
        return 'X'

    if texts:
        # grąžinam pirmą rastą dažnio tekstą
        val, kind, num = texts[0]
        if kind == 'times_per_month':
            return f'{num} kartas per mėn.'
        return val

    if numbers:
        # grąžinam pirmą skaičių kaip tekstą
        return str(numbers[0])

    return ''

# -------- Pagrindinė atnaujinimo logika --------
def update_act_from_schedule(act_df: pd.DataFrame, schedule_df: pd.DataFrame, month_cols: list, recalc_prices: bool = False):
    # Akto stulpeliai
    name_col = act_df.columns[0]
    period_col = next((c for c in act_df.columns if 'Periodi' in normalize_text(c)), None)
    qty_col = next((c for c in act_df.columns if 'Plotas' in normalize_text(c) or 'kiekis' in normalize_text(c)), None)
    rate_col = next((c for c in act_df.columns if 'įkainis' in normalize_text(c).lower()), None)
    price_col = next((c for c in act_df.columns if 'Kaina' in normalize_text(c)), None)

    # Grafiko pagrindinis (pirmas) stulpelis – pavadinimas/užduotis
    sched_task_col = schedule_df.columns[0]

    updated_df = act_df.copy()
    change_log = []

    # Pastatome žodyną: užduotis -> marker (iš pasirinktų stulpelių)
    schedule_lookup = {}
    for _, row in schedule_df.iterrows():
        task = normalize_text(row.get(sched_task_col, ''))
        if not task:
            continue
        marker = collect_marker_from_columns(row, month_cols)
        schedule_lookup[task] = marker

    # Eilučių atnaujinimas
    for idx, row in updated_df.iterrows():
        task_name = normalize_text(row[name_col])
        if not task_name:
            continue

        # tiesioginis match
        marker = schedule_lookup.get(task_name)

        # fuzzy match
        if not marker:
            for k, v in schedule_lookup.items():
                if k and (k.lower() in task_name.lower() or task_name.lower() in k.lower()):
                    marker = v
                    break

        if not marker:
            continue

        before_period = updated_df.at[idx, period_col] if period_col else None
        before_qty = updated_df.at[idx, qty_col] if qty_col else None

        # Periodiškumas
        if period_col:
            if marker.upper() == 'X':
                updated_df.at[idx, period_col] = 'X'
            else:
                matched = False
                for rgx, kind in FREQUENCY_PATTERNS:
                    m = rgx.search(marker)
                    if m:
                        val = m.group(1)
                        updated_df.at[idx, period_col] = (
                            f'{val} kartas per mėn.' if kind == 'times_per_month' else marker
                        )
                        matched = True
                        break
                if not matched:
                    updated_df.at[idx, period_col] = marker

        # Kiekis
        if qty_col:
            try:
                num = float(marker.replace(',', '.'))
                updated_df.at[idx, qty_col] = num
            except Exception:
                pass

        # Kaina (nebūtina): qty * rate
        if recalc_prices and price_col and rate_col and qty_col:
            try:
                rate_val = updated_df.at[idx, rate_col]
                qty_val = updated_df.at[idx, qty_col]
                if pd.notna(rate_val) and pd.notna(qty_val):
                    price = float(str(rate_val).replace(',', '.')) * float(str(qty_val).replace(',', '.'))
                    updated_df.at[idx, price_col] = round(price, 2)
            except Exception:
                pass

        after_period = updated_df.at[idx, period_col] if period_col else None
        after_qty = updated_df.at[idx, qty_col] if qty_col else None
        if before_period != after_period or before_qty != after_qty:
            change_log.append({
                'Eilutė': idx,
                'Patalpa/užduotis': task_name,
                'Žyma grafike (sujungta iš pasirinktų stulpelių)': marker,
                'Periodiškumas (prieš→po)': f"{before_period} → {after_period}",
                'Kiekis (prieš→po)': f"{before_qty} → {after_qty}"
            })

    return updated_df, pd.DataFrame(change_log)

# -------- UI --------
st.title('Akto atnaujinimas iš grafiko')
st.caption('Palaiko grafikus: .ods / .xlsx / .xls. Akto formatas: .xls arba .xlsx.')

col1, col2 = st.columns(2)
with col1:
    schedule_file = st.file_uploader('Įkelkite grafiko failą (.ods/.xlsx/.xls)', type=['ods','xlsx','xls'])
    year = st.number_input('Metai', min_value=2000, max_value=2100, value=datetime.now().year)
    month_name_lt = ['Sausis','Vasaris','Kovas','Balandis','Gegužė','Birželis','Liepa','Rugpjūtis','Rugsėjis','Spalis','Lapkritis','Gruodis']
    month_idx = st.selectbox('Mėnuo', options=list(range(1,13)), format_func=lambda i: month_name_lt[i-1], index=datetime.now().month-1)

with col2:
    act_file = st.file_uploader('Įkelkite akto failą (.xls/.xlsx)', type=['xls','xlsx'])
    recalc_prices = st.checkbox('Perskaičiuoti kainą (qty × įkainis), jei stulpeliai rasti', value=False)

process = st.button('Atnaujinti aktą pagal grafiką')

if process:
    if not schedule_file or not act_file:
        st.error('Įkelkite **abu** failus: grafiką ir aktą.')
        st.stop()

    try:
        # Nuskaitome grafiką
        sched_df, sched_sheet = parse_schedule(schedule_file)
        st.subheader('Grafiko peržiūra')
        st.write("Grafiko stulpeliai:", list(map(str, sched_df.columns)))
        st.dataframe(sched_df.head(10), use_container_width=True)

        # Automatinė mėnesio kolonos paieška (jei grafike stulpelyje yra mėnesio tekstas)
        auto_month_col = find_month_column(sched_df, month_idx)
        if auto_month_col:
            st.info(f"Automatiškai rasta mėnesio kolona: **{auto_month_col}**")
        else:
            st.warning("Automatiškai mėnesio kolonos nerasta — pasirinkite rankiniu būdu.")

        # Rekomendacija: dažniausiai grafikas turi kelias kolonas (pvz., darbo dienas) po mėnesio grupės.
        st.markdown("**Pasirinkite stulpelius, kurie priklauso pasirinktam mėnesiui / mėnesių grupei** (pvz., Pirmadienis–Penktadienis po „GRUODIS–SAUSIS–VASARIS“):")
        default_selection = [auto_month_col] if auto_month_col else []
        selected_cols = st.multiselect(
            "Mėnesio/grupės stulpeliai",
            options=list(map(str, sched_df.columns)),
            default=default_selection
        )

        if not selected_cols:
            st.error("Nepasirinkote mėnesio/grupės stulpelių. Pasirinkite bent vieną.")
            st.stop()

        # Nuskaitome aktą
        act_df, act_sheet = load_act(act_file)
        st.subheader('Akto peržiūra')
        st.dataframe(act_df.head(10), use_container_width=True)

        # Atnaujinimas
        updated_df, log_df = update_act_from_schedule(act_df, sched_df, selected_cols, recalc_prices)

        st.success('Aktas atnaujintas. Žemiau matysite peržiūrą ir galėsite atsisiųsti .xlsx.')

        with st.expander('Pakeitimų žurnalas'):
            if log_df.empty:
                st.info('Pakeitimų nerasta (gal grafike pasirinktuose stulpeliuose nėra žymų arba pavadinimai nesutampa).')
            else:
                st.dataframe(log_df, use_container_width=True)

        with st.expander('Atnaujinto akto peržiūra (pirmos 20 eil.)'):
            st.dataframe(updated_df.head(20), use_container_width=True)

        # Atsisiuntimas
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            updated_df.to_excel(writer, index=False, sheet_name=act_sheet if act_sheet else 'Akto_lapas')
        output.seek(0)
        out_name = f"Aktas_atnaujintas_{year}_{month_idx:02d}.xlsx"
        st.download_button(
            label='Atsisiųsti atnaujintą aktą (.xlsx)',
            data=output,
            file_name=out_name,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        st.error(f'Klaida: {e}')
        st.exception(e)

st.markdown('---')
st.markdown('**Pastabos:**')
st.markdown('- Jei .ods neskaitomas, įsitikinkite, kad diegiant buvo pridėtas paketas `odfpy` (žr. requirements).')
st.markdown('- `.xls` skaitymui naudojama `xlrd≥2.0.1`, `.xlsx` – `openpyxl`.')
st.markdown('- Jei grafike mėnesių grupė (pvz. GRUODIS–SAUSIS–VASARIS) apima kelias kolonas (darbo dienas), pasirinkite **visas** tas kolonas – aplikacija sujungs žymas.')

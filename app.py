# app.py
# Streamlit app: Akto atnaujinimas iš grafiko (.ods/.xlsx/.xls)
# Author: M365 Copilot for Sigita Abasovienė

import streamlit as st
import pandas as pd
import re
import io
import tempfile
from datetime import datetime

st.set_page_config(page_title="Akto atnaujinimas iš grafiko", layout="wide")

# -------- Pagalbiniai --------
LT_MONTHS = {
    1: ['SAUSIS','SAUSIO','Sausis','Sausio'],
    2: ['VASARIS','VASARIO','Vasaris','Vasario'],
    3: ['KOVAS','KOVO','Kovas','Kovo'],
    4: ['BALANDIS','BALANDŽIO','Balandis','Balandžio','BALANDŽIS'],
    5: ['GEGUŽĖ','GEGUŽĖS','Gegužė','Gegužės'],
    6: ['BIRŽELIS','BIRŽELIO','Birželis','Birželio'],
    7: ['LIEPA','LIEPOS','Liepa','Liepos'],
    8: ['RUGPJŪTIS','RUGPJŪČIO','Rugpjūtis','Rugpjūčio'],
    9: ['RUGSĖJIS','RUGSĖJO','Rugsėjis','Rugsėjo'],
    10:['SPALIS','SPALIO','Spalis','Spalio'],
    11:['LAPKritis','LAPKRITIS','Lapkritis','Lapkričio'],
    12:['GRUODIS','GRUODŽIO','Gruodis','Gruodžio']
}

FREQUENCY_PATTERNS = [
    (re.compile(r'\b(\d+)\s*kart(?:as|ai)\s*per\s*mėn', re.I), 'times_per_month'),
    (re.compile(r'\b(\d+[\.,]?\d*)\s*kv\.m\.?', re.I), 'sqm'),
    (re.compile(r'\b(\d+[\.,]?\d*)\s*vnt\.?', re.I), 'units'),
]


def normalize_text(x):
    if pd.isna(x):
        return ''
    if isinstance(x, (int, float)):
        # keep integers as int-like strings to match act rows
        return str(x).strip()
    return str(x).strip()


def save_to_temp(uploaded_file, suffix: str):
    # Some pandas engines prefer file paths, so write to temp
    t = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    t.write(uploaded_file.read())
    t.flush()
    t.close()
    return t.name


def parse_schedule(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith('.ods'):
        path = save_to_temp(uploaded_file, '.ods')
        # engine='odf' requires odfpy
        df = pd.read_excel(path, engine='odf', header=0)
        return df
    elif name.endswith('.xlsx'):
        path = save_to_temp(uploaded_file, '.xlsx')
        xls = pd.ExcelFile(path, engine='openpyxl')
        sheet = xls.sheet_names[0]
        df = xls.parse(sheet, header=0)
        return df
    elif name.endswith('.xls'):
        path = save_to_temp(uploaded_file, '.xls')
        xls = pd.ExcelFile(path, engine='xlrd')
        sheet = xls.sheet_names[0]
        df = xls.parse(sheet, header=0)
        return df
    else:
        raise ValueError('Nepalaikomas grafiko formatas. Naudokite .ods, .xlsx arba .xls')


def load_act(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith('.xlsx'):
        path = save_to_temp(uploaded_file, '.xlsx')
        xls = pd.ExcelFile(path, engine='openpyxl')
        sheet = xls.sheet_names[0]
        df = xls.parse(sheet, header=0)
        return df, sheet
    elif name.endswith('.xls'):
        path = save_to_temp(uploaded_file, '.xls')
        xls = pd.ExcelFile(path, engine='xlrd')
        sheet = xls.sheet_names[0]
        df = xls.parse(sheet, header=0)
        return df, sheet
    else:
        raise ValueError('Nepalaikomas akto formatas. Naudokite .xls arba .xlsx')


def find_month_column(df: pd.DataFrame, month: int):
    # Try normal headers first
    candidates = [c for c in df.columns if any(m in normalize_text(c) for m in LT_MONTHS[month])]
    if candidates:
        return candidates[0]
    # Try multi-row header: first row as header
    if df.shape[0] > 0:
        df2 = df.copy()
        df2.columns = [normalize_text(v) for v in df2.iloc[0].values]
        df2 = df2.iloc[1:]
        for c in df2.columns:
            if any(m in normalize_text(c) for m in LT_MONTHS[month]):
                return c
    return None


def update_act_from_schedule(act_df: pd.DataFrame, schedule_df: pd.DataFrame, month: int, recalc_prices: bool = False):
    # Identify act columns
    name_col = act_df.columns[0]
    period_col = next((c for c in act_df.columns if 'Periodi' in normalize_text(c)), None)
    qty_col = next((c for c in act_df.columns if 'Plotas' in normalize_text(c) or 'kiekis' in normalize_text(c)), None)
    rate_col = next((c for c in act_df.columns if 'įkainis' in normalize_text(c).lower()), None)
    price_col = next((c for c in act_df.columns if 'Kaina' in normalize_text(c)), None)

    # Schedule columns
    sched_task_col = schedule_df.columns[0]
    month_col = find_month_column(schedule_df, month)
    if month_col is None:
        raise ValueError('Nepavyko rasti pasirinktų mėnesio stulpelio grafike.')

    # Build lookup
    schedule_lookup = {}
    for _, row in schedule_df.iterrows():
        task = normalize_text(row[sched_task_col])
        marker = normalize_text(row.get(month_col, ''))
        if task:
            schedule_lookup[task] = marker

    updated_df = act_df.copy()
    change_log = []

    for idx, row in updated_df.iterrows():
        task_name = normalize_text(row[name_col])
        if not task_name:
            continue
        marker = schedule_lookup.get(task_name)
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
                'Žyma grafike': marker,
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
    else:
        try:
            sched_df = parse_schedule(schedule_file)
            act_df, sheet_name = load_act(act_file)

            # Rodome pirmas kelias eilutes peržiūrai
            with st.expander('Grafiko peržiūra (pirmos 10 eil.)'):
                st.dataframe(sched_df.head(10))
            with st.expander('Akto peržiūra (pirmos 10 eil.)'):
                st.dataframe(act_df.head(10))

            updated_df, log_df = update_act_from_schedule(act_df, sched_df, month_idx, recalc_prices)

            st.success('Aktas atnaujintas. Žemiau matysite peržiūrą ir galėsite atsisiųsti .xlsx.')

            with st.expander('Pakeitimų žurnalas'):
                if log_df.empty:
                    st.info('Pakeitimų nerasta (gal grafike šiam mėnesiui nėra žymų arba pavadinimai nesutampa).')
                else:
                    st.dataframe(log_df)

            with st.expander('Atnaujinto akto peržiūra (pirmos 20 eil.)'):
                st.dataframe(updated_df.head(20))

            # Paruošiame atsisiuntimą
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                updated_df.to_excel(writer, index=False, sheet_name=sheet_name)
            output.seek(0)
            out_name = f"Aktas_atnaujintas_{year}_{month_idx:02d}.xlsx"
            st.download_button(label='Atsisiųsti atnaujintą aktą (.xlsx)', data=output, file_name=out_name, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        except Exception as e:
            st.error(f'Klaida: {e}')
            st.exception(e)

st.markdown('---')
st.markdown('**Pastabos:**')
st.markdown('- Jei .ods neskaitomas, įsitikinkite, kad diegiant buvo pridėtas paketas `odfpy` (žr. requirements).')
st.markdown('- .xls skaitymui naudojama `xlrd<2`.')

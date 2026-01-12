
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
    s = s.replace('\u2013','-').replace('\u2014','-').replace('\u00a0',' ')
    s = ' '.join(s.split())
    return s.lower()

def normalize_text(x):
    if pd.isna(x):
        return ''
    if isinstance(x, (int, float)):
        return str(x).strip()
    return str(x).strip()

def save_to_temp(uploaded_file, suffix: str):
    t = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    t.write(uploaded_file.read())
    t.flush()
    t.close()
    return t.name

# =========== NAUJA: grafiko skaitymas RAW (header=None) ===========
def parse_schedule_raw(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith('.ods'):
        path = save_to_temp(uploaded_file, '.ods')
        df = pd.read_excel(path, engine='odf', header=None)  # RAW
        return df, None
    elif name.endswith('.xlsx'):
        path = save_to_temp(uploaded_file, '.xlsx')
        # Pasiūlome lapo pasirinkimą
        xls = pd.ExcelFile(path, engine='openpyxl')
        sheet = st.selectbox("Pasirinkite grafiko lapą", options=xls.sheet_names)
        df = pd.read_excel(path, sheet_name=sheet, engine='openpyxl', header=None)
        return df, sheet
    elif name.endswith('.xls'):
        path = save_to_temp(uploaded_file, '.xls')
        xls = pd.ExcelFile(path, engine='xlrd')
        sheet = st.selectbox("Pasirinkite grafiko lapą", options=xls.sheet_names)
        df = pd.read_excel(path, sheet_name=sheet, engine='xlrd', header=None)
        return df, sheet
    else:
        raise ValueError('Nepalaikomas grafiko formatas. Naudokite .ods, .xlsx arba .xls')

# =========== NAUJA: suklijuojame kelių eilučių antraštes ===========
def build_headers_from_rows(raw_df: pd.DataFrame, header_rows: int):
    """
    header_rows – kiek pirmų eilučių sudaro antraštę (pvz., 2 ar 3).
    Grąžina: (df, headers) – df be antraštės eilučių, headers – suformuoti pavadinimai.
    Pavadinimą formuojame jungdami kiekvieno stulpelio viršutines header_rows vertes su ' | '.
    Pvz. "GRUODIS–SAUSIS–VASARIS | Pirmadienis"
    """
    # paimame pirmas N eilučių kaip header
    hdr = raw_df.iloc[:header_rows, :].fillna('')
    # likusi dalis – realūs duomenys
    data = raw_df.iloc[header_rows:, :].reset_index(drop=True)

    # suformuojame pavadinimus
    headers = []
    ncols = raw_df.shape[1]
    for c in range(ncols):
        parts = [normalize_text(hdr.iat[r, c]) for r in range(header_rows)]
        # pašaliname tuščias dalis
        parts = [p for p in parts if p]
        col_name = ' | '.join(parts) if parts else f'Col_{c}'
        headers.append(col_name)

    data.columns = headers
    return data, headers

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

def month_aliases(month:int):
    return [a.lower() for a in LT_MONTH_ALIASES.get(month, [])]

def find_month_columns_from_headers(headers: list, month: int):
    aliases = month_aliases(month)
    found = []
    for h in headers:
        sh = _norm(h)
        if any(a in sh for a in aliases):
            found.append(h)
    return found

def is_weekday_column_name(c: str) -> bool:
    sc = _norm(c)
    return any(w in sc for w in WEEKDAY_ALIASES)

def collect_marker_from_columns(row: pd.Series, selected_cols: list):
    texts = []
    numbers = []
    x_flag = False
    for c in selected_cols:
        val = normalize_text(row.get(c, ''))
        if not val:
            continue
        if 'x' in val.lower():
            x_flag = True
        for rgx, kind in FREQUENCY_PATTERNS:
            m = rgx.search(val)
            if m:
                texts.append((val, kind, m.group(1)))
        try:
            n = float(val.replace(',', '.'))
            numbers.append(n)
        except:
            pass
    if x_flag:
        return 'X'
    if texts:
        val, kind, num = texts[0]
        if kind == 'times_per_month':
            return f'{num} kartas per mėn.'
        return val
    if numbers:
        return str(numbers[0])
    return ''

def update_act_from_schedule(act_df: pd.DataFrame, schedule_df: pd.DataFrame, month_cols: list, recalc_prices: bool = False):
    name_col = act_df.columns[0]
    period_col = next((c for c in act_df.columns if 'Periodi' in normalize_text(c)), None)
    qty_col = next((c for c in act_df.columns if 'Plotas' in normalize_text(c) or 'kiekis' in normalize_text(c)), None)
    rate_col = next((c for c in act_df.columns if 'įkainis' in normalize_text(c).lower()), None)
    price_col = next((c for c in act_df.columns if 'Kaina' in normalize_text(c)), None)

    sched_task_col = schedule_df.columns[0]

    updated_df = act_df.copy()
    change_log = []

    schedule_lookup = {}
    for _, row in schedule_df.iterrows():
        task = normalize_text(row.get(sched_task_col, ''))
        if not task:
            continue
        marker = collect_marker_from_columns(row, month_cols)
        schedule_lookup[task] = marker

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

        if period_col:
            if marker.upper() == 'X':
                updated_df.at[idx, period_col] = 'X'
            else:
                matched = False
                for rgx, kind in FREQUENCY_PATTERNS:
                    m = rgx.search(marker)
                    if m:
                        val = m.group(1)
                        updated_df.at[idx, period_col] = (f'{val} kartas per mėn.' if kind=='times_per_month' else marker)
                        matched = True
                        break
                if not matched:
                    updated_df.at[idx, period_col] = marker

        if qty_col:
            try:
                num = float(marker.replace(',', '.'))
                updated_df.at[idx, qty_col] = num
            except Exception:
                pass

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
                'Žyma grafike (sujungta)': marker,
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
        # 1) Grafikas RAW
        raw_sched_df, sched_sheet = parse_schedule_raw(schedule_file)

        st.subheader('Grafiko žali duomenys (pirmos 10 eilučių)')
        st.dataframe(raw_sched_df.head(10), use_container_width=True)

        # 2) Kiek pirmų eilučių yra antraštės?
        header_rows = st.number_input('Kiek pirmų eilučių sudaro antraštę?', min_value=1, max_value=10, value=3, step=1)

        sched_df, headers = build_headers_from_rows(raw_sched_df, header_rows)
        st.write("Sukurtos antraštės (pirmi 40):", headers[:40])
        st.subheader('Grafiko peržiūra su naujomis antraštėmis')
        st.dataframe(sched_df.head(10), use_container_width=True)

        # 3) Automatinė paieška pagal pasirinkto mėnesio aliasus
        auto_cols = find_month_columns_from_headers(headers, month_idx)
        if auto_cols:
            st.info(f"Rastos su mėnesiu susijusios kolonos ({month_name_lt[month_idx-1]}): {auto_cols[:8]}{'...' if len(auto_cols)>8 else ''}")
        else:
            st.warning("Pagal mėnesio aliasus automatiškai nieko nerasta — pasirinkite rankiniu būdu.")

        # 4) Rankinis pasirinkimas kelių stulpelių (pvz. dienos po 'GRUODIS–SAUSIS–VASARIS')
        selected_cols = st.multiselect(
            "Pasirinkite stulpelius, kurie priklauso pasirinktai mėnesių grupei (pvz., Pirmadienis–Penktadienis).",
            options=headers,
            default=auto_cols[:10] if auto_cols else []
        )
        if not selected_cols:
            st.error("Nepasirinkote jokių stulpelių. Pasirinkite bent vieną stulpelį.")
            st.stop()

        # 5) Akto nuskaitymas
        act_df, act_sheet = load_act(act_file)
        st.subheader('Akto peržiūra')
        st.dataframe(act_df.head(10), use_container_width=True)

        # 6) Atnaujinimas
        updated_df, log_df = update_act_from_schedule(act_df, sched_df, selected_cols, recalc_prices)
        st.success('Aktas atnaujintas. Žemiau – peržiūra ir atsisiuntimas.')

        with st.expander('Pakeitimų žurnalas'):
            if log_df.empty:
                st.info('Pakeitimų nerasta (gal pasirinktuose stulpeliuose nėra žymų arba pavadinimai nesutampa).')
            else:
                st.dataframe(log_df, use_container_width=True)

        with st.expander('Atnaujinto akto peržiūra (pirmos 20 eil.)'):
            st.dataframe(updated_df.head(20), use_container_width=True)

        # 7) Atsisiuntimas
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
st.markdown('- Jei .ods neskaitomas, įsitikinkite, kad `requirements.txt` yra `odfpy>=1.4.1`.')
st.markdown('- `.xls` skaitymui naudokite `xlrd>=2.0.1`, `.xlsx` – `openpyxl`.')
st.markdown('- Jei grafiko antraštės turi kelias eilutes ir (ar) sulietas celes, padidinkite „Kiek pirmų eilučių sudaro antraštę?“ ir pasirinkite atitinkamus stulpelius.')

# app.py
# Streamlit app: Akto atnaujinimas iš grafiko (.ods/.xlsx/.xls) iš konkretaus diapazono
# Author: M365 Copilot for Sigita Abasovienė

import streamlit as st
import pandas as pd
import re
import io
import tempfile
from datetime import datetime

st.set_page_config(page_title="Akto atnaujinimas iš grafiko", layout="wide")

# -------- Iš anksto nustatyti mėnesių blokų diapazonai (A1) + mėnesių sąrašas --------
# Pastaba: mėnesiai nurodomi kaip skaičiai (1=Sausis ... 12=Gruodis)
PRESET_BLOCKS = [
    ("Lapkritis · C5:G47", "C5:G47", [11]),
    ("Gruodis–Sausis–Vasaris · H5:L47", "H5:L47", [12, 1, 2]),
    ("Kovas · M5:Q47", "M5:Q47", [3]),
    ("Balandis · R5:V47", "R5:V47", [4]),
    ("Gegužė–Birželis–Liepa–Rugpjūtis–Rugsėjis · W5:AA47", "W5:AA47", [5, 6, 7, 8, 9]),
    ("Spalis · AB5:AF47", "AB5:AF47", [10]),
]

MONTH_NAME_LT = {
    1: 'Sausis', 2: 'Vasaris', 3: 'Kovas', 4: 'Balandis', 5: 'Gegužė', 6: 'Birželis',
    7: 'Liepa', 8: 'Rugpjūtis', 9: 'Rugsėjis', 10: 'Spalis', 11: 'Lapkritis', 12: 'Gruodis'
}

# -------- Pagalbiniai --------
LT_MONTH_ALIASES = {
    1:  ['SAUSIS','SAUSIO','SAU','SAU.','SAUS'],
    2:  ['VASARIS','VASARIO','VAS','VAS.'],
    3:  ['KOVAS','KOVO','KOV','KOV.'],
    4:  ['BALANDIS','BALANDŽIO','BAL','BAL.','BALANDZIO'],
    5:  ['GEGUŽĖ','GEGUŽĖS','GEG','GEG.','GEGUZE','GEGUZES'],
    6:  ['BIRŽELIS','BIRŽELIO','BIR','BIR.','BIRZELIS','BIRZELIO'],
    7:  ['LIEPA','LIEPOS','LIE','LIE.'],
    8:  ['RUGPJŪTIS','RUGPJŪČIO','RGP','RGP.','RUGPJ','RUGP'],
    9:  ['RUGSĖJIS','RUGSĖJO','RUGS','RUGS.','RGS','RGS.','RUGSEJIS','RUGSEJO'],
    10: ['SPALIS','SPALIO','SPA','SPA.'],
    11: ['LAPKritis','LAPKRIČIO','LAPK','LAPK.','LAPKRICIO','Lapkritis','Lapkričio'],
    12: ['GRUODIS','GRUODŽIO','GRU','GRU.','GRUODZIS','GRUODZIO'],
}
WEEKDAY_ALIASES = [
    'pirmadienis','antradienis','trečiadienis','ketvirtadienis','penktadienis',
    'šeštadienis','sekmadienis','i-dienis','ii-dienis','iii-dienis','iv-dienis','v-dienis'
]
FREQUENCY_PATTERNS = [
    (re.compile(r'\b(\d+)\s*kart(?:as|ai)\s*per\s*mėn', re.I), 'times_per_month'),
    (re.compile(r'\b(\d+[.,]?\d*)\s*kv\.m\.?', re.I), 'sqm'),
    (re.compile(r'\b(\d+[.,]?\d*)\s*vnt\.?', re.I), 'units'),
]

def _norm(s: str) -> str:
    if s is None:
        return ''
    s = str(s)
    s = s.replace('\u2013','-').replace('\u2014','-').replace('\u00a0',' ')
    s = ' '.join(s.split())
    return s.lower()

def normalize_text(x):
    if pd.isna(x): return ''
    if isinstance(x, (int, float)): return str(x).strip()
    return str(x).strip()

def save_to_temp(uploaded_file, suffix: str):
    t = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    t.write(uploaded_file.read())
    t.flush()
    t.close()
    return t.name

# -------- A1 diapazono parsinimas --------
def col_letter_to_index(col: str) -> int:
    col = re.sub(r'[^A-Za-z]', '', col).upper()
    val = 0
    for ch in col:
        val = val * 26 + (ord(ch) - ord('A') + 1)
    return val - 1

def parse_a1_cell(cell: str):
    m = re.match(r'([A-Za-z]+)(\d+)', cell.strip())
    if not m:
        raise ValueError(f"Neteisingas A1 adresas: {cell}")
    col = col_letter_to_index(m.group(1))
    row = int(m.group(2)) - 1
    return row, col

def parse_a1_range(a1: str, shape):
    parts = a1.split(':')
    if len(parts) != 2:
        raise ValueError("Diapazoną nurodykite A1 formatu, pvz.: G3:N80")
    r0, c0 = parse_a1_cell(parts[0])
    r1, c1 = parse_a1_cell(parts[1])
    r0, r1 = sorted([r0, r1])
    c0, c1 = sorted([c0, c1])
    max_r, max_c = shape[0]-1, shape[1]-1
    r0 = max(0, min(r0, max_r)); r1 = max(0, min(r1, max_r))
    c0 = max(0, min(c0, max_c)); c1 = max(0, min(c1, max_c))
    return r0, r1, c0, c1

# ---------- RAW grafiko nuskaitymas (be antraščių) ----------
def parse_schedule_raw(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith('.ods'):
        path = save_to_temp(uploaded_file, '.ods')
        df = pd.read_excel(path, engine='odf', header=None)  # RAW
        return df, None
    elif name.endswith('.xlsx'):
        path = save_to_temp(uploaded_file, '.xlsx')
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

# ---------- Header wizard: ffill+bfill per sulietas celes ----------
def build_headers(raw_df: pd.DataFrame, start_row: int, header_rows_up: int):
    top = raw_df.iloc[start_row-header_rows_up:start_row+1, :].copy().fillna('')
    top_ff = top.T.ffill().T
    top_full = top_ff.T.bfill().T
    data = raw_df.iloc[start_row+1:, :].reset_index(drop=True)
    headers = []
    for c in range(top_full.shape[1]):
        parts = [normalize_text(top_full.iat[r, c]) for r in range(top_full.shape[0])]
        parts = [p for p in parts if p]
        col_name = ' | '.join(parts) if parts else f'Col_{c}'
        headers.append(col_name)
    data.columns = headers
    return data, headers

# ---------- Pagalbinės paieškos ----------
def month_aliases(month:int):
    return [a.lower() for a in LT_MONTH_ALIASES.get(month, [])]

def _tokens(s: str) -> list:
    s = _norm(s)
    return [t for t in re.split(r'[^a-z0-9ąčęėįšųūž]+', s) if t]

def find_month_columns_from_headers(headers: list, month: int):
    aliases = month_aliases(month)
    found = []
    for h in headers:
        toks = _tokens(h)
        if any(a in toks for a in aliases):
            found.append(h)
    return found

# ---------- Akto nuskaitymas ----------
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

# ---------- Žymų sujungimas iš kelių stulpelių ----------
def collect_marker_from_columns(row: pd.Series, selected_cols: list):
    texts, numbers, x_flag = [], [], False
    for c in selected_cols:
        val = normalize_text(row.get(c, ''))
        if not val: continue
        if 'x' in val.lower(): x_flag = True
        for rgx, kind in FREQUENCY_PATTERNS:
            m = rgx.search(val)
            if m: texts.append((val, kind, m.group(1)))
        try:
            n = float(val.replace(',', '.'))
            numbers.append(n)
        except: pass
    if x_flag: return 'X'
    if texts:
        val, kind, num = texts[0]
        return f'{num} kartas per mėn.' if kind == 'times_per_month' else val
    if numbers: return str(numbers[0])
    return ''

# ---------- Pagrindinė atnaujinimo logika ----------
def update_act_from_schedule(act_df: pd.DataFrame, schedule_df: pd.DataFrame, month_cols: list, recalc_prices: bool = False):
    name_col = act_df.columns[0]
    period_col = next((c for c in act_df.columns if 'Periodi' in normalize_text(c)), None)
    qty_col    = next((c for c in act_df.columns if 'Plotas' in normalize_text(c) or 'kiekis' in normalize_text(c)), None)
    rate_col   = next((c for c in act_df.columns if 'įkainis' in normalize_text(c).lower()), None)
    price_col  = next((c for c in act_df.columns if 'Kaina'  in normalize_text(c)), None)
    sched_task_col = schedule_df.columns[0]

    updated_df = act_df.copy()
    change_log = []
    schedule_lookup = {}

    for _, row in schedule_df.iterrows():
        task = normalize_text(row.get(sched_task_col, ''))
        if not task: continue
        marker = collect_marker_from_columns(row, month_cols)
        schedule_lookup[task] = marker

    for idx, row in updated_df.iterrows():
        task_name = normalize_text(row[name_col])
        if not task_name: continue

        marker = schedule_lookup.get(task_name)
        if not marker:
            for k, v in schedule_lookup.items():
                if k and (k.lower() in task_name.lower() or task_name.lower() in k.lower()):
                    marker = v; break
        if not marker: continue

        before_period = updated_df.at[idx, period_col] if period_col else None
        before_qty    = updated_df.at[idx, qty_col]    if qty_col    else None

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
                        matched = True; break
                if not matched:
                    updated_df.at[idx, period_col] = marker

        if qty_col:
            try:
                updated_df.at[idx, qty_col] = float(marker.replace(',', '.'))
            except: pass

        if recalc_prices and price_col and rate_col and qty_col:
            try:
                rate_val = updated_df.at[idx, rate_col]
                qty_val  = updated_df.at[idx, qty_col]
                if pd.notna(rate_val) and pd.notna(qty_val):
                    price = float(str(rate_val).replace(',', '.')) * float(str(qty_val).replace(',', '.'))
                    updated_df.at[idx, price_col] = round(price, 2)
            except: pass

        after_period = updated_df.at[idx, period_col] if period_col else None
        after_qty    = updated_df.at[idx, qty_col]    if qty_col    else None
        if before_period != after_period or before_qty != after_qty:
            change_log.append({
                'Eilutė': idx,
                'Patalpa/užduotis': task_name,
                'Žyma grafike (sujungta)': marker,
                'Periodiškumas (prieš→po)': f"{before_period} → {after_period}",
                'Kiekis (prieš→po)': f"{before_qty} → {after_qty}"
            })

    return updated_df, pd.DataFrame(change_log)

# ---------- UI ----------
st.title('Akto atnaujinimas iš grafiko (konkretus diapazonas)')
st.caption('Grafikas: .ods / .xlsx / .xls. Aktas: .xls / .xlsx.')

# Įkėlimas
schedule_file = st.file_uploader('Įkelkite grafiko failą (.ods/.xlsx/.xls)', type=['ods','xlsx','xls'])
act_file = st.file_uploader('Įkelkite akto failą (.xls/.xlsx)', type=['xls','xlsx'])
recalc_prices = st.checkbox('Perskaičiuoti kainą (qty × įkainis), jei stulpeliai rasti', value=False)

if schedule_file and act_file:
    try:
        # RAW grafikas
        raw_sched_df, sched_sheet = parse_schedule_raw(schedule_file)
        st.caption(f"Grafiko forma: {raw_sched_df.shape[0]} eil. × {raw_sched_df.shape[1]} kol.")
        st.dataframe(raw_sched_df.head(20), use_container_width=True)

        # Diapazono pasirinkimas iš presetų + mėnesio pasirinkimas pagal bloką
        st.markdown("### Pasirinkite mėnesio bloką")
        preset_label = st.selectbox(
            "Mėnesio blokas (A1)",
            options=[p[0] for p in PRESET_BLOCKS], index=1  # default: Gruodis–Sausis–Vasaris
        )
        preset_map = {p[0]: (p[1], p[2]) for p in PRESET_BLOCKS}
        a1_default, months_in_block = preset_map[preset_label]

        # --- Numatytas mėnuo pagal pasirinktą bloką ---
        if 'last_preset' not in st.session_state:
            st.session_state.last_preset = None
            st.session_state.month_default_idx = 0

        # Jei pasikeitė blokas, nustatome numatytą mėnesį (pirmą sąraše)
        if st.session_state.last_preset != preset_label:
            st.session_state.month_default_idx = 0
            st.session_state.last_preset = preset_label

        month_choice = st.selectbox(
            'Mėnuo šiame bloke',
            options=months_in_block,
            index=st.session_state.month_default_idx,
            format_func=lambda m: MONTH_NAME_LT[m],
            key='month_choice_select'
        )

        a1 = st.text_input('A1 diapazonas (galite pakoreguoti ranka)', value=a1_default)
        r0, r1, c0, c1 = parse_a1_range(a1.strip(), raw_sched_df.shape)
        block = raw_sched_df.iloc[int(r0):int(r1)+1, int(c0):int(c1)+1].reset_index(drop=True)
        st.subheader('Pasirinktas grafiko blokas (peržiūra)')
        st.dataframe(block.head(20), use_container_width=True)

        # Header wizard
        guess_row = 0
        for i in range(min(40, block.shape[0])):
            vals = [normalize_text(v) for v in list(block.iloc[i, :].values)]
            score = sum(1 for v in vals if any(w in _norm(v) for w in WEEKDAY_ALIASES))
            if score >= 3:
                guess_row = i; break
        start_row = st.number_input('Eilutė bloke, kur yra bazinės antraštės (pvz., Pirmadienis…)', min_value=0, max_value=int(block.shape[0]-2), value=int(guess_row), step=1)
        header_rows_up = st.number_input('Kiek eilučių virš jos – „grupės“ antraštė', min_value=0, max_value=10, value=1, step=1)

        sched_df, headers = build_headers(block, start_row, header_rows_up)
        st.write("Sukurtos antraštės (pirmi 40):", headers[:40])
        st.dataframe(sched_df.head(12), use_container_width=True)

        # Automatinė paieška pagal parinktą mėnesį
        auto_cols = find_month_columns_from_headers(headers, month_choice)
        if auto_cols:
            st.info(f"Rastos su {MONTH_NAME_LT[month_choice]} susijusios kolonos: {auto_cols[:10]}{'...' if len(auto_cols)>10 else ''}")
        else:
            st.warning("Pagal mėnesio aliasus automatiškai nieko nerasta — pasirinkite rankiniu būdu.")

        substr = st.text_input("Papildomas rankinis filtras antraštėms (pvz., 'gruodis'/'sausis'...)", value='')
        manual_hits = []
        if substr.strip():
            ss = _norm(substr)
            manual_hits = [h for h in headers if ss in _norm(h)]
            if manual_hits:
                st.success(f"Pagal filtrą rasta: {manual_hits[:10]}{'...' if len(manual_hits)>10 else ''}")
            else:
                st.warning("Pagal filtrą nieko nerasta.")

        preselect = (auto_cols or manual_hits)[:10]
        selected_cols = st.multiselect(
            "Pasirinkite stulpelius (pvz., Pirmadienis–Penktadienis)",
            options=headers,
            default=preselect
        )
        if not selected_cols:
            st.error("Nepasirinkote jokių stulpelių. Pasirinkite bent vieną.")
            st.stop()

        # Akto nuskaitymas
        act_df, act_sheet = load_act(act_file)
        st.subheader('Akto peržiūra')
        st.dataframe(act_df.head(10), use_container_width=True)

        # Atnaujinimas
        updated_df, log_df = update_act_from_schedule(act_df, sched_df, selected_cols, recalc_prices)
        st.success('Aktas atnaujintas. Žemiau – peržiūra ir atsisiuntimas.')

        with st.expander('Pakeitimų žurnalas'):
            if log_df.empty:
                st.info('Pakeitimų nerasta (gal pasirinktuose stulpeliuose nėra žymų arba pavadinimai nesutampa).')
            else:
                st.dataframe(log_df, use_container_width=True)

        with st.expander('Atnaujinto akto peržiūra (pirmos 20 eil.)'):
            st.dataframe(updated_df.head(20), use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            updated_df.to_excel(writer, index=False, sheet_name=act_sheet if act_sheet else 'Akto_lapas')
        output.seek(0)
        out_name = f"Aktas_atnaujintas_{datetime.now().year}_{month_choice:02d}.xlsx"
        st.download_button(
            label='Atsisiųsti atnaujintą aktą (.xlsx)',
            data=output,
            file_name=out_name,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        st.error(f'Klaida: {e}')
        st.exception(e)
else:
    st.info('Įkelkite grafiką ir aktą, tada pasirinkite bloką ir mėnesį.')

st.markdown('---')
st.markdown('**Pastabos:**')
st.markdown('- `.ods` skaitymui: `odfpy`; `.xls`: `xlrd≥2.0.1`; `.xlsx`: `openpyxl`.')
st.markdown('- Pasirinkite bloką iš sąrašo (pvz., H5:L47) ir atitinkamą mėnesį tame bloke.')
st.markdown('- Jei antraštės sulietos, naudokite ffill/bfill wizard’ą (įtraukta šiame kode).')

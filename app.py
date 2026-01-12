
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
    """A -> 0, B -> 1, ... Z -> 25, AA -> 26, AB -> 27, etc."""
    col = re.sub(r'[^A-Za-z]', '', col).upper()
    val = 0
    for ch in col:
        val = val * 26 + (ord(ch) - ord('A') + 1)
    return val - 1

def parse_a1_cell(cell: str):
    """G3 -> (row_idx=2, col_idx=6) (0-based)"""
    m = re.match(r'([A-Za-z]+)(\d+)', cell.strip())
    if not m:
        raise ValueError(f"Neteisingas A1 adresas: {cell}")
    col = col_letter_to_index(m.group(1))
    row = int(m.group(2)) - 1
    return row, col

def parse_a1_range(a1: str, shape):
    """G3:N80 -> (r0,r1,c0,c1) inclusive (0-based), clipped to df shape."""
    parts = a1.split(':')
    if len(parts) != 2:
        raise ValueError("Diapazoną nurodykite A1 formatu, pvz.: G3:N80")
    r0, c0 = parse_a1_cell(parts[0])
    r1, c1 = parse_a1_cell(parts[1])
    # normalizuojam tvarką
    r0, r1 = sorted([r0, r1])
    c0, c1 = sorted([c0, c1])
    # ribojam pagal df
    max_r, max_c = shape[0]-1, shape[1]-1
    r0 = max(0, min(r0, max_r))
    r1 = max(0, min(r1, max_r))
    c0 = max(0, min(c0, max_c))
    c1 = max(0, min(c1, max_c))
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
    """
    start_row – eilutė, kur yra bazinės stulpelių antraštės (pvz., Pirmadienis…).
    header_rows_up – kiek eilučių virš start_row įtraukti į antraštę (pvz., 1 – GRUODIS–SAUSIS–VASARIS).
    """
    top = raw_df.iloc[start_row-header_rows_up:start_row+1, :].copy().fillna('')
    # ffill (į dešinę) ir bfill (į kairę) per transponavimą (kad atkartotų merged)
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
        # 1) RAW grafikas (be antraščių)
        raw_sched_df, sched_sheet = parse_schedule_raw(schedule_file)
        st.subheader('Grafiko RAW (pirmos 30 eilučių)')
        st.dataframe(raw_sched_df.head(30), use_container_width=True)

        # 2) Diapazono pasirinkimas (A1 arba skaitinis)
        st.markdown("### Diapazonas grafike")
        a1 = st.text_input("A1 diapazonas (pvz., G3:N80). Jei nenaudosi A1, užpildyk skaitinius laukus žemiau.", value="")
        r0=r1=c0=c1=None

        if a1.strip():
            try:
                r0, r1, c0, c1 = parse_a1_range(a1.strip(), raw_sched_df.shape)
            except Exception as e:
                st.error(f"Neteisingas diapazonas: {e}")
                st.stop()
        else:
            st.write("Arba nurodyk skaitinį diapazoną (0-based indeksai):")
            c0 = st.number_input("Pradžios kolona (0=A, 1=B...)", min_value=0, max_value=int(raw_sched_df.shape[1]-1), value=6)
            c1 = st.number_input("Pabaigos kolona", min_value=0, max_value=int(raw_sched_df.shape[1]-1), value=13)
            r0 = st.number_input("Pradžios eilutė", min_value=0, max_value=int(raw_sched_df.shape[0]-1), value=2)
            r1 = st.number_input("Pabaigos eilutė", min_value=0, max_value=int(raw_sched_df.shape[0]-1), value=80)

        # Apkarpome grafiko bloką
        block = raw_sched_df.iloc[int(r0):int(r1)+1, int(c0):int(c1)+1].reset_index(drop=True)
        st.subheader('Pasirinktas grafiko blokas (peržiūra)')
        st.dataframe(block.head(20), use_container_width=True)

        # 3) Header wizard šiame bloke: start_row ir virš jos esantis „grupės“ sluoksnis
        # Automatinis spėjimas: eilutė su >=3 savaitės dienomis
        guess_row = 0
        for i in range(min(40, block.shape[0])):
            vals = [normalize_text(v) for v in list(block.iloc[i, :].values)]
            score = sum(1 for v in vals if any(w in _norm(v) for w in WEEKDAY_ALIASES))
            if score >= 3:
                guess_row = i; break

        start_row = st.number_input('Eilutė bloke, kur yra bazinės antraštės (pvz., Pirmadienis…)', min_value=0, max_value=int(block.shape[0]-2), value=int(guess_row), step=1)
        header_rows_up = st.number_input('Kiek eilučių virš jos – „grupės“ antraštė (pvz., GRUODIS–SAUSIS–VASARIS)', min_value=0, max_value=10, value=1, step=1)

        sched_df, headers = build_headers(block, start_row, header_rows_up)
        st.write("Sukurtos antraštės (pirmi 40):", headers[:40])
        st.subheader('Grafiko peržiūra su naujomis antraštėmis (bloke)')
        st.dataframe(sched_df.head(12), use_container_width=True)

        # 4) Automatiškai randam kolonas, kuriose minima pasirinkto mėnesio alias
        auto_cols = find_month_columns_from_headers(headers, month_idx)
        if auto_cols:
            st.info(f"Rastos su mėnesiu susijusios kolonos ({month_name_lt[month_idx-1]}): {auto_cols[:10]}{'...' if len(auto_cols)>10 else ''}")
        else:
            st.warning("Pagal mėnesio aliasus automatiškai nieko nerasta — pasirinkite rankiniu būdu.")

        # 5) Rankinis „substring“ filtras (pvz., 'gruodis-sausis-vasaris' arba 'gruodis')
        substr = st.text_input("Papildomas rankinis filtras antraštėms (pvz., 'gruodis-sausis-vasaris' arba 'gruodis')", value='')
        manual_hits = []
        if substr.strip():
            ss = _norm(substr)
            manual_hits = [h for h in headers if ss in _norm(h)]
            if manual_hits:
                st.success(f"Pagal filtrą rasta: {manual_hits[:10]}{'...' if len(manual_hits)>10 else ''}")
            else:
                st.warning("Pagal filtrą nieko nerasta.")

        # 6) Pasirenkame mėnesio/grupės kolonas (pvz., Pirmadienis–Penktadienis po „GRUODIS–SAUSIS–VASARIS“)
        preselect = (auto_cols or manual_hits)[:10]
        selected_cols = st.multiselect(
            "Pasirinkite stulpelius, kurie priklauso pasirinktai mėnesių grupei (šiame bloke).",
            options=headers,
            default=preselect
        )
        if not selected_cols:
            st.error("Nepasirinkote jokių stulpelių. Pasirinkite bent vieną.")
            st.stop()

        # 7) Akto nuskaitymas
        act_df, act_sheet = load_act(act_file)
        st.subheader('Akto peržiūra')
        st.dataframe(act_df.head(10), use_container_width=True)

        # 8) Atnaujinimas
        updated_df, log_df = update_act_from_schedule(act_df, sched_df, selected_cols, recalc_prices)
        st.success('Aktas atnaujintas. Žemiau – peržiūra ir atsisiuntimas.')

        with st.expander('Pakeitimų žurnalas'):
            if log_df.empty:
                st.info('Pakeitimų nerasta (gal pasirinktuose stulpeliuose nėra žymų arba pavadinimai nesutampa).')
            else:
                st.dataframe(log_df, use_container_width=True)

        with st.expander('Atnaujinto akto peržiūra (pirmos 20 eil.)'):
            st.dataframe(updated_df.head(20), use_container_width=True)

        # 9) Atsisiuntimas
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
st.markdown('- `.ods` skaitymui: `odfpy`; `.xls`: `xlrd≥2.0.1`; `.xlsx`: `openpyxl`.')
st.markdown('- Diapazonas A1 (pvz., `G3:N80`) patogus, kai grafiko sekcija yra tik dalyje lapo.')
st.markdown('- Jei antraštės sulietos, naudokite ffill+bfill wizard’ą (šiame kode jau įdiegta).')

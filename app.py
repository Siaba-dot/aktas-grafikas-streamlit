
# app.py
# Streamlit: Akto atnaujinimas iš grafiko (.ods/.xlsx/.xls) su diapazonais,
# su „header wizard“ (sulietų antraščių ffill+bfill), mėnesių blokais,
# ir darbo dienų skaičiavimu (konkretiems metams/mėnesiui), su kintamų švenčių korekcija.

import streamlit as st
import pandas as pd
import re
import io
import tempfile
import calendar
import datetime
from datetime import datetime as dt, date, timedelta

st.set_page_config(page_title="Akto atnaujinimas iš grafiko", layout="wide")

# ========= PRESET BLOKAI (A1 diapazonai + mėnesių sąrašai) =========
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

# ========= Pagalbinės teksto funkcijos =========
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

# ========= Failų nuskaitymas =========
def save_to_temp(uploaded_file, suffix: str):
    t = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    t.write(uploaded_file.read())
    t.flush()
    t.close()
    return t.name

# A1 adresų parsinimas
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
        raise ValueError("Diapazoną nurodykite A1 formatu, pvz.: C5:G47")
    r0, c0 = parse_a1_cell(parts[0])
    r1, c1 = parse_a1_cell(parts[1])
    r0, r1 = sorted([r0, r1])
    c0, c1 = sorted([c0, c1])
    max_r, max_c = shape[0]-1, shape[1]-1
    r0 = max(0, min(r0, max_r)); r1 = max(0, min(r1, max_r))
    c0 = max(0, min(c0, max_c)); c1 = max(0, min(c1, max_c))
    return r0, r1, c0, c1

# RAW grafiko nuskaitymas (be header)
def parse_schedule_raw(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith('.ods'):
        path = save_to_temp(uploaded_file, '.ods')
        df = pd.read_excel(path, engine='odf', header=None)
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

# Header wizard: ffill+bfill per sulietas celes (formuoja prasmingus pavadinimus)
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

# Paieška pagal mėnesio aliasus
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

# Akto nuskaitymas
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

# Žymų sujungimas iš kelių stulpelių (X/tekstai/skaičiai)
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

# Pagrindinė atnaujinimo logika
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

# ========= Kintamų švenčių skaičiavimas ir darbo dienos =========
# Velykos (Vakarų/Gregoriano kalendoriaus) – Meeus/Jones/Butcher algoritmas
def easter_sunday(year: int) -> date:
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2*e + 2*i - h - k) % 7
    m = (a + 11*h + 22*l) // 451
    month = (h + l - 7*m + 114) // 31
    day = ((h + l - 7*m + 114) % 31) + 1
    return date(year, month, day)

def first_sunday(year: int, month: int) -> date:
    d = date(year, month, 1)
    while d.weekday() != 6:  # 0=Mon ... 6=Sun
        d += timedelta(days=1)
    return d

def lt_fixed_holidays(year: int) -> set:
    return {
        date(year, 1, 1),   # Naujieji metai
        date(year, 2, 16),  # Valstybės atkūrimo diena
        date(year, 3, 11),  # Nepriklausomybės atkūrimo diena
        date(year, 5, 1),   # Darbo diena
        date(year, 6, 24),  # Joninės
        date(year, 7, 6),   # Valstybės diena
        date(year, 8, 15),  # Žolinė
        date(year, 11, 1),  # Visų šventųjų
        date(year, 12, 24), # Kūčios
        date(year, 12, 25), # Kalėdos I
        date(year, 12, 26), # Kalėdos II
    }

def lt_variable_holidays(year: int,
                         include_easter_sunday=True,
                         include_easter_monday=True,
                         include_pentecost_sunday=True,
                         include_mothers_day=True,
                         include_fathers_day=True) -> set:
    es = easter_sunday(year)
    em = es + timedelta(days=1)
    ps = es + timedelta(days=49)  # Sekminės (sekmadienis)
    md = first_sunday(year, 5)
    fd = first_sunday(year, 6)
    s = set()
    if include_easter_sunday:  s.add(es)
    if include_easter_monday:  s.add(em)
    if include_pentecost_sunday: s.add(ps)
    if include_mothers_day:    s.add(md)
    if include_fathers_day:    s.add(fd)
    return s

def darbo_dienos(year: int, month: int, holidays: set) -> int:
    cal = calendar.Calendar()
    count = 0
    for d in cal.itermonthdates(year, month):
        if d.month == month and d.weekday() < 5 and d not in holidays:
            count += 1
    return count

def holidays_in_month(year: int, month: int, holidays: set) -> list:
    """Grąžina šventines, kurios patenka į pasirinktą mėnesį."""
    return sorted([h for h in holidays if h.month == month])

# ========= UI =========
st.title('Akto atnaujinimas iš grafiko (diapazonai + darbo dienos)')
st.caption('Grafikas: .ods / .xlsx / .xls. Aktas: .xls / .xlsx.')

schedule_file = st.file_uploader('Įkelkite grafiko failą (.ods/.xlsx/.xls)', type=['ods','xlsx','xls'])
act_file      = st.file_uploader('Įkelkite akto failą (.xls/.xlsx)', type=['xls','xlsx'])
recalc_prices = st.checkbox('Perskaičiuoti kainą (qty × įkainis), jei stulpeliai rasti', value=False)

if schedule_file and act_file:
    try:
        # RAW grafikas
        raw_sched_df, sched_sheet = parse_schedule_raw(schedule_file)
        st.caption(f"Grafiko forma: {raw_sched_df.shape[0]} eil. × {raw_sched_df.shape[1]} kol.")
        st.dataframe(raw_sched_df.head(20), use_container_width=True)

        # Pasirink bloką + mėnesį + metus
        st.markdown("### Pasirinkite mėnesio bloką ir konkretų mėnesį")
        preset_label = st.selectbox(
            "Mėnesio blokas (A1)",
            options=[p[0] for p in PRESET_BLOCKS],
            index=1  # default: Gruodis–Sausis–Vasaris
        )
        preset_map = {p[0]: (p[1], p[2]) for p in PRESET_BLOCKS}
        a1_default, months_in_block = preset_map[preset_label]

        # Metai
        year_choice = st.number_input("Metai", min_value=2000, max_value=2100, value=dt.now().year)

        # Default mėnuo: einamasis, jei yra bloke, kitaip pirmas bloke
        now_month = dt.now().month
        default_idx = months_in_block.index(now_month) if now_month in months_in_block else 0

        month_choice = st.selectbox(
            "Konkretaus mėnesio pasirinkimas šiame bloke",
            options=months_in_block,
            index=default_idx,
            format_func=lambda m: MONTH_NAME_LT[m]
        )

        # A1 diapazonas (galima koreguoti ranka)
        a1 = st.text_input("A1 diapazonas (galite pakoreguoti ranka)", value=a1_default)

        # Iškarpytas blokas
        r0, r1, c0, c1 = parse_a1_range(a1.strip(), raw_sched_df.shape)
        block = raw_sched_df.iloc[int(r0):int(r1)+1, int(c0):int(c1)+1].reset_index(drop=True)
        st.subheader('Pasirinktas grafiko blokas (peržiūra)')
        st.dataframe(block.head(20), use_container_width=True)

        # Header wizard šiame bloke
        guess_row = 0
        for i in range(min(40, block.shape[0])):
            vals = [normalize_text(v) for v in list(block.iloc[i, :].values)]
            score = sum(1 for v in vals if any(w in _norm(v) for w in WEEKDAY_ALIASES))
            if score >= 3:
                guess_row = i; break
        start_row = st.number_input('Eilutė bloke, kur yra bazinės antraštės (pvz., Pirmadienis…)',
                                    min_value=0, max_value=int(block.shape[0]-2), value=int(guess_row), step=1)
        header_rows_up = st.number_input('Kiek eilučių virš jos – „grupės“ antraštė',
                                         min_value=0, max_value=10, value=1, step=1)

        sched_df, headers = build_headers(block, start_row, header_rows_up)
        st.write("Sukurtos antraštės (pirmi 40):", headers[:40])
        st.dataframe(sched_df.head(12), use_container_width=True)

        # Automatinė paieška pagal konkretų mėnesį
        auto_cols = find_month_columns_from_headers(headers, month_choice)
        if auto_cols:
            st.info(f"Rastos su {MONTH_NAME_LT[month_choice]} susijusios kolonos: {auto_cols[:10]}{'...' if len(auto_cols)>10 else ''}")
        else:
            st.warning("Pagal mėnesio aliasus automatiškai nieko nerasta — pasirinkite rankiniu būdu.")

        # Papildomas substring filtras + multiselect
        substr = st.text_input("Papildomas rankinis filtras antraštėms (pvz., 'gruodis'/'sausis'...)",
                               value='')
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

        # ==== ŠVENTINĖS (FIKSUOTOS + KINTAMOS) IR DARBO DIENOS ====
        st.markdown("### Šventinės dienos ir darbo dienų skaičiavimas")

        # Fiksuotos (pagal metus)
        fixed_h = lt_fixed_holidays(year_choice)
        with st.expander("Fiksuotos LT šventinės (siūlomos)", expanded=False):
            st.write(sorted(list(fixed_h)))

        # Kintamos: įtraukti/neįtraukti (checkbox)
        st.caption("Kintamos šventinės – automatinis paskaičiavimas (Velykos, Sekminės, Motinos/Tėvo diena):")
        include_easter_sunday  = st.checkbox("Įtraukti Velykų sekmadienį", value=True)
        include_easter_monday  = st.checkbox("Įtraukti Velykų pirmadienį (darbo dienoms turi įtaką)", value=True)
        include_pentecost_sun  = st.checkbox("Įtraukti Sekminių sekmadienį", value=True)
        include_mothers_day    = st.checkbox("Įtraukti Motinos dieną (sekmadienis)", value=True)
        include_fathers_day    = st.checkbox("Įtraukti Tėvo dieną (sekmadienis)", value=True)

        variable_h = lt_variable_holidays(
            year_choice,
            include_easter_sunday,
            include_easter_monday,
            include_pentecost_sun,
            include_mothers_day,
            include_fathers_day
        )
        with st.expander("Kintamos šventinės (auto)", expanded=False):
            st.write(sorted(list(variable_h)))

        # Rankinis pridėjimas/šalinimas (CSV įvedimas)
        st.caption("Papildomos šventinės (rankiniu būdu): įveskite YYYY-MM-DD per kablelį.")
        extra_add_text = st.text_input("Pridėti datas", value="")
        extra_del_text = st.text_input("Pašalinti datas", value="")

        extra_add_set, extra_del_set = set(), set()
        def parse_dates_csv(csv_text: str) -> set:
            s = set()
            if csv_text.strip():
                for part in csv_text.split(','):
                    tok = part.strip()
                    if not tok: continue
                    try:
                        y, m, d = map(int, tok.split('-'))
                        s.add(date(y, m, d))
                    except:
                        st.warning(f"Praleista neteisinga data: {tok}")
            return s

        extra_add_set = parse_dates_csv(extra_add_text)
        extra_del_set = parse_dates_csv(extra_del_text)

        # Sudedame galutinį rinkinį
        holidays_final = (fixed_h.union(variable_h).union(extra_add_set)).difference(extra_del_set)
        holidays_month = holidays_in_month(year_choice, month_choice, holidays_final)

        # Darbo dienos
        wd = darbo_dienos(year_choice, month_choice, holidays_final)
        st.metric(f"Darbo dienos {MONTH_NAME_LT[month_choice]} {year_choice} (be šventinių)", wd)

        with st.expander("Šventinės, patenkančios į pasirinktą mėnesį", expanded=False):
            if holidays_month:
                st.write(holidays_month)
            else:
                st.info("Šiame mėnesyje šventinių (po korekcijos) nėra arba jos nepatenka į darbo dienas.")

        # ==== Akto atnaujinimas ir atsisiuntimas ====
        act_df, act_sheet = load_act(act_file)
        st.subheader('Akto peržiūra')
        st.dataframe(act_df.head(10), use_container_width=True)

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
        out_name = f"Aktas_atnaujintas_{year_choice}_{month_choice:02d}.xlsx"
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
    st.info('Įkelkite grafiką ir aktą, tada pasirinkite bloką, konkretų mėnesį ir metus.')

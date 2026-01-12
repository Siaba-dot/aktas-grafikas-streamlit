
# app.py
# Streamlit: Akto atnaujinimas iš grafiko (.ods/.xlsx/.xls) su WOW dizainu,
# teisingais savaitės dienų skaičiavimais (pagal pasirinktus metus + šventines),
# ir Excel formulių įrašymu:
# - C7: "{MĖNESIO KILM.} 1-{paskutinė}" (pvz., "BALANDŽIO 1-30")
# - A6: paskutinė mėnesio data (YYYY-MM-DD)
# - B8: fiksuotas "Plotas kv m./kiekis/val"
# - Skaičiavimai: Pir..Sek skaičiai (atsižvelgiant į šventines)
# - Periodiškumas: IF per "X" žymas + sumos iš "Skaičiavimai"
# - Kaina: Plotas * Įkainis * Periodiškumas (tarpinių neapvaliname; tik rezultatui formatas 0.00)
# - PVM neskaičiuojamas; pridedama "Iš viso (be PVM)"
# - Sekminės: laikomos "darbo diena" (neįtraukiamos į šventines)

import streamlit as st
import pandas as pd
import re
import io
import tempfile
import calendar
from datetime import datetime as dt, date, timedelta

# ------------------ PUSLAPIO NUSTATYMAI + WOW CSS ------------------

st.set_page_config(page_title="Akto atnaujinimas iš grafiko", layout="wide")

def inject_wow_css(accent="#7C3AED"):
    st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;800&display=swap');
    html, body, [class*="css"]  {{ font-family: 'Inter', system-ui, -apple-system, 'Segoe UI', Roboto, Ubuntu, 'Helvetica Neue', Arial, 'Noto Sans', sans-serif; }}
    .stApp {{
      background: radial-gradient(1200px 600px at 10% 10%, rgba(124,58,237,0.08), transparent 55%),
                  radial-gradient(800px 400px at 90% 20%, rgba(20,184,166,0.06), transparent 60%),
                  linear-gradient(120deg, #0B1221 0%, #0B1221 100%);
    }}
    .hero {{
      padding: 1.2rem 1.4rem;
      border-radius: 18px;
      backdrop-filter: blur(10px);
      background: linear-gradient(135deg, rgba(17,24,39,0.6), rgba(17,24,39,0.35));
      border: 1px solid rgba(255,255,255,0.08);
      box-shadow: 0 20px 40px rgba(124,58,237,0.20);
      margin-bottom: 12px;
    }}
    .hero h1 {{ font-weight: 800; letter-spacing: 0.5px; color: #ECF2FF; margin: 0 0 0.25rem 0; }}
    .hero p  {{ color: rgba(236,242,255,0.85); margin: 0.25rem 0 0 0; }}
    .badge {{
      display: inline-block; padding: 4px 10px; border-radius: 999px;
      background: rgba(124,58,237,0.16); color: #C4B5FD; font-weight: 600;
      border: 1px solid rgba(124,58,237,0.28); margin-right: 8px;
    }}
    .card {{
      padding: 1rem 1.2rem; border-radius: 16px;
      background: linear-gradient(135deg, rgba(17,24,39,0.6), rgba(17,24,39,0.35));
      border: 1px solid rgba(255,255,255,0.08);
      transition: box-shadow 0.2s ease, transform 0.2s ease;
    }}
    .card:hover {{ transform: translateY(-2px); box-shadow: 0 14px 30px rgba(124,58,237,0.18); }}
    .card h3 {{ margin: 0; color: #E5E7EB; }}
    .card .value {{ font-size: 28px; font-weight: 700; color: #FFFFFF; }}
    .stButton>button, .stDownloadButton>button {{
      border-radius: 12px; padding: 0.6rem 1rem; font-weight: 600;
      border: 1px solid rgba(255,255,255,0.12); color: #ECF2FF;
      background: linear-gradient(135deg, {accent}, #4C1D95);
      box-shadow: 0 8px 16px rgba(124,58,237,0.28);
    }}
    .stButton>button:hover, .stDownloadButton>button:hover {{
      box-shadow: 0 12px 24px rgba(124,58,237,0.40); filter: brightness(1.02);
    }}
    .stTabs [role="tablist"] button[role="tab"] {{
      padding: 10px 14px; border-radius: 12px; margin-right: 6px;
      background: rgba(255,255,255,0.05); color: #E5E7EB;
    }}
    .stTabs [role="tablist"] button[aria-selected="true"] {{
      background: rgba(124,58,237,0.18); color: #FFFFFF; border: 1px solid rgba(124,58,237,0.28);
    }}
    .stDataFrame thead tr th {{ background: rgba(255,255,255,0.06) !important; color: #ECF2FF !important; }}
    </style>
    """, unsafe_allow_html=True)

with st.sidebar:
    st.subheader("🎨 Dizaino akcentas")
    accent = st.color_picker("Pasirinkite akcento spalvą", "#7C3AED")
inject_wow_css(accent)

st.markdown("""
<div class="hero">
  <span class="badge">Akto atnaujinimas</span>
  <span class="badge">Streamlit · WOW UI</span>
  <h1>Automatinis akto užpildymas iš grafiko</h1>
  <p>Metai ir mėnuo → teisingos dienos, šventinės, <strong>Periodiškumas</strong> ir <strong>Kaina</strong> formulėmis.</p>
</div>
""", unsafe_allow_html=True)

# ------------------ PRESET BLOKAI + MĖNESIAI ------------------

PRESET_BLOCKS = [
    ("Lapkritis · C5:G47", "C5:G47", [11]),
    ("Gruodis–Sausis–Vasaris · H5:L47", "H5:L47", [12, 1, 2]),
    ("Kovas · M5:Q47", "M5:Q47", [3]),
    ("Balandis · R5:V47", "R5:V47", [4]),
    ("Gegužė–Birželis–Liepa–Rugpjūtis–Rugsėjis · W5:AA47", "W5:AA47", [5, 6, 7, 8, 9]),
    ("Spalis · AB5:AF47", "AB5:AF47", [10]),
]

MONTH_NAME_LT = {1:'Sausis',2:'Vasaris',3:'Kovas',4:'Balandis',5:'Gegužė',6:'Birželis',7:'Liepa',8:'Rugpjūtis',9:'Rugsėjis',10:'Spalis',11:'Lapkritis',12:'Gruodis'}
MONTH_GEN_LT  = {1:'SAUSIO',2:'VASARIO',3:'KOVO',4:'BALANDŽIO',5:'GEGUŽĖS',6:'BIRŽELIO',7:'LIEPOS',8:'RUGPJŪČIO',9:'RUGSĖJO',10:'SPALIO',11:'LAPKRIČIO',12:'GRUODŽIO'}

# ------------------ PAGALBINĖS FUNKCIJOS ------------------

def _norm(s: str) -> str:
    if s is None: return ''
    s = str(s).replace('\u2013','-').replace('\u2014','-').replace('\u00a0',' ')
    return ' '.join(s.split()).lower()

def normalize_text(x):
    if pd.isna(x): return ''
    if isinstance(x, (int, float)): return str(x).strip()
    return str(x).strip()

def _tokens(s: str) -> list:
    s = _norm(s)
    return [t for t in re.split(r'[^a-z0-9ąčęėįšųūž]+', s) if t]

def save_to_temp(uploaded_file, suffix: str):
    t = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    t.write(uploaded_file.read()); t.flush(); t.close()
    return t.name

def col_letter_to_index(col: str) -> int:
    col = re.sub(r'[^A-Za-z]', '', col).upper()
    val = 0
    for ch in col: val = val * 26 + (ord(ch) - ord('A') + 1)
    return val - 1

def parse_a1_cell(cell: str):
    m = re.match(r'([A-Za-z]+)(\d+)', cell.strip())
    if not m: raise ValueError(f"Neteisingas A1 adresas: {cell}")
    col = col_letter_to_index(m.group(1)); row = int(m.group(2)) - 1
    return row, col

def parse_a1_range(a1: str, shape):
    parts = a1.split(':')
    if len(parts) != 2: raise ValueError("Diapazoną nurodykite A1 formatu, pvz.: C5:G47")
    r0, c0 = parse_a1_cell(parts[0]); r1, c1 = parse_a1_cell(parts[1])
    r0, r1 = sorted([r0, r1]); c0, c1 = sorted([c0, c1])
    max_r, max_c = shape[0]-1, shape[1]-1
    r0 = max(0, min(r0, max_r)); r1 = max(0, min(r1, max_r))
    c0 = max(0, min(c0, max_c)); c1 = max(0, min(c1, max_c))
    return r0, r1, c0, c1

# RAW grafiko nuskaitymas
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

# Header wizard: ffill+bfill per sulietas celes
WEEKDAY_ALIASES = ['pirmadienis','antradienis','trečiadienis','ketvirtadienis','penktadienis','šeštadienis','sekmadienis','i-dienis','ii-dienis','iii-dienis','iv-dienis','v-dienis']
def build_headers(raw_df: pd.DataFrame, start_row: int, header_rows_up: int):
    top = raw_df.iloc[start_row-header_rows_up:start_row+1, :].copy().fillna('')
    top_ff = top.T.ffill().T
    top_full = top_ff.T.bfill().T
    data = raw_df.iloc[start_row+1:, :].reset_index(drop=True)
    headers = []
    for c in range(top_full.shape[1]):
        parts = [normalize_text(top_full.iat[r, c]) for r in range(top_full.shape[0])]
        parts = [p for p in parts if p]
        headers.append(' \n '.join(parts) if parts else f'Col_{c}')
    data.columns = headers
    return data, headers

# Mėnesio aliasai grafiko paieškai
LT_MONTH_ALIASES = {
    1:['SAUSIS','SAUSIO','SAU','SAU.','SAUS'], 2:['VASARIS','VASARIO','VAS','VAS.'],
    3:['KOVAS','KOVO','KOV','KOV.'], 4:['BALANDIS','BALANDŽIO','BAL','BAL.','BALANDZIO'],
    5:['GEGUŽĖ','GEGUŽĖS','GEG','GEG.','GEGUZE','GEGUZES'], 6:['BIRŽELIS','BIRŽELIO','BIR','BIR.','BIRZELIS','BIRZELIO'],
    7:['LIEPA','LIEPOS','LIE','LIE.'], 8:['RUGPJŪTIS','RUGPJŪČIO','RGP','RGP.','RUGPJ','RUGP'],
    9:['RUGSĖJIS','RUGSĖJO','RUGS','RUGS.','RGS','RGS.','RUGSEJIS','RUGSEJO'], 10:['SPALIS','SPALIO','SPA','SPA.'],
    11:['LAPKritis','LAPKRIČIO','LAPK','LAPK.','LAPKRICIO','Lapkritis','Lapkričio'], 12:['GRUODIS','GRUODŽIO','GRU','GRU.','GRUODZIS','GRUODZIO']
}
def month_aliases(month:int):
    return [a.lower() for a in LT_MONTH_ALIASES.get(month, [])]
def find_month_columns_from_headers(headers: list, month: int):
    aliases = month_aliases(month); found = []
    for h in headers:
        toks = _tokens(h)
        if any(a in toks for a in aliases): found.append(h)
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

# Žymų sujungimas iš grafiko stulpelių
FREQUENCY_PATTERNS = [
    (re.compile(r'\b(\d+)\s*kart(?:as|ai)\s*per\s*mėn', re.I), 'times_per_month'),
    (re.compile(r'\b(\d+[.,]?\d*)\s*kv\.m\.?', re.I), 'sqm'),
    (re.compile(r'\b(\d+[.,]?\d*)\s*vnt\.?', re.I), 'units'),
]
def collect_marker_from_columns(row: pd.Series, selected_cols: list):
    texts, numbers, x_flag = [], [], False
    for c in selected_cols:
        val = normalize_text(row.get(c, ''))
        if not val: continue
        if 'x' in _norm(val): x_flag = True
        for rgx, kind in FREQUENCY_PATTERNS:
            m = rgx.search(val); 
            if m: texts.append((val, kind, m.group(1)))
        try:
            n = float(val.replace(',', '.')); numbers.append(n)
        except: pass
    if x_flag: return 'X'
    if texts:
        val, kind, num = texts[0]
        return f'{num} kartas per mėn.' if kind == 'times_per_month' else val
    if numbers: return str(numbers[0])
    return ''

# Akto atnaujinimas (periodiškumą skaičiuosime formulėmis)
def update_act_from_schedule(act_df: pd.DataFrame, schedule_df: pd.DataFrame, month_cols: list):
    name_col = act_df.columns[0]
    period_col = next((c for c in act_df.columns if 'periodi' in _norm(c)), None)
    updated_df = act_df.copy()
    schedule_lookup = {}
    for _, row in schedule_df.iterrows():
        task = normalize_text(row.get(schedule_df.columns[0], ''))
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
        if marker and period_col:
            updated_df.at[idx, period_col] = marker
    return updated_df

# Šventinės ir darbo dienų skaičiavimas
def easter_sunday(year: int) -> date:
    a = year % 19; b = year // 100; c = year % 100
    d = b // 4; e = b % 4; f = (b + 8) // 25; g = (b - f + 1) // 3
    h = (19*a + b - d - g + 15) % 30; i = c // 4; k = c % 4
    l = (32 + 2*e + 2*i - h - k) % 7; m = (a + 11*h + 22*l) // 451
    month = (h + l - 7*m + 114) // 31
    day   = ((h + l - 7*m + 114) % 31) + 1
    return date(year, month, day)

def first_sunday(year: int, month: int) -> date:
    d = date(year, month, 1)
    while d.weekday() != 6: d += timedelta(days=1)
    return d

def lt_fixed_holidays(year: int) -> set:
    """
    Fiksuotos šventinės:
    01-01, 02-16, 03-11, 05-01, 06-24, 07-06, 08-15, 11-01, 11-02, 12-24, 12-25, 12-26
    """
    return {
        date(year, 1, 1),
        date(year, 2, 16),
        date(year, 3, 11),
        date(year, 5, 1),
        date(year, 6, 24),
        date(year, 7, 6),
        date(year, 8, 15),
        date(year, 11, 1),
        date(year, 11, 2),
        date(year, 12, 24),
        date(year, 12, 25),
        date(year, 12, 26),
    }

def lt_variable_holidays(year: int,
    include_easter_sunday=True, 
    include_easter_monday=True,
    include_pentecost_sunday=False,   # <-- Sekminės default: NEįtraukiame
    include_mothers_day=True, 
    include_fathers_day=True) -> set:
    """
    Kintamos šventinės:
    Velykų sekmadienis, Velykų pirmadienis,
    (Sekminės – if include_pentecost_sunday=True),
    Motinos diena (gegužės 1-as sekmadienis), Tėvo diena (birželio 1-as sekmadienis)
    """
    es = easter_sunday(year)       # Velykų sekmadienis
    em = es + timedelta(days=1)    # Velykų pirmadienis (pvz., 2026-04-06)
    ps = es + timedelta(days=49)   # Sekminių sekmadienis
    md = first_sunday(year, 5)     # Motinos diena
    fd = first_sunday(year, 6)     # Tėvo diena
    s = set()
    if include_easter_sunday:    s.add(es)
    if include_easter_monday:    s.add(em)
    if include_pentecost_sunday: s.add(ps)  # default False
    if include_mothers_day:      s.add(md)
    if include_fathers_day:      s.add(fd)
    return s

def holidays_in_month(year: int, month: int, holidays: set) -> list:
    return sorted([h for h in holidays if h.month == month])

# Savaitės dienų aliasai (platus)
WEEKDAY_ALIAS_MAP = {
    0: ['pirmadienis','pirmadien','pir','pirm','i-dienis','i','1','mon','monday'],
    1: ['antradienis','antradien','antr','ant','ii-dienis','ii','2','tue','tuesday'],
    2: ['trečiadienis','treciadienis','trečiadien','treciadien','tre','trec','iii-dienis','iii','3','wed','wednesday'],
    3: ['ketvirtadienis','ketvirtadien','ket','ketv','ketvirtadie','iv-dienis','iv','4','thu','thursday'],
    4: ['penktadienis','penktadien','penk','pen','penktadie','v-dienis','v','5','fri','friday'],
    5: ['šeštadienis','sestadienis','šeštadien','sestadien','ses','vi','6','sat','saturday'],
    6: ['sekmadienis','sekmadien','sek','vii','7','sun','sunday'],
}
SPECIAL_GROUPS = {'savaitgalis': {5, 6}, 'darbo': {0,1,2,3,4}, 'visos': {0,1,2,3,4,5,6}}

def header_to_weekday_set(header: str) -> set:
    toks = set(_tokens(header))
    for key, s in SPECIAL_GROUPS.items():
        if key in toks: return set(s)
    chosen = set()
    for wd, aliases in WEEKDAY_ALIAS_MAP.items():
        if any(a in toks for a in aliases): chosen.add(wd)
    return chosen

def map_headers_to_weekdaysets(headers: list) -> dict:
    mapping = {}
    for h in headers:
        wset = header_to_weekday_set(h)
        if wset: mapping[h] = wset
    return mapping

def count_selected_workdays(year:int, month:int, weekday_set:set, holidays:set) -> int:
    if not weekday_set: return 0
    cal = calendar.Calendar(); cnt = 0
    for d in cal.itermonthdates(year, month):
        if d.month != month: continue
        if d.weekday() in weekday_set and d not in holidays:
            cnt += 1
    return cnt

def list_selected_workdays(year:int, month:int, weekday_set:set, holidays:set) -> list:
    if not weekday_set: return []
    cal = calendar.Calendar(); out = []
    for d in cal.itermonthdates(year, month):
        if d.month != month: continue
        if d.weekday() in weekday_set and d not in holidays:
            out.append(d)
    return out

def render_month_period_genitive(month:int, year:int) -> str:
    last_day = calendar.monthrange(year, month)[1]
    gen_name = MONTH_GEN_LT.get(month, 'MĖNESIO').upper()
    return f"{gen_name} 1-{last_day}"

def excel_col_letter(idx_zero_based: int) -> str:
    letters = ''
    n = idx_zero_based + 1
    while n:
        n, r = divmod(n - 1, 26)
        letters = chr(65 + r) + letters
    return letters

def write_month_end_date(ws, cell_addr: str, year: int, month: int):
    last_day = calendar.monthrange(year, month)[1]
    cell = ws[cell_addr]
    cell.value = date(year, month, last_day)
    cell.number_format = 'yyyy-mm-dd'

# ------------------ UI: TAB'AI ------------------

tab1, tab2, tab3, tab4 = st.tabs(["📤 Įkėlimas", "📅 Grafikas", "🧮 Skaičiavimai", "📄 Aktas"])

with tab1:
    schedule_file = st.file_uploader('Įkelkite grafiko failą (.ods/.xlsx/.xls)', type=['ods','xlsx','xls'])
    act_file = st.file_uploader('Įkelkite akto failą (.xls/.xlsx)', type=['xls','xlsx'])

if not (schedule_file and act_file):
    st.info('Įkelkite grafiką ir aktą, tada pereikite per skiltis.')
    st.stop()

# ------------------ PAGRINDINĖ LOGIKA ------------------

# RAW grafikas
raw_sched_df, sched_sheet = parse_schedule_raw(schedule_file)

with tab2:
    st.caption(f"Grafiko forma: {raw_sched_df.shape[0]} eil. × {raw_sched_df.shape[1]} kol.")
    st.dataframe(raw_sched_df.head(20), use_container_width=True)

    st.markdown("### Pasirinkite mėnesio bloką ir konkretų mėnesį")
    preset_label = st.selectbox("Mėnesio blokas (A1)", options=[p[0] for p in PRESET_BLOCKS], index=1)
    preset_map = {p[0]: (p[1], p[2]) for p in PRESET_BLOCKS}
    a1_default, months_in_block = preset_map[preset_label]

    year_choice = st.number_input("Metai", min_value=2000, max_value=2100, value=dt.now().year)
    now_month = dt.now().month
    default_idx = months_in_block.index(now_month) if now_month in months_in_block else 0
    month_choice = st.selectbox("Konkretaus mėnesio pasirinkimas šiame bloke",
                                 options=months_in_block, index=default_idx,
                                 format_func=lambda m: MONTH_NAME_LT[m])

    month_cell_addr = st.text_input("Įrašyti mėnesį į langelį (A1 adresas)", value="C7")
    a1 = st.text_input("A1 diapazonas (galite pakoreguoti ranka)", value=a1_default)

    r0, r1, c0, c1 = parse_a1_range(a1.strip(), raw_sched_df.shape)
    block = raw_sched_df.iloc[int(r0):int(r1)+1, int(c0):int(c1)+1].reset_index(drop=True)
    st.subheader('Pasirinktas grafiko blokas (peržiūra)')
    st.dataframe(block.head(20), use_container_width=True)

    # Header wizard
    guess_row = 0
    for i in range(min(40, block.shape[0])):
        vals = [normalize_text(v) for v in list(block.iloc[i, :].values)]
        score = sum(1 for v in vals if any(w in _norm(v) for w in WEEKDAY_ALIASES))
        if score >= 3: guess_row = i; break
    start_row = st.number_input('Eilutė bloke, kur yra bazinės antraštės (pvz., Pirmadienis…)',
                                min_value=0, max_value=int(block.shape[0]-2), value=int(guess_row), step=1)
    header_rows_up = st.number_input('Kiek eilučių virš jos – „grupės“ antraštė',
                                     min_value=0, max_value=10, value=1, step=1)
    sched_df, headers = build_headers(block, start_row, header_rows_up)
    st.write("Sukurtos antraštės (pirmi 40):", headers[:40])
    st.dataframe(sched_df.head(12), use_container_width=True)

    # Auto radimas pagal mėnesį
    auto_cols = find_month_columns_from_headers(headers, month_choice)
    if auto_cols:
        st.info(f"Rastos su {MONTH_NAME_LT[month_choice]} susijusios kolonos: {auto_cols[:10]}{'...' if len(auto_cols)>10 else ''}")
    else:
        st.warning("Pagal mėnesio aliasus automatiškai nieko nerasta — pasirinkite rankiniu būdu.")

    substr = st.text_input("Papildomas rankinis filtras antraštėms", value='')
    manual_hits = [h for h in headers if substr.strip() and _norm(substr) in _norm(h)]
    if manual_hits: st.success(f"Pagal filtrą rasta: {manual_hits[:10]}{'...' if len(manual_hits)>10 else ''}")
    preselect = (auto_cols or manual_hits)[:10]
    selected_cols = st.multiselect("Pasirinkite grafiko stulpelius (pvz., Pirmadienis–Penktadienis)",
                                   options=headers, default=preselect)
    if not selected_cols:
        st.error("Nepasirinkote jokių stulpelių. Pasirinkite bent vieną.")
        st.stop()

with tab3:
    st.markdown("### Šventinės dienos ir skaičiavimai")
    fixed_h = lt_fixed_holidays(year_choice)

    include_easter_sunday = st.checkbox("Įtraukti Velykų sekmadienį", value=True)
    include_easter_monday = st.checkbox("Įtraukti Velykų pirmadienį (darbo dienoms turi įtaką)", value=True)
    # Sekminės — BY DEFAULT NEįtraukiame (value=False)
    include_pentecost_sun  = st.checkbox("Įtraukti Sekminių sekmadienį", value=False)
    include_mothers_day    = st.checkbox("Įtraukti Motinos dieną (gegužės 1-as sekmadienis)", value=True)
    include_fathers_day    = st.checkbox("Įtraukti Tėvo dieną (birželio 1-as sekmadienis)", value=True)

    variable_h = lt_variable_holidays(
        year_choice,
        include_easter_sunday, 
        include_easter_monday,
        include_pentecost_sun, 
        include_mothers_day, 
        include_fathers_day
    )

    st.caption("Papildomos šventinės (rankiniu būdu): įveskite YYYY-MM-DD per kablelį.")
    extra_add_text = st.text_input("Pridėti datas", value="")
    extra_del_text = st.text_input("Pašalinti datas", value="")

    def parse_dates_csv(csv_text: str) -> set:
        s = set()
        if csv_text.strip():
            for part in csv_text.split(','):
                tok = part.strip()
                if not tok: continue
                try:
                    y, m, d = map(int, tok.split('-')); s.add(date(y, m, d))
                except:
                    st.warning(f"Praleista neteisinga data: {tok}")
        return s

    extra_add_set = parse_dates_csv(extra_add_text)
    extra_del_set = parse_dates_csv(extra_del_text)

    # Sudarome galutinį rinkinį
    holidays_final = (fixed_h.union(variable_h).union(extra_add_set)).difference(extra_del_set)

    # --- Saugiklis: Sekminės laikome "darbo diena" (pašaliname iš šventinių) ---
    es = easter_sunday(year_choice)
    ps = es + timedelta(days=49)  # Sekminių sekmadienis
    holidays_final.discard(ps)

    holidays_month = holidays_in_month(year_choice, month_choice, holidays_final)

    # Skaičiuojame weekday skaičius (Pir..Sek)
    weekday_counts = {w: count_selected_workdays(year_choice, month_choice, {w}, holidays_final) for w in range(7)}
    weekday_names  = {0:'Pir',1:'An',2:'Tre',3:'Ket',4:'Pen',5:'Šeš',6:'Sek'}

    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.markdown(f'<div class="card"><h3>Mėnuo</h3><div class="value">{MONTH_NAME_LT[month_choice]} {year_choice}</div></div>', unsafe_allow_html=True)
    with col_b:
        wd_union = sum(weekday_counts[w] for w in range(0,5))  # bendros darbo dienos (informacinė)
        st.markdown(f'<div class="card"><h3>Darbo dienos</h3><div class="value">{wd_union}</div></div>', unsafe_allow_html=True)
    with col_c:
        last_day = calendar.monthrange(year_choice, month_choice)[1]
        st.markdown(f'<div class="card"><h3>Intervalas</h3><div class="value">1–{last_day}</div></div>', unsafe_allow_html=True)

    with st.expander("Šventinės, patenkančios į pasirinktą mėnesį"):
        if holidays_month: st.write(holidays_month)
        else: st.info("Šiame mėnesyje šventinių (po korekcijos) nėra arba jos nepatenka į darbo dienas.")

with tab4:
    # Akto duomenys
    act_df, act_sheet = load_act(act_file)
    st.subheader('Akto peržiūra')
    st.dataframe(act_df.head(12), use_container_width=True)

    updated_df = update_act_from_schedule(act_df, sched_df, selected_cols)

    # Įkainio konvertavimas į skaičių (be apvalinimo)
    act_cols = list(updated_df.columns)
    rate_col = next((c for c in act_cols if 'įkain' in _norm(c) or 'ikain' in _norm(c) or 'mato vnt' in _norm(c)), None)
    if rate_col:
        def to_float_preserve(x):
            s = normalize_text(x)
            if not s: return x
            try: return float(s.replace(',', '.'))
            except: return x
        updated_df[rate_col] = updated_df[rate_col].apply(to_float_preserve)

    # Rašymas į Excel + FORMULĖS (užtikrinta)
    with st.spinner("Kuriame formules ir generuojame Excel..."):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            target_sheet_name = act_sheet if act_sheet else 'Akto_lapas'

            # Užtikriname stulpelių buvimą
            act_cols = list(updated_df.columns)
            period_col = next((c for c in act_cols if 'periodi' in _norm(c)), None)
            if period_col is None:
                updated_df.insert(len(act_cols), 'Periodiškumas', '')
                period_col = 'Periodiškumas'; act_cols = list(updated_df.columns)
            price_col = next((c for c in act_cols if 'kaina' in _norm(c)), None)
            if price_col is None:
                updated_df.insert(len(act_cols), 'Kaina', '')
                price_col = 'Kaina'; act_cols = list(updated_df.columns)

            updated_df.to_excel(writer, index=False, sheet_name=target_sheet_name)
            ws = writer.sheets[target_sheet_name]

            # Skaičiavimai lapas: Pir..Sek skaičiai
            book = writer.book
            calc_ws = book.create_sheet('Skaičiavimai')
            calc_ws['A1'] = 'Metai'; calc_ws['B1'] = 'Mėnuo'
            calc_ws['A2'] = year_choice; calc_ws['B2'] = month_choice
            for i, h in enumerate(['Pir','An','Tre','Ket','Pen','Šeš','Sek']):
                calc_ws.cell(row=1, column=3+i, value=h)       # C1..I1
                calc_ws.cell(row=2, column=3+i, value=weekday_counts[i])  # C2..I2
            calc_ws['A4'] = 'Šventinės (mėnesio)'; r = 5
            for h in holidays_month:
                calc_ws.cell(row=r, column=1, value=str(h)); r += 1

            # C7 – "{MĖNESIO KILM.} 1-{last}"
            try:
                ws[month_cell_addr] = render_month_period_genitive(month_choice, year_choice)
            except Exception as e:
                st.warning(f"Nepavyko įrašyti mėnesio į {month_cell_addr}: {e}")

            # A6 – paskutinė mėnesio diena (yyyy-mm-dd)
            try:
                write_month_end_date(ws, 'A6', year_choice, month_choice)
            except Exception as e:
                st.warning(f"Nepavyko įrašyti datos į A6: {e}")

            # B8 – FIKSUOTAS stulpelio pavadinimas
            ws["B8"] = "Plotas kv m./kiekis/val"

            # Formulės: Periodiškumas + Kaina (0.00 formatas)
            act_cols = list(updated_df.columns)
            period_idx = act_cols.index(period_col)
            price_idx  = act_cols.index(price_col)
            qty_col    = next((c for c in act_cols if 'plotas' in _norm(c) or 'kiekis' in _norm(c) or 'kv' in _norm(c) or 'val' in _norm(c)), None)
            qty_idx    = act_cols.index(qty_col) if qty_col else None
            rate_idx   = act_cols.index(rate_col) if rate_col else None

            # Akto antraščių atpažinimas kaip savaitės dienų stulpeliai
            act_weekday_map = map_headers_to_weekdaysets(act_cols)

            def weekday_count_cell(w: int) -> str:
                # C2..I2 (Pir..Sek)
                col_letter = chr(67 + w)  # 0->C, 1->D, ... 6->I
                return f"Skaičiavimai!{col_letter}2"

            header_count_expr = {}
            for h, wset in act_weekday_map.items():
                parts = [weekday_count_cell(w) for w in sorted(wset)]
                header_count_expr[h] = "+".join(parts) if parts else "0"

            n_rows = len(updated_df)
            for i in range(2, n_rows + 2):
                # Periodiškumas = Σ IF(header_cell == "X", Skaičiavimai!<weekday sum>, 0)
                parts = []
                for h, expr in header_count_expr.items():
                    col_idx = act_cols.index(h)
                    cell_ref = f"{excel_col_letter(col_idx)}{i}"
                    parts.append(f"IF({cell_ref}=\"X\",{expr},0)")
                formula_period = "=" + ("+".join(parts) if parts else "0")
                ws[f"{excel_col_letter(period_idx)}{i}"] = formula_period

                # Kaina = Plotas * Įkainis * Periodiškumas (be ROUND; tik formatas 0.00)
                if qty_idx is not None and rate_idx is not None:
                    plot_cell   = f"{excel_col_letter(qty_idx)}{i}"
                    rate_cell   = f"{excel_col_letter(rate_idx)}{i}"
                    period_cell = f"{excel_col_letter(period_idx)}{i}"
                    price_cell  = f"{excel_col_letter(price_idx)}{i}"
                    ws[price_cell] = f"=IFERROR({plot_cell}*{rate_cell}*{period_cell},0)"
                    ws[price_cell].number_format = '0.00'

            # Iš viso (be PVM): suma kainų
            total_row = n_rows + 3
            ws[f"A{total_row}"] = "Iš viso (be PVM):"
            sum_cell = f"{excel_col_letter(price_idx)}{total_row}"
            ws[sum_cell] = f"=SUM({excel_col_letter(price_idx)}2:{excel_col_letter(price_idx)}{n_rows+1})"
            ws[sum_cell].number_format = '0.00'

        output.seek(0)

    st.toast("✅ Formulės įrašytos. Galite atsisiųsti.", icon="🎉")
    out_name = f"Aktas_atnaujintas_{year_choice}_{month_choice:02d}.xlsx"
    st.download_button(
        label='Atsisiųsti atnaujintą aktą (.xlsx)',
        data=output,
        file_name=out_name,
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

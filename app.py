
# app.py — Akto atnaujinimas iš grafiko su EXCEL FORMULĖMIS (LT locale, ";" skyriklis)
# - 1–7 eil. paliekamos; antraštės 8-oje; duomenys nuo 9-os
# - A6: paskutinė mėnesio diena (formulė), C7: mėnesio tekstas (formulė)
# - A8 = "Savaitės dienos", C8..G8 = "Pirmadienis".."Penktadienis", B7 tuščia
# - Grafiko žymėjimai perkelti į AKTO lapo bloką C10:G50
# - Periodiškumas/Kaina/SUMA — tik Excel formulės
# - K51 — bendra suma; J51 — "Suma:"
# - PVM netaikomas

import streamlit as st
import pandas as pd
import re
import io
import tempfile
import calendar
from datetime import datetime as dt, date, timedelta
from openpyxl.utils import get_column_letter, column_index_from_string

st.set_page_config(page_title="Akto atnaujinimas iš grafiko", layout="wide")

# ===== KONFIGURACIJA / KONSTANTOS =====
HEADER_TOP_ROW = 7               # 7 eilė – viršus
HEADER_BOTTOM_ROW = 8            # 8 eilė – antraštės
DATA_START_ROW = HEADER_BOTTOM_ROW + 1  # 9 eilė – duomenys
MARK_BLOCK_TOP_LEFT = ("C", 10)  # žymėjimų blokas (C10:G50)
MARK_BLOCK_WIDTH = 5             # C..G (Pir..Pen)
MARK_BLOCK_HEIGHT = 41           # 10..50 (41 eilučių vietos)
FIXED_SUM_CELL = "K51"           # bendra suma
SUM_LABEL_CELL = "J51"           # "Suma:"

# ===== PRESET BLOKAI grafiko A1 diapazonams (peržiūra) =====
PRESET_BLOCKS = [
    ("Lapkritis · C5:G47", "C5:G47", [11]),
    ("Gruodis–Sausis–Vasaris · H5:L47", "H5:L47", [12, 1, 2]),
    ("Kovas · M5:Q47", "M5:Q47", [3]),
    ("Balandis · R5:V47", "R5:V47", [4]),
    ("Gegužė–Birželis–Liepa–Rugpjūtis–Rugsėjis · W5:AA47", "W5:AA47", [5, 6, 7, 8, 9]),
    ("Spalis · AB5:AF47", "AB5:AF47", [10]),
]
MONTH_NAME_LT = {1:'Sausis',2:'Vasaris',3:'Kovas',4:'Balandis',5:'Gegužė',6:'Birželis',
                 7:'Liepa',8:'Rugpjūtis',9:'Rugsėjis',10:'Spalis',11:'Lapkritis',12:'Gruodis'}

# ===== Pagalbinės =====
WEEKDAY_ALIASES = [
    'pirmadienis','antradienis','trečiadienis','ketvirtadienis','penktadienis',
    'šeštadienis','sekmadienis','i-dienis','ii-dienis','iii-dienis','iv-dienis','v-dienis'
]

def _norm(s: str) -> str:
    if s is None: return ''
    s = str(s).replace('\u2013','-').replace('\u2014','-').replace('\u00a0',' ')
    return ' '.join(s.split()).lower()

def normalize_text(x):
    if pd.isna(x): return ''
    if isinstance(x, (int, float)): return str(x).strip()
    return str(x).strip()

def save_to_temp(uploaded_file, suffix: str):
    t = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    t.write(uploaded_file.read()); t.flush(); t.close()
    return t.name

def col_letter_to_index(col: str) -> int:
    col = re.sub(r'[^A-Za-z]', '', col).upper()
    val = 0
    for ch in col: val = val*26 + (ord(ch)-ord('A')+1)
    return val-1

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

def _tokens(s: str) -> list:
    s = _norm(s)
    return [t for t in re.split(r'[^a-z0-9ąčęėįšųūž]+', s) if t]

def month_aliases(month:int):
    LT_MONTH_ALIASES = {
        1:['SAUSIS','SAUSIO','SAU','SAU.','SAUS'], 2:['VASARIS','VASARIO','VAS','VAS.'],
        3:['KOVAS','KOVO','KOV','KOV.'], 4:['BALANDIS','BALANDŽIO','BAL','BAL.','BALANDZIO'],
        5:['GEGUŽĖ','GEGUŽĖS','GEG','GEG.','GEGUZE','GEGUZES'], 6:['BIRŽELIS','BIRŽELIO','BIR','BIR.','BIRZELIS','BIRZELIO'],
        7:['LIEPA','LIEPOS','LIE','LIE.'], 8:['RUGPJŪTIS','RUGPJŪČIO','RGP','RGP.','RUGPJ','RUGP'],
        9:['RUGSĖJIS','RUGSĖJO','RUGS','RUGS.','RGS','RGS.','RUGSEJIS','RUGSEJO'],
        10:['SPALIS','SPALIO','SPA','SPA.'], 11:['LAPKritis','LAPKRIČIO','LAPK','LAPK.','LAPKRICIO','Lapkritis','Lapkričio'],
        12:['GRUODIS','GRUODŽIO','GRU','GRU.','GRUODZIS','GRUODZIO'],
    }
    return [a.lower() for a in LT_MONTH_ALIASES.get(month, [])]

def find_month_columns_from_headers(headers: list, month: int):
    aliases = month_aliases(month)
    found = []
    for h in headers:
        toks = _tokens(h)
        if any(a in toks for a in aliases): found.append(h)
    return found

# ===== Grafiko nuskaitymas peržiūrai =====
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

def build_headers(raw_df: pd.DataFrame, start_row: int, header_rows_up: int):
    top = raw_df.iloc[start_row-header_rows_up:start_row+1, :].copy().fillna('')
    top_ff = top.T.ffill().T
    top_full = top_ff.T.bfill().T
    data = raw_df.iloc[start_row+1:, :].reset_index(drop=True)
    headers = []
    for c in range(top_full.shape[1]):
        parts = [normalize_text(top_full.iat[r, c]) for r in range(top_full.shape[0])]
        parts = [p for p in parts if p]
        headers.append(' \n'.join(parts) if parts else f'Col_{c}')
    data.columns = headers
    return data, headers

# ===== Akto nuskaitymas su 7–8 antraštėmis =====
def build_headers_from_two_rows(df_raw: pd.DataFrame, top_row_1based: int, bottom_row_1based: int):
    r_top = top_row_1based - 1
    r_bot = bottom_row_1based - 1
    top = df_raw.iloc[[r_top, r_bot], :].copy().fillna("")
    tb = top.T
    tb_ff = tb.ffill()
    tb_full = tb_ff.bfill()
    top_full = tb_full.T
    headers = []
    for c in range(top_full.shape[1]):
        parts = [str(top_full.iat[r, c]).strip() for r in range(top_full.shape[0])]
        parts_clean = [p for p in parts if p]
        col_name = " / ".join(parts_clean) if parts_clean else f"Col_{c}"
        headers.append(col_name)
    df_data = df_raw.iloc[bottom_row_1based:, :].reset_index(drop=True)
    df_data.columns = headers
    return df_data, headers

def load_act_both(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith('.xlsx'):
        path = save_to_temp(uploaded_file, '.xlsx')
        xls = pd.ExcelFile(path, engine='openpyxl')
        sheet = st.selectbox("Pasirinkite akto lapą", options=xls.sheet_names)
        df_raw = xls.parse(sheet, header=None)
        df_table, _ = build_headers_from_two_rows(df_raw, HEADER_TOP_ROW, HEADER_BOTTOM_ROW)
        return df_raw, df_table, sheet
    elif name.endswith('.xls'):
        path = save_to_temp(uploaded_file, '.xls')
        xls = pd.ExcelFile(path, engine='xlrd')
        sheet = st.selectbox("Pasirinkite akto lapą", options=xls.sheet_names)
        df_raw = xls.parse(sheet, header=None)
        df_table, _ = build_headers_from_two_rows(df_raw, HEADER_TOP_ROW, HEADER_BOTTOM_ROW)
        return df_raw, df_table, sheet
    else:
        raise ValueError('Nepalaikomas akto formatas. Naudokite .xls arba .xlsx')

def find_col_exact_or_prefix(headers_list, target):
    for i, h in enumerate(headers_list):
        if str(h).strip() == target:
            return i + 1
    for i, h in enumerate(headers_list):
        if str(h).strip().startswith(target):
            return i + 1
    return -1

# ===== UI =====
st.title('Akto atnaujinimas iš grafiko (diapazonai + formulės)')
st.caption('Grafikas: .ods/.xlsx/.xls, Aktas: .xls/.xlsx. PVM – netaikomas.')

schedule_file = st.file_uploader('Įkelkite grafiko failą (.ods/.xlsx/.xls)', type=['ods','xlsx','xls'])
act_file = st.file_uploader('Įkelkite akto failą (.xls/.xlsx)', type=['xls','xlsx'])

if schedule_file and act_file:
    try:
        # --- Grafikas peržiūrai ---
        raw_sched_df, sched_sheet = parse_schedule_raw(schedule_file)
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

        a1 = st.text_input("A1 diapazonas", value=a1_default)
        r0, r1, c0, c1 = parse_a1_range(a1.strip(), raw_sched_df.shape)
        block = raw_sched_df.iloc[int(r0):int(r1)+1, int(c0):int(c1)+1].reset_index(drop=True)
        st.subheader('Pasirinktas grafiko blokas (peržiūra)')
        st.dataframe(block.head(20), use_container_width=True)

        # Header wizard bloke
        guess_row = 0
        for i in range(min(40, block.shape[0])):
            vals = [normalize_text(v) for v in list(block.iloc[i, :].values)]
            score = sum(1 for v in vals if any(w in _norm(v) for w in WEEKDAY_ALIASES))
            if score >= 3: guess_row = i; break
        start_row = st.number_input('Eilutė bloke, kur yra bazinės antraštės (pvz., Pirmadienis…)',
                                    min_value=0, max_value=int(block.shape[0]-2),
                                    value=int(guess_row), step=1)
        header_rows_up = st.number_input('Kiek eilučių virš jos – „grupės“ antraštė',
                                         min_value=0, max_value=10, value=1, step=1)
        sched_df, headers = build_headers(block, start_row, header_rows_up)
        st.write("Sukurtos antraštės (pirmi 40):", headers[:40])
        st.dataframe(sched_df.head(12), use_container_width=True)

        # automatinė paieška pagal mėnesį + rankinis filtras
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
        preselect = (auto_cols or manual_hits)[:5]  # mums reikia Pir..Pen (5 stulp.)
        selected_cols = st.multiselect("Pasirinkite 5 stulpelius (Pir..Pen)", options=headers, default=preselect)
        if not selected_cols:
            st.error("Nepasirinkote jokių stulpelių. Pasirinkite bent vieną.")
            st.stop()

        # --- Aktas: nuskaitymas su 7–8 eil. antraštėmis; 1–7 viršus paliekamas ---
        df_raw_act, act_df, act_sheet = load_act_both(act_file)
        st.subheader('Akto peržiūra (nuo 9-os eil.)')
        st.dataframe(act_df.head(10), use_container_width=True)

        headers_act = list(act_df.columns)
        # privalomi pavadinimai
        col_plotas  = find_col_exact_or_prefix(headers_act, "Plotas kv m./kiekis/val")
        col_ikainis = find_col_exact_or_prefix(headers_act, "įkainis")
        col_period  = find_col_exact_or_prefix(headers_act, "Periodiškumas")
        col_kaina   = find_col_exact_or_prefix(headers_act, "Kaina")

        missing = []
        for name, idx in [("Plotas kv m./kiekis/val", col_plotas),
                          ("įkainis", col_ikainis),
                          ("Periodiškumas", col_period),
                          ("Kaina", col_kaina)]:
            if idx == -1: missing.append(name)
        if missing:
            st.error("Trūksta stulpelių (po 7–8 eil. suformavimo): " + ", ".join(missing))
            st.stop()

        # --- EKSPORTAS SU FORMULĖMIS ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Aktas — header 8-oje, duomenys nuo 9-os
            act_df.to_excel(writer, index=False, sheet_name=act_sheet, startrow=HEADER_BOTTOM_ROW-1)
            wb = writer.book
            ws_act = wb[act_sheet]

            # Grąžiname VIRŠŲ (1–7 eil.) iš originalo
            n_cols_top = df_raw_act.shape[1]
            for r in range(1, 8):  # 1..7 eil.
                for c in range(1, n_cols_top + 1):
                    val = df_raw_act.iat[r-1, c-1] if (c-1) < df_raw_act.shape[1] else None
                    if pd.isna(val): val = None
                    ws_act.cell(row=r, column=c, value=val)

            # Skaičiavimai: Pir..Pen (365 formulės), remiasi į A2/B2; Šventinės E2:E200
            calc_ws = wb.create_sheet("Skaičiavimai")
            calc_ws["A1"] = "Metai";  calc_ws["B1"] = "Mėnuo"
            calc_ws["A2"] = int(year_choice); calc_ws["B2"] = int(month_choice)
            calc_ws["E1"] = "Šventinės (E2:E200)"

            # Automatiškai įrašome Velykų pirmadienį (jei patenka į pasirinktą mėn.)
            def easter_sunday(year: int) -> date:
                a = year % 19; b = year // 100; c = year % 100
                d = b // 4; e = b % 4; f = (b + 8) // 25; g = (b - f + 1) // 3
                h = (19*a + b - d - g + 15) % 30; i = c // 4; k = c % 4
                l = (32 + 2*e + 2*i - h - k) % 7; m = (a + 11*h + 22*l) // 451
                month = (h + l - 7*m + 114) // 31
                day = ((h + l - 7*m + 114) % 31) + 1
                return date(year, month, day)
            es = easter_sunday(int(year_choice))
            em = es + timedelta(days=1)  # Velykų pirmadienis
            if em.year == int(year_choice) and em.month == int(month_choice):
                calc_ws["E2"] = em.strftime("%Y-%m-%d")  # įvestis kaip tekstas YYYY-MM-DD

            # Pir..Pen (C1..G1); Formulės C2..G2 su ";" (LT locale), atimant šventines
            headers_w = ["Pir","An","Tre","Ket","Pen"]
            letters   = ["C","D","E","F","G"]
            for i, h in enumerate(headers_w):
                calc_ws.cell(row=1, column=3+i, value=h)  # C1..G1
                # wd_num = i+1 (1=Pir..5=Pen)
                calc_ws[f"{letters[i]}2"] = (
                    f'=LET('
                    f'y;$A$2; m;$B$2; '
                    f'days;SEQUENCE(DAY(EOMONTH(DATE(y;m;1);0))); '
                    f'd;DATE(y;m;days); '
                    f'wd;WEEKDAY(d;2); '
                    f'hol;FILTER($E$2:$E$200;(MONTH($E$2:$E$200)=m)*(YEAR($E$2:$E$200)=y);""); '
                    f'COUNTIF(wd;{i+1})-IFERROR(COUNTIF(WEEKDAY(hol;2);{i+1});0) '
                    f')'
                )
            try:
                wb.calculation.fullCalcOnLoad = True
            except Exception:
                pass

            # Viršaus formulės (A6/C7)
            ws_act["A6"] = "=EOMONTH(DATE(Skaičiavimai!$A$2;Skaičiavimai!$B$2;1);0)"
            ws_act["A6"].number_format = "yyyy-mm-dd"
            ws_act["C7"] = (
                '=CHOOSE(Skaičiavimai!$B$2;'
                '"SAUSIO";"VASARIO";"KOVO";"BALANDŽIO";"GEGUŽĖS";"BIRŽELIO";'
                '"LIEPOS";"RUGPJŪČIO";"RUGSĖJO";"SPALIO";"LAPKRIČIO";"GRUODŽIO")'
                '&" "&1&"-"&DAY(EOMONTH(DATE(Skaičiavimai!$A$2;Skaičiavimai!$B$2;1);0))'
            )

            # Viršaus/antraščių pataisymai
            ws_act["B7"] = None                 # B7 tuščias
            ws_act["A8"] = "Savaitės dienos"    # A8 — fiksuotas tekstas
            ws_act["C8"] = "Pirmadienis"
            ws_act["D8"] = "Antradienis"
            ws_act["E8"] = "Trečiadienis"
            ws_act["F8"] = "Ketvirtadienis"
            ws_act["G8"] = "Penktadienis"

            # ===== GRAFIKO ŽYMĖJIMŲ PERKĖLIMAS Į AKTĄ (C10:G50) =====
            mark_cols = selected_cols[:MARK_BLOCK_WIDTH] if selected_cols else []
            top_col_letter, top_row = MARK_BLOCK_TOP_LEFT
            top_col_index = column_index_from_string(top_col_letter)  # C -> 3
            max_rows_to_fill = min(MARK_BLOCK_HEIGHT, len(sched_df))
            for ci, col_name in enumerate(mark_cols):
                dest_col = top_col_index + ci   # C..G
                for r_off in range(max_rows_to_fill):
                    src_val = normalize_text(sched_df.iloc[r_off].get(col_name, ""))
                    val = "X" if src_val and ("x" in src_val.lower()) else ""  # X/tuščia
                    ws_act.cell(row=top_row + r_off, column=dest_col, value=val)

            # ===== PERIODIŠKUMO IR KAINOS FORMULĖS (nuo 9-os) =====
            headers_act = list(act_df.columns)
            def find_col_exact_or_prefix(headers_list, target):
                for i, h in enumerate(headers_list):
                    if str(h).strip() == target: return i + 1
                for i, h in enumerate(headers_list):
                    if str(h).strip().startswith(target): return i + 1
                return -1
            col_plotas  = find_col_exact_or_prefix(headers_act, "Plotas kv m./kiekis/val")
            col_ikainis = find_col_exact_or_prefix(headers_act, "įkainis")
            col_period  = find_col_exact_or_prefix(headers_act, "Periodiškumas")
            col_kaina   = find_col_exact_or_prefix(headers_act, "Kaina")

            def addr(col_idx_1based: int, row_1based: int) -> str:
                return f"{get_column_letter(col_idx_1based)}{row_1based}"

            n_data   = len(act_df)
            first_row = DATA_START_ROW
            last_row  = DATA_START_ROW + n_data - 1

            # Periodiškumas = Σ IF(Cr="X";Skaičiavimai!C2;0) + ... + IF(Gr="X";Skaičiavimai!G2;0)
            letters_map = ["C","D","E","F","G"]  # Pir..Pen
            for r in range(first_row, last_row + 1):
                plot_cell   = addr(col_plotas,  r)
                rate_cell   = addr(col_ikainis, r)
                period_cell = addr(col_period,  r)
                price_cell  = addr(col_kaina,   r)

                mark_row = r + 1  # 9 -> 10, 10 -> 11, ... (blokas prasideda 10)
                parts = []
                for i, mcol in enumerate(letters_map):  # i=0..4
                    mark_cell = f"{mcol}{mark_row}"
                    base_cell = f"Skaičiavimai!{letters_map[i]}2"
                    parts.append(f'IF({mark_cell}="X";{base_cell};0)')
                ws_act[period_cell].value = "=" + ("+".join(parts) if parts else "0")

                # Kaina = Plotas * Įkainis * Periodiškumas (be ROUND; rodymas 0.00)
                ws_act[price_cell].value = (
                    f'=IFERROR(VALUE({plot_cell});{plot_cell})'
                    f'*IFERROR(VALUE({rate_cell});{rate_cell})'
                    f'*IFERROR(VALUE({period_cell});{period_cell})'
                )
                ws_act[price_cell].number_format = "0.00"

            # ===== GALUTINĖ SUMA — K51 ir žyma J51 =====
            sum_formula = f"=SUM({addr(col_kaina, first_row)}:{addr(col_kaina, last_row)})"
            ws_act[FIXED_SUM_CELL] = sum_formula
            ws_act[FIXED_SUM_CELL].number_format = "0.00"
            ws_act[SUM_LABEL_CELL] = "Suma:"

        output.seek(0)
        out_name = f"Aktas_atnaujintas_{year_choice}_{month_choice:02d}.xlsx"
        st.download_button(
            label='Atsisiųsti atnaujintą aktą (.xlsx)',
            data=output,
            file_name=out_name,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        st.success('✅ Paruošta: Periodiškumas/kaina — tik formulės, Velykos atimtos, žymėjimai AKTE (C10:G50), bendra suma K51.')
    except Exception as e:
        st.error(f'Klaida: {e}')
        st.exception(e)
else:
    st.info('Įkelkite grafiką ir aktą, tada pasirinkite bloką, konkretų mėnesį ir metus.')


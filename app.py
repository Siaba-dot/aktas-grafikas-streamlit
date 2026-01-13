
# app.py — Akto atnaujinimas iš grafiko su EXCEL FORMULĖMIS
# - 1–7 eil. paliekamos; antraštės 8-oje; duomenys nuo 9-os
# - A8 = "Savaitės dienos"; C8 = "Pirmadienis"; B7 tuščias
# - A6/C7 formulės pagal Skaičiavimai!A2/B2
# - Eilučių Periodiškumas/Kaina = tik Excel formulės
# - Galutinė suma tik K51 (K stulpelio gale)
# - Grafiko įkėlimas (.ods/.xlsx/.xls), A1 diapazonas, header wizard, filtrai palikti
# - PVM netaikomas

import streamlit as st
import pandas as pd
import re
import io
import tempfile
import calendar
from datetime import datetime as dt, date, timedelta
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Akto atnaujinimas iš grafiko", layout="wide")

# ========================= KONSTANTOS =========================
HEADER_TOP_ROW = 7              # 7 eilė – viršutinė (paliekama)
HEADER_BOTTOM_ROW = 8           # 8 eilė – antraštės
DATA_START_ROW = HEADER_BOTTOM_ROW + 1  # 9 eilė – duomenys
FIXED_SUM_CELL = "K51"          # galutinė bendra suma – fiksuota vieta

# ========================= PRESET BLOKAI ======================
PRESET_BLOCKS = [
    ("Lapkritis · C5:G47", "C5:G47", [11]),
    ("Gruodis–Sausis–Vasaris · H5:L47", "H5:L47", [12, 1, 2]),
    ("Kovas · M5:Q47", "M5:Q47", [3]),
    ("Balandis · R5:V47", "R5:V47", [4]),
    ("Gegužė–Birželis–Liepa–Rugpjūtis–Rugsėjis · W5:AA47", "W5:AA47", [5, 6, 7, 8, 9]),
    ("Spalis · AB5:AF47", "AB5:AF47", [10]),
]
MONTH_NAME_LT = {
    1:'Sausis', 2:'Vasaris', 3:'Kovas', 4:'Balandis', 5:'Gegužė', 6:'Birželis',
    7:'Liepa', 8:'Rugpjūtis', 9:'Rugsėjis', 10:'Spalis', 11:'Lapkritis', 12:'Gruodis'
}

# ========================= Pagalbinės =========================
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

# ========================= Grafiko nuskaitymas =========================
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
    """Header wizard (sulietos antraštės) — ffill+bfill."""
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

# ========================= Akto nuskaitymas =========================
def build_headers_from_two_rows(df_raw: pd.DataFrame, top_row_1based: int, bottom_row_1based: int):
    """Sujungia 7 ir 8 eilutes į vieną antraščių rinkinį, taikant ffill + bfill per stulpelį."""
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
    """Grąžina: (df_raw_be_header, df_table_su_header_is_7-8, act_sheet_name)"""
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
    """Ieško tikslios antraštės arba prefikso (kai sujungta „X / Y“). Grąžina 1-based indeksą arba -1."""
    for i, h in enumerate(headers_list):
        if str(h).strip() == target:
            return i + 1
    for i, h in enumerate(headers_list):
        if str(h).strip().startswith(target):
            return i + 1
    return -1

# ========================= UI =========================
st.title('Akto atnaujinimas iš grafiko (diapazonai + formulės)')
st.caption('Grafikas: .ods / .xlsx / .xls. Aktas: .xls / .xlsx. PVM – netaikomas.')

schedule_file = st.file_uploader('Įkelkite grafiko failą (.ods/.xlsx/.xls)', type=['ods','xlsx','xls'])
act_file = st.file_uploader('Įkelkite akto failą (.xls/.xlsx)', type=['xls','xlsx'])

if schedule_file and act_file:
    try:
        # ==== GRAFIKAS ====
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

        a1 = st.text_input("A1 diapazonas (galite pakoreguoti ranka)", value=a1_default)
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
        preselect = (auto_cols or manual_hits)[:10]
        selected_cols = st.multiselect("Pasirinkite stulpelius (pvz., Pirmadienis–Penktadienis)",
                                       options=headers, default=preselect)
        if not selected_cols:
            st.error("Nepasirinkote jokių stulpelių. Pasirinkite bent vieną.")
            st.stop()

        # ==== AKTAS ====
        df_raw_act, act_df, act_sheet = load_act_both(act_file)
        st.subheader('Akto peržiūra (nuo 9-os eil.)')
        st.dataframe(act_df.head(10), use_container_width=True)

        headers_act = list(act_df.columns)
        # privalomi pavadinimai
        col_plotas  = find_col_exact_or_prefix(headers_act, "Plotas kv m./kiekis/val")
        col_ikainis = find_col_exact_or_prefix(headers_act, "įkainis")
        col_period  = find_col_exact_or_prefix(headers_act, "Periodiškumas")
        col_kaina   = find_col_exact_or_prefix(headers_act, "Kaina")

        weekday_targets = ["Pirmadienis","Antradienis","Trečiadienis",
                           "Ketvirtadienis","Penktadienis","Šeštadienis","Sekmadienis"]
        weekday_cols_map = {t: find_col_exact_or_prefix(headers_act, t) for t in weekday_targets}
        has_weekday_x = any(idx != -1 for idx in weekday_cols_map.values())

        missing = []
        for name, idx in [("Plotas kv m./kiekis/val", col_plotas),
                          ("įkainis", col_ikainis),
                          ("Periodiškumas", col_period),
                          ("Kaina", col_kaina)]:
            if idx == -1: missing.append(name)
        if missing:
            st.error("Trūksta stulpelių (po 7–8 eil. suformavimo): " + ", ".join(missing))
            st.stop()

        # ==== EKSPORTAS SU FORMULĖMIS ====
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # 1) Grafikas (peržiūrai)
            sched_df.to_excel(writer, index=False, sheet_name='Grafikas')

            # 2) Aktas — header 8-oje, duomenys nuo 9-os
            act_df.to_excel(writer, index=False, sheet_name=act_sheet, startrow=HEADER_BOTTOM_ROW-1)
            wb = writer.book
            ws_act = wb[act_sheet]

            # 2a) Grąžiname VIRŠŲ (1–7 eil.) iš originalo
            n_cols_top = df_raw_act.shape[1]
            for r in range(1, 8):  # 1..7 eil.
                for c in range(1, n_cols_top + 1):
                    val = df_raw_act.iat[r-1, c-1] if (c-1) < df_raw_act.shape[1] else None
                    if pd.isna(val): val = None
                    ws_act.cell(row=r, column=c, value=val)

            # 3) Skaičiavimai: Pir..Sek (365 formulės), remiasi į A2/B2
            calc_ws = wb.create_sheet("Skaičiavimai")
            calc_ws["A1"] = "Metai";  calc_ws["B1"] = "Mėnuo"
            calc_ws["A2"] = int(year_choice); calc_ws["B2"] = int(month_choice)
            calc_ws["E1"] = "Šventinės (E2:E200)"
            calc_ws["D2"] = "Įveskite YYYY-MM-DD; bus atimta iš mėnesio dienų."
            headers_w = ["Pir", "An", "Tre", "Ket", "Pen", "Šeš", "Sek"]
            letters   = ["C","D","E","F","G","H","I"]
            for i, h in enumerate(headers_w):
                calc_ws.cell(row=1, column=3+i, value=h)  # C1..I1
                calc_ws[f"{letters[i]}2"] = (
                    f'=LET('
                    f'y,$A$2, m,$B$2, '
                    f'days,SEQUENCE(DAY(EOMONTH(DATE(y,m,1),0))), '
                    f'dates,DATE(y,m,days), '
                    f'week,WEEKDAY(dates,2), '
                    f'hol,FILTER($E$2:$E$200,(MONTH($E$2:$E$200)=m)*(YEAR($E$2:$E$200)=y),""), '
                    f'COUNTIF(week,{i+1})-COUNTIF(WEEKDAY(hol,2),{i+1}) '
                    f')'
                )
            try:
                wb.calculation.fullCalcOnLoad = True
            except Exception:
                pass

            # --- Viršaus formulės (A6/C7) ---
            ws_act["A6"] = "=EOMONTH(DATE(Skaičiavimai!$A$2,Skaičiavimai!$B$2,1),0)"
            ws_act["A6"].number_format = "yyyy-mm-dd"
            ws_act["C7"] = (
                '=CHOOSE(Skaičiavimai!$B$2,'
                '"SAUSIO","VASARIO","KOVO","BALANDŽIO","GEGUŽĖS","BIRŽELIO",'
                '"LIEPOS","RUGPJŪČIO","RUGSĖJO","SPALIO","LAPKRIČIO","GRUODŽIO")'
                '&" "&1&"-"&DAY(EOMONTH(DATE(Skaičiavimai!$A$2,Skaičiavimai!$B$2,1),0))'
            )

            # --- Tikslūs viršaus/antraščių pataisymai ---
            ws_act["B7"] = None                 # B7 tuščias
            ws_act["A8"] = "Savaitės dienos"    # A8 privalo likti šis tekstas
            ws_act["C8"] = "Pirmadienis"       # C8 – Pirmadienis (ne mėnesio tekstas)

            # 4) Formulės eilutėms: 9..(9 + len(act_df) - 1)
            n_data   = len(act_df)
            first_row = DATA_START_ROW
            last_row  = DATA_START_ROW + n_data - 1

            def addr(col_idx_1based: int, row_1based: int) -> str:
                return f"{get_column_letter(col_idx_1based)}{row_1based}"

            # Eilučių FORMULĖS (Periodiškumas + Kaina)
            for r in range(first_row, last_row + 1):
                plot_cell   = addr(col_plotas,  r)
                rate_cell   = addr(col_ikainis, r)
                period_cell = addr(col_period,  r)
                price_cell  = addr(col_kaina,   r)

                # PERIODIŠKUMAS:
                if has_weekday_x:
                    parts = []
                    letters_map = ["C","D","E","F","G","H","I"]  # Skaičiavimai!C2..I2
                    wd_to_col = {
                        1: weekday_cols_map["Pirmadienis"],
                        2: weekday_cols_map["Antradienis"],
                        3: weekday_cols_map["Trečiadienis"],
                        4: weekday_cols_map["Ketvirtadienis"],
                        5: weekday_cols_map["Penktadienis"],
                        6: weekday_cols_map["Šeštadienis"],
                        7: weekday_cols_map["Sekmadienis"],
                    }
                    for wd_num, wd_idx in wd_to_col.items():
                        if wd_idx != -1:
                            wd_cell = addr(wd_idx, r)
                            base_cell = f"Skaičiavimai!{letters_map[wd_num-1]}2"
                            parts.append(f'IF({wd_cell}="X",{base_cell},0)')
                    ws_act[period_cell].value = "=" + ("+".join(parts) if parts else "0")
                else:
                    # fallback: X -> 1; skaičius tekste -> VALUE; kita -> 0
                    ws_act[period_cell].value = f'=IF(LOWER({period_cell})="x",1,IFERROR(VALUE({period_cell}),0))'

                # KAINA = Plotas * Įkainis * Periodiškumas (be ROUND; rodymas 0.00)
                ws_act[price_cell].value = (
                    f'=IFERROR(VALUE({plot_cell}),{plot_cell})'
                    f'*IFERROR(VALUE({rate_cell}),{rate_cell})'
                    f'*IFERROR(VALUE({period_cell}),{period_cell})'
                )
                ws_act[price_cell].number_format = "0.00"

            # 5) GALUTINĖ SUMA — tik K51
            sum_formula = f"=SUM({addr(col_kaina, first_row)}:{addr(col_kaina, last_row)})"
            ws_act[FIXED_SUM_CELL] = sum_formula
            ws_act[FIXED_SUM_CELL].number_format = "0.00"

        output.seek(0)
        out_name = f"Aktas_atnaujintas_{year_choice}_{month_choice:02d}.xlsx"
        st.download_button(
            label='Atsisiųsti atnaujintą aktą (.xlsx)',
            data=output,
            file_name=out_name,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        st.success('✅ Paruošta: A8="Savaitės dienos", C8="Pirmadienis", viršus išsaugotas, formulės nuo 9-os, bendra suma K51.')
    except Exception as e:
        st.error(f'Klaida: {e}')
        st.exception(e)
else:
    st.info('Įkelkite grafiką ir aktą, tada pasirinkite bloką, konkretų mėnesį ir metus.')

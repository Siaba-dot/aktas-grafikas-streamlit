
# app.py
import io
import re
from datetime import date, datetime
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook.workbook import Workbook

# ================== Page config & WOW CSS ==================
st.set_page_config(page_title="Aktas + ODS grafikas → X + Periodiškumas", page_icon="🧾", layout="wide")

NEON_PRIMARY = "#6EE7F9"   # cyan
NEON_SECOND  = "#A78BFA"   # violet
BG_GRAD_1    = "#0b0f19"
BG_GRAD_2    = "#12182b"

st.markdown(f"""
<style>
.stApp {{
  background: linear-gradient(135deg, {BG_GRAD_1} 0%, {BG_GRAD_2} 100%);
  color: #e6eefc; font-family: Inter, Segoe UI, system-ui, -apple-system, sans-serif;
}}
div[data-testid="stVerticalBlock"] > div {{
  background: rgba(255, 255, 255, 0.04);
  border: 1px solid rgba(255, 255, 255, 0.08);
  backdrop-filter: blur(6px);
  border-radius: 14px; padding: 16px 18px;
  box-shadow: 0 10px 30px rgba(0,0,0,0.25);
}}
h1, h2, h3 {{ color: #f3f6ff; position: relative; }}
h1::after, h2::after {{
  content: ""; position: absolute; left: 0; bottom: -6px; width: 92px; height: 3px;
  background: linear-gradient(90deg, {NEON_PRIMARY}, {NEON_SECOND}); border-radius: 2px;
}}
.stButton>button {{
  background: linear-gradient(90deg, {NEON_PRIMARY}33, {NEON_SECOND}33);
  color: #eaf7ff; border: 1px solid {NEON_PRIMARY}55; border-radius: 10px;
  padding: 0.55rem 1rem; transition: all .2s ease;
}}
.stButton>button:hover {{
  box-shadow: 0 0 18px {NEON_PRIMARY}88, inset 0 0 10px {NEON_SECOND}44;
  transform: translateY(-1px);
}}
[data-testid="stFileUploaderDropzone"] > div > div {{
  border: 1px dashed {NEON_PRIMARY}88 !important; background: rgba(255,255,255,0.03);
}}
thead tr th {{ background: rgba(255,255,255,0.05) !important; }}
section[data-testid="stSidebar"] {{
  background: #0e1424aa; border-right: 1px solid rgba(255,255,255,0.06);
}}
</style>
""", unsafe_allow_html=True)

st.title("🧾 ODS → X (Pn–Pn) → Periodiškumas")
st.caption("Formulės lieka. „Kaina“ – TRUNC iki 2 d. be apvalinimo (jei įjungsi). PVM – nereikia.")

# ================== Helpers ==================

HEADER_ROW_INDEX = 8  # antraštės akte yra 8 eilutėje

def norm(s: str) -> str:
    s = (str(s) if s is not None else "").strip().lower()
    s = (s.replace("ą","a").replace("č","c").replace("ę","e").replace("ė","e")
           .replace("į","i").replace("š","s").replace("ų","u").replace("ū","u").replace("ž","z")
           .replace("–","-").replace("—","-"))
    s = re.sub(r"\s+", " ", s)
    return s

# Tik darbo dienos (Mon–Fri)
WD_NAMES = ["Pirmadienis", "Antradienis", "Trečiadienis", "Ketvirtadienis", "Penktadienis"]
WD_IDX = { "pirmadienis":0, "antradienis":1, "treciadienis":2, "ketvirtadienis":3, "penktadienis":4 }

def month_weekday_counts(year: int, month: int) -> Dict[int, int]:
    import calendar
    last_day = calendar.monthrange(year, month)[1]
    counts = {i: 0 for i in range(7)}
    for d in range(1, last_day + 1):
        counts[date(year, month, d).weekday()] += 1
    # grąžinam tik Mon–Fri
    return {k: v for k, v in counts.items() if k in (0,1,2,3,4)}

def try_parse_date(val) -> Optional[date]:
    if val is None or val == "":
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    if isinstance(val, (int, float)):
        # excel serial fallback
        try:
            dt = pd.to_datetime(val, unit="D", origin="1899-12-30", errors="coerce")
            if pd.notnull(dt):
                return dt.date()
        except Exception:
            pass
    if isinstance(val, str):
        s = val.strip()
        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%Y.%m.%d", "%d.%m.%Y"):
            try:
                return datetime.strptime(s, fmt).date()
            except Exception:
                continue
        dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
        if pd.notnull(dt):
            return dt.date()
    return None

def build_header_map(ws: Worksheet, header_row: int) -> Dict[str, int]:
    m: Dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=col).value
        if v is None:
            continue
        m[norm(v)] = col
    return m

def find_end_row(ws: Worksheet, start_row: int) -> int:
    end_markers = {norm("Suma be PVM"), norm("Iš viso")}
    for r in range(start_row, ws.max_row + 1):
        row_text = " ".join(
            str(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)
            if ws.cell(r, c).value is not None
        )
        if not row_text:
            continue
        rt = norm(row_text)
        if any(em in rt for em in end_markers):
            return r - 1
    return ws.max_row

def detect_name_col(ws: Worksheet, start_row: int, end_row: int, header_map: Dict[str,int]) -> int:
    """
    Heuristika: 'paslaugos pavadinimo' stulpelis – kairėje nuo 'Mato vnt.' / 'įkainis' / 'Periodiškumas',
    kuriame daugiausiai tekstinių reikšmių.
    """
    anchors = [norm("Mato vnt."), norm("įkainis"), norm("Periodiškumas")]
    anchor_cols = [header_map[a] for a in anchors if a in header_map]
    max_anchor = min(anchor_cols) if anchor_cols else ws.max_column
    best_col, best_score = 1, -1
    for col in range(1, max_anchor):
        score = 0
        for r in range(start_row, end_row + 1):
            v = ws.cell(r, col).value
            if isinstance(v, str) and v.strip():
                score += 1
        if score > best_score:
            best_col, best_score = col, score
    return best_col

def collect_holidays_from_sheet(ws: Worksheet, header_row: int, search_limit_rows: int = 60) -> List[date]:
    holidays: List[date] = []
    target_key = norm("Šventinės dienos")
    for r in range(1, header_row + search_limit_rows):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and norm(v) == target_key:
                blanks = 0; rr = r + 1
                while rr <= min(ws.max_row, r + search_limit_rows):
                    val = ws.cell(rr, c).value
                    if val is None or (isinstance(val, str) and not val.strip()):
                        blanks += 1
                        if blanks >= 3: break
                        rr += 1; continue
                    d = try_parse_date(val)
                    if d: holidays.append(d); blanks = 0
                    rr += 1
    return sorted(set(holidays))

def wrap_formula_trunc(value: Optional[str], decimals: int = 2) -> Optional[str]:
    if not isinstance(value, str) or not value.startswith("="):
        return value
    up = value.upper()
    if "TRUNC(" in up or "ROUNDDOWN(" in up:
        return value
    return f"=TRUNC({value[1:]},{decimals})"

def read_schedule_ods(ods_file) -> pd.DataFrame:
    """Skaito ODS grafiką (pirmą lapą)."""
    return pd.read_excel(ods_file, engine="odf")

def autodetect_columns(df: pd.DataFrame) -> Tuple[Optional[str], Optional[str]]:
    """
    Bando atspėti, kuris stulpelis yra Data, o kuris – Patalpa/Paslauga/Zona.
    """
    # Data kandidatės
    date_candidates = []
    for col in df.columns:
        s = df[col].dropna()
        if s.empty: continue
        parsed = s.apply(lambda x: pd.to_datetime(x, errors="coerce", dayfirst=True))
        ratio = parsed.notna().mean()
        if ratio > 0.6:
            date_candidates.append(col)

    # Paslaugos/patalpos kandidatės – tekstiniai stulpeliai su mažai unikalių reikšmių
    svc_candidates = []
    for col in df.columns:
        s = df[col].dropna()
        if s.empty: continue
        if s.map(lambda x: isinstance(x, str)).mean() > 0.7 and s.nunique(dropna=True) <= max(200, len(s)*0.9):
            svc_candidates.append(col)

    date_col = date_candidates[0] if date_candidates else None
    svc_col  = None
    # mėgink atspėti pagal pavadinimą
    for col in df.columns:
        if norm(col) in (norm("Patalpa"), norm("Zona"), norm("Paslauga"), norm("Pavadinimas")):
            svc_col = col; break
    if not svc_col and svc_candidates:
        svc_col = svc_candidates[0]
    return date_col, svc_col

def extract_act_rows(ws: Worksheet, header_map: Dict[str,int]) -> Tuple[int, int, int, Dict[int,str], Dict[int,int]]:
    """
    Iš akto ištraukia:
      - start_row / end_row;
      - name_col (eilutės pavadinimas: 'Kabinetai', 'WC patalpos' ir pan.);
      - row_names: {row_index -> pavadinimas}
      - day_cols: {weekday_index(0..4) -> column_index}
    """
    start_row = HEADER_ROW_INDEX + 1
    end_row = find_end_row(ws, start_row)
    name_col = detect_name_col(ws, start_row, end_row, header_map)

    row_names: Dict[int, str] = {}
    for r in range(start_row, end_row + 1):
        v = ws.cell(r, name_col).value
        if isinstance(v, str) and v.strip():
            row_names[r] = v.strip()

    day_cols: Dict[int, int] = {}
    for label, wd_idx in WD_IDX.items():
        if label in header_map:
            day_cols[wd_idx] = header_map[label]

    return start_row, end_row, name_col, row_names, day_cols

def best_match(target: str, candidates: List[str]) -> Optional[str]:
    """
    Paprastas „fuzzy“: exact -> startswith -> substring (normalize).
    Vengiame papildomų paketų, kad neišpūsti requirements.
    """
    t = norm(target)
    # exact
    for c in candidates:
        if norm(c) == t:
            return c
    # startswith
    for c in candidates:
        if norm(c).startswith(t) or t.startswith(norm(c)):
            return c
    # substring
    for c in candidates:
        if t in norm(c) or norm(c) in t:
            return c
    return None

def mark_X_from_schedule(
    ws: Worksheet,
    header_map: Dict[str,int],
    year: int, month: int,
    df_sched: pd.DataFrame,
    date_col: str,
    svc_col: str,
    clear_existing: bool = True
) -> int:
    """
    Pagal ODS grafiką:
      - priderina ODS paslaugos pavadinimus prie akto eilučių;
      - kiekvienai eilutei sužymi X atitinkamuose Pn–Pn stulpeliuose (pagal to mėn. datas);
      - (pasirinktinai) išvalo senus X.
    Grąžina: kiek eilučių buvo paliesta.
    """
    start_row, end_row, name_col, row_names, day_cols = extract_act_rows(ws, header_map)
    if not day_cols:
        raise RuntimeError("Neradau Pirmadienis–Penktadienis stulpelių akte.")

    # Filtruojam tik šio mėnesio įrašus su data
    df = df_sched.copy()
    df = df.dropna(subset=[date_col, svc_col])
    df["_d"] = df[date_col].apply(lambda x: try_parse_date(x))
    df = df[df["_d"].notna()]
    df["y"] = df["_d"].apply(lambda d: d.year)
    df["m"] = df["_d"].apply(lambda d: d.month)
    df = df[(df["y"] == year) & (df["m"] == month)]

    # Mappinam ODS paslaugų pavadinimus prie akto eilučių
    act_names = list(row_names.values())
    touched_rows = 0

    # Grupė: pagal paslaugą -> sąrašas datų
    for svc, g in df.groupby(svc_col):
        match = best_match(str(svc), act_names)
        if not match:
            continue
        # Raskim to atitikmens eilutes indeksą (-us) (gali būti kelios vienodo pavadinimo)
        target_rows = [r for r, nm in row_names.items() if best_match(match, [nm]) is not None]
        if not target_rows:
            continue

        # Išvesti X pagal darbo dienas, kuriose ši paslauga yra grafike
        wd_set = set(d.weekday() for d in g["_d"])
        wd_set = {wd for wd in wd_set if wd in (0,1,2,3,4)}  # tik Pn–Pn

        for r in target_rows:
            # išvalom senus X (tik Mon–Fri), jei pasirinkta
            if clear_existing:
                for wd, col in day_cols.items():
                    ws.cell(r, col).value = None

            # uždedam X
            for wd in wd_set:
                col = day_cols.get(wd)
                if col:
                    ws.cell(r, col).value = "X"
            touched_rows += 1

    return touched_rows

def apply_periodiskumas_mon_fri(
    wb: Workbook,
    year: int, month: int,
    exclude_holidays: bool = True,
    enforce_trunc_on_kaina: bool = True,
    decimals: int = 2,
) -> Tuple[int, Dict[int, int]]:
    """
    - Skaičiuoja tik Mon–Fri.
    - (Pasirinktinai) minusuoja šventes (iš akto skyriaus „Šventinės dienos“).
    - Įrašo 'Periodiškumas' pagal pažymėtus X Mon–Fri.
    - 'Kaina' formulėms pritaiko TRUNC(...;2), jei įjungta.
    """
    ws = wb.active
    header_map = build_header_map(ws, HEADER_ROW_INDEX)

    # Mon–Fri stulpeliai
    day_cols: Dict[int, int] = {}
    for label, wd in WD_IDX.items():
        if label in header_map:
            day_cols[wd] = header_map[label]

    period_col = header_map.get(norm("Periodiškumas"))
    kaina_col  = header_map.get(norm("Kaina"))
    if not day_cols or not period_col:
        raise RuntimeError("Neradau Mon–Fri stulpelių ar 'Periodiškumas' stulpelio (8-oje eilutėje).")

    wd_counts = month_weekday_counts(year, month)

    if exclude_holidays:
        holidays = collect_holidays_from_sheet(ws, HEADER_ROW_INDEX)
        for d in holidays:
            if d.year == year and d.month == month:
                wd = d.weekday()
                if wd in wd_counts:
                    wd_counts[wd] = max(wd_counts[wd] - 1, 0)

    start_row = HEADER_ROW_INDEX + 1
    end_row = find_end_row(ws, start_row)
    updated = 0

    for r in range(start_row, end_row + 1):
        marked_wd: List[int] = []
        for wd, c in day_cols.items():
            val = ws.cell(row=r, column=c).value
            if isinstance(val, str) and val.strip().upper() == "X":
                marked_wd.append(wd)
        if not marked_wd:
            continue
        period = sum(wd_counts.get(wd, 0) for wd in marked_wd)
        ws.cell(row=r, column=period_col).value = int(period)
        updated += 1

        if enforce_trunc_on_kaina and kaina_col:
            cell = ws.cell(row=r, column=kaina_col)
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cell.value = wrap_formula_trunc(cell.value, decimals=decimals)

    return updated, wd_counts

# ================== Sidebar ==================
with st.sidebar:
    st.header("⚙️ Nustatymai")
    col1, col2 = st.columns(2)
    with col1:
        target_year = st.number_input("Metai", 2020, 2100, datetime.now().year, step=1)
    with col2:
        target_month = st.number_input("Mėnuo", 1, 12, datetime.now().month, step=1)

    exclude_holidays = st.checkbox("Neįtraukti švenčių (iš skyriaus „Šventinės dienos“)", value=True)
    enforce_trunc = st.checkbox("„Kaina“ – TRUNC iki 2 d. po kablelio (be apvalinimo)", value=True)
    clear_existing_x = st.checkbox("Perrašyti X pagal ODS (išvalyti senus)", value=True)
    obj_filter = st.text_input("Filtras pagal objektą (neprivaloma)", placeholder="pvz., Ignalina")

# ================== Main UI ==================
st.subheader("1) Įkelk aktą (.xlsx) su formulėmis ir ODS grafiką")
act_file = st.file_uploader("Aktas (Excel .xlsx)", type=["xlsx"])
ods_file = st.file_uploader("Grafikas (LibreOffice .ods)", type=["ods"])

date_col_name = None
svc_col_name = None
df_sched = None

if ods_file:
    try:
        df_sched = read_schedule_ods(ods_file)
        # Autodetekcija
        date_guess, svc_guess = autodetect_columns(df_sched)
        st.success("ODS nuskaitytas.")
        with st.expander("🔎 Peržiūra (pirmos 50 eilučių)"):
            st.dataframe(df_sched.head(50), use_container_width=True)

        st.subheader("2) Nurodyk stulpelius grafike")
        cols = list(df_sched.columns)
        date_col_name = st.selectbox("Stulpelis su data", options=cols, index=cols.index(date_guess) if date_guess in cols else 0)
        svc_col_name  = st.selectbox("Stulpelis su paslauga/patalpa/zona", options=cols, index=cols.index(svc_guess) if svc_guess in cols else 0)

        # Pasirenkamas filtras pagal objektą
        if obj_filter:
            # jei ODS turi "Objektas" stulpelį – pabandom filtruoti
            obj_cols = [c for c in cols if norm(c) in (norm("Objektas"), norm("Objektas pavadinimas"), norm("Padalinys"))]
            if obj_cols:
                df_sched = df_sched[df_sched[obj_cols[0]].astype(str).str.contains(obj_filter, case=False, na=False)]
                st.caption(f"🔍 Pridėtas objektų filtras pagal '{obj_filter}'")

    except Exception as e:
        st.exception(e)
        st.error("Nepavyko nuskaityti ODS. Patikrink, ar tai teisingas .ods failas.")

if st.button("🔄 Uždėti X pagal ODS ir perskaičiuoti periodiškumą", type="primary", use_container_width=True):
    if not act_file or df_sched is None or not date_col_name or not svc_col_name:
        st.warning("Įkelk **aktą (.xlsx)**, **ODS** ir nurodyk **Data** bei **Paslauga** stulpelius.")
        st.stop()
    try:
        with st.spinner("Atidarau aktą..."):
            wb = load_workbook(filename=act_file, data_only=False)

        ws = wb.active
        header_map = build_header_map(ws, HEADER_ROW_INDEX)

        with st.spinner("Žymiu X pagal ODS..."):
            touched = mark_X_from_schedule(
                ws=ws,
                header_map=header_map,
                year=int(target_year), month=int(target_month),
                df_sched=df_sched,
                date_col=date_col_name,
                svc_col=svc_col_name,
                clear_existing=clear_existing_x
            )

        with st.spinner("Skaičiuoju „Periodiškumą“ (Pn–Pn)..."):
            updated, wd_counts = apply_periodiskumas_mon_fri(
                wb=wb,
                year=int(target_year),
                month=int(target_month),
                exclude_holidays=exclude_holidays,
                enforce_trunc_on_kaina=enforce_trunc,
                decimals=2,
            )

        out = io.BytesIO()
        wb.save(out); out.seek(0)
        label = f"{int(target_year)}-{int(target_month):02d}"
        st.success(f"X pažymėta {touched} eilutėse. Periodiškumas atnaujintas {updated} eilutėse. ({label})")
        st.json({
            "Pirmadienių": wd_counts.get(0,0),
            "Antradienių": wd_counts.get(1,0),
            "Trečiadienių": wd_counts.get(2,0),
            "Ketvirtadienių": wd_counts.get(3,0),
            "Penktadienių": wd_counts.get(4,0),
        })

        st.download_button(
            "⬇️ Parsisiųsti atnaujintą aktą",
            data=out,
            file_name=f"Aktas_atnaujintas_{label}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    except Exception as e:
        st.exception(e)
        st.error("Nepavyko pažymėti X / perskaičiuoti. Patikrink ODS stulpelius ir akto antraštes (8 eilutę).")

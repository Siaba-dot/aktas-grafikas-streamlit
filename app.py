
# app.py
# Streamlit: Akto atnaujinimas (.xls/.xlsx) su EXCEL FORMULĖMIS
# - Periodiškumas skaičiuojamas FORMULĖMIS pagal kalendorių (Pir–Sek) ir atimant šventines (Skaičiavimai!E2:E200)
# - Kaina = Plotas kv m./kiekis/val × įkainis × Periodiškumas (be ROUND; rodymas 0.00)
# - PVM netaikomas
# - WOW dark neon UI

import streamlit as st
import pandas as pd
import io
import tempfile
import calendar
from datetime import datetime as dt
from openpyxl.utils import get_column_letter

# ---------- PAGE CONFIG + WOW CSS ----------
st.set_page_config(page_title="Akto atnaujinimas iš grafiko", layout="wide")

def inject_wow_css(accent="#00FF88"):
    st.markdown(
        f"""
        <style>
        .stApp {{
            background: #0b1020;
            color: #e6e6ea;
        }}
        .stButton>button {{
            background: linear-gradient(90deg, {accent}, #7C3AED);
            color: white; border: 0; border-radius: 8px; padding: 0.6rem 1rem; font-weight: 600;
        }}
        .stDownloadButton>button {{
            background: linear-gradient(90deg, #7C3AED, {accent});
            color: white; border: 0; border-radius: 8px; padding: 0.6rem 1rem; font-weight: 600;
        }}
        .stTextInput>div>div>input, .stNumberInput input {{
            background: #131a33; color: #e6e6ea; border: 1px solid #213157;
        }}
        .stDataFrame [data-testid="stTable"] {{
            background: #0f1430;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )

with st.sidebar:
    st.subheader("🎨 Dizaino akcentas")
    accent = st.color_picker("Pasirinkite akcento spalvą", "#00FF88")
inject_wow_css(accent)

st.title("Akto atnaujinimas · Excel formulės (be statinių reikšmių)")
st.caption("Įkainiai, kiekiai ir periodiškumai – tik per Excel formules. PVM netaikomas.")

# ---------- FILE UPLOAD ----------
act_file = st.file_uploader("Įkelkite akto failą (.xls/.xlsx)", type=["xls", "xlsx"])
year_choice = st.number_input("Metai", min_value=2000, max_value=2100, value=dt.now().year)
month_choice = st.number_input("Mėnuo (1–12)", min_value=1, max_value=12, value=dt.now().month)

if not act_file:
    st.info("Įkelkite akto failą ir nurodykite metus/mėnesį.")
    st.stop()

# ---------- READ ACT ----------
def save_to_temp(uploaded_file, suffix: str):
    t = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    t.write(uploaded_file.read()); t.flush(); t.close()
    return t.name

path = save_to_temp(act_file, "." + act_file.name.split(".")[-1].lower())
engine = "openpyxl" if path.endswith(".xlsx") else "xlrd"
xls = pd.ExcelFile(path, engine=engine)
act_sheet = st.selectbox("Pasirinkite akto lapą", options=xls.sheet_names)
act_df = xls.parse(act_sheet, header=0)

st.subheader("Akto peržiūra")
st.dataframe(act_df.head(15), use_container_width=True)

# ---------- REQUIRED COLUMNS ----------
headers = [str(c).strip() for c in act_df.columns]
def col_index_exact(col_name: str) -> int:
    if col_name in headers:
        return headers.index(col_name) + 1  # 1-based index for openpyxl addresses
    else:
        return -1

col_plotas  = col_index_exact("Plotas kv m./kiekis/val")
col_ikainis = col_index_exact("įkainis")
col_period  = col_index_exact("Periodiškumas")
col_kaina   = col_index_exact("Kaina")

missing = []
for name, idx in [("Plotas kv m./kiekis/val", col_plotas),
                  ("įkainis", col_ikainis),
                  ("Periodiškumas", col_period),
                  ("Kaina", col_kaina)]:
    if idx == -1:
        missing.append(name)
if missing:
    st.error(f"Trūksta šių stulpelių akte: {', '.join(missing)}. Įsitikink pavadinimais (tiksliai!).")
    st.stop()

# Weekday columns (optional, jei yra X žymos)
weekday_cols_map = {
    "Pirmadienis": None, "Antradienis": None, "Trečiadienis": None,
    "Ketvirtadienis": None, "Penktadienis": None, "Šeštadienis": None, "Sekmadienis": None
}
for name in list(weekday_cols_map.keys()):
    weekday_cols_map[name] = col_index_exact(name)

has_weekday_x = any(idx != -1 for idx in weekday_cols_map.values())

st.info("Formulėms pritaikytas akto struktūros žemėlapis:")
st.code(f"""
Plotas kv m./kiekis/val -> kol. #{col_plotas}
įkainis                  -> kol. #{col_ikainis}
Periodiškumas            -> kol. #{col_period}
Kaina                    -> kol. #{col_kaina}
Savaitės dienos          -> {[name for name, idx in weekday_cols_map.items() if idx != -1]}
""")

# ---------- WRITE EXCEL WITH FORMULAS ----------
with st.spinner("Generuoju Excel su formulėmis..."):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 1) Rašome AKTO lapą (duomenis kaip yra)
        act_df.to_excel(writer, index=False, sheet_name=act_sheet)

        # 2) Kuriame SKAIČIAVIMAI lapą: Metai/Mėnuo + formulių bazė
        wb = writer.book
        calc_ws = wb.create_sheet("Skaičiavimai")

        # Įvestys (ne „apskaičiuotos“): metai/mėnuo
        calc_ws["A1"] = "Metai"; calc_ws["B1"] = "Mėnuo"
        calc_ws["A2"] = int(year_choice); calc_ws["B2"] = int(month_choice)

        # Vietos šventinėms (įvestys vartotojo): E2:E200
        calc_ws["D1"] = "Pastaba"
        calc_ws["D2"] = "Į šį stulpelį įveskite šventinių datų sąrašą (YYYY-MM-DD)."
        calc_ws["E1"] = "Šventinės datos (E2:E200)"
        for r in range(2, 201):
            calc_ws[f"E{r}"] = None  # paliekam tuščia kaip įvestį

        # Weekday antraštės (Pir..Sek) ir FORMULĖS C2:I2
        headers_w = ["Pir", "An", "Tre", "Ket", "Pen", "Šeš", "Sek"]
        for i, h in enumerate(headers_w):
            calc_ws.cell(row=1, column=3+i, value=h)  # C1..I1

        # Dinaminės 365 formulės C2..I2: kiek kartų tenka atitinkama savaitės diena mėnesyje, atėmus šventines
        # Paaiškinimas:
        # dates = visos mėnesio datos; week = WEEKDAY(dates,2) -> 1..7 (Pir..Sek)
        # hol = FILTER(E2:E200, MONTH(E2:E200)=m, "") -> šventinės to mėnesio dienos
        # COUNTIF(week, wd) - COUNTIF(WEEKDAY(hol,2), wd)
        letters = ["C","D","E","F","G","H","I"]
        for wd_num, col_letter in enumerate(letters, start=1):
            calc_ws[f"{col_letter}2"] = (
                f'=LET('
                f'y,$A$2, m,$B$2, '
                f'days,SEQUENCE(DAY(EOMONTH(DATE(y,m,1),0))), '
                f'dates,DATE(y,m,days), '
                f'week,WEEKDAY(dates,2), '
                f'hol,FILTER($E$2:$E$200,(MONTH($E$2:$E$200)=m)*(YEAR($E$2:$E$200)=y),""), '
                f'COUNTIF(week,{wd_num})-COUNTIF(WEEKDAY(hol,2),{wd_num}) '
                f')'
            )

        # Perskaičiavimas atidarant
        try:
            wb.calculation.fullCalcOnLoad = True
        except Exception:
            pass

        # 3) Įrašome FORMULES į AKTO lapą
        ws = wb[act_sheet]
        last_row = ws.max_row

        def addr(ci, ri):
            return f"{get_column_letter(ci)}{ri}"

        # SUM viršuje (nebūtina, bet patogu): į 7-ą eilutę įdedame SUM formules (jei tokia yra).
        # Jei 7 eilutė neužimta, paliksime kaip yra — tai netrukdo.
        sum_row = 7
        # SUM Periodiškumas
        ws[f"{get_column_letter(col_period)}{sum_row}"] = f"=SUM({get_column_letter(col_period)}2:{get_column_letter(col_period)}{last_row})"
        ws[f"{get_column_letter(col_period)}{sum_row}"].number_format = "0.00"
        # SUM Kaina
        ws[f"{get_column_letter(col_kaina)}{sum_row}"] = f"=SUM({get_column_letter(col_kaina)}2:{get_column_letter(col_kaina)}{last_row})"
        ws[f"{get_column_letter(col_kaina)}{sum_row}"].number_format = "0.00"
        # SUM Plotas (jei reikia)
        ws[f"{get_column_letter(col_plotas)}{sum_row}"] = f"=SUM({get_column_letter(col_plotas)}2:{get_column_letter(col_plotas)}{last_row})"
        ws[f"{get_column_letter(col_plotas)}{sum_row}"].number_format = "0.00"

        # Eilučių formulės (nuo 2-os eil.)
        for r in range(2, last_row+1):
            plot_cell   = addr(col_plotas,  r)
            rate_cell   = addr(col_ikainis, r)
            period_cell = addr(col_period,  r)
            price_cell  = addr(col_kaina,   r)

            # PERIODIŠKUMAS:
            # Jei yra savaitės dienų X-stulpeliai, sumuojame IF(X, Skaičiavimai!C2..I2, 0).
            # Jei NĖRA, taikom fallback: IF(LOWER(Periodiškumas)="x",1,IFERROR(VALUE(Periodiškumas),0))
            if has_weekday_x:
                parts = []
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
                        # C2..I2 atitinka wd_num=1..7
                        base_cell = f"Skaičiavimai!{letters[wd_num-1]}2"
                        parts.append(f'IF({wd_cell}="X",{base_cell},0)')
                ws[period_cell] = "=" + ("+".join(parts) if parts else "0")
            else:
                ws[period_cell] = f'=IF(LOWER({period_cell})="x",1,IFERROR(VALUE({period_cell}),0))'

            # KAINA = Plotas * Įkainis * Periodiškumas (be ROUND; formatas 0.00 tik rodymui)
            ws[price_cell] = f'=IFERROR(VALUE({plot_cell}),{plot_cell})*IFERROR(VALUE({rate_cell}),{rate_cell})*IFERROR(VALUE({period_cell}),{period_cell})'
            ws[price_cell].number_format = "0.00"

    output.seek(0)
    out_name = f"Aktas_atnaujintas_{year_choice}_{month_choice:02d}.xlsx"
    st.download_button(
        label="Atsisiųsti atnaujintą aktą (.xlsx)",
        data=output,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.success("✅ Paruošta. Akto faile visos skaičiuojamos reikšmės (Periodiškumas, Kaina, SUM) yra Excel formulės.")
st.info("Šventines įveskite į Skaičiavimai!E2:E200 (YYYY-MM-DD). Formulės automatiškai jas atims iš mėnesio dienų. PVM – netaikomas.")

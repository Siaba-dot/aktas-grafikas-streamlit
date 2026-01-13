
# app.py — Akto atnaujinimas su EXCEL FORMULĖMIS
# - Antraštės (headers) yra 8-oje eilutėje: header=7 (pandas) ir startrow=7 (rašymui)
# - Kaina = Plotas * Įkainis * Periodiškumas (be ROUND; rodymas 0.00)
# - Periodiškumas iš savaitės X (jei yra) arba fallback iš teksto
# - PVM netaikomas. Jokios galutinės reikšmės nerašomos kaip skaičiai/tekstai — tik formulės.

import streamlit as st
import pandas as pd
import io
import tempfile
import calendar
from datetime import datetime as dt, date
from openpyxl.utils import get_column_letter

HEADER_ROW = 8            # kur yra antraštės (1-based)
DATA_START_ROW = HEADER_ROW + 1  # pirmoji duomenų eilutė (1-based)

st.set_page_config(page_title="Akto atnaujinimas (formulės)", layout="wide")
st.title("Akto atnaujinimas · formulės iš 8-os eilutės antraščių")

# --- Įkėlimas ---
act_file = st.file_uploader("Įkelkite akto failą (.xls/.xlsx)", type=["xls","xlsx"])
year_choice = st.number_input("Metai", min_value=2000, max_value=2100, value=dt.now().year)
month_choice = st.number_input("Mėnuo (1–12)", min_value=1, max_value=12, value=dt.now().month)

def save_to_temp(uploaded_file, suffix: str):
    t = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    t.write(uploaded_file.read()); t.flush(); t.close()
    return t.name

if not act_file:
    st.info("Įkelkite akto failą ir nurodykite metus/mėnesį.")
    st.stop()

# --- Nuskaitymas: header=7 (8-oji eilutė) ---
path = save_to_temp(act_file, "." + act_file.name.split(".")[-1].lower())
engine = "openpyxl" if path.endswith(".xlsx") else "xlrd"
xls = pd.ExcelFile(path, engine=engine)
act_sheet = st.selectbox("Pasirinkite akto lapą", options=xls.sheet_names)

# Svarbu: header=HEADER_ROW-1, kad pandas kolonų pavadinimus paimtų iš 8-os eilutės
act_df = xls.parse(act_sheet, header=HEADER_ROW-1)

st.subheader("Akto peržiūra (duomenys nuo 9-os eil.)")
st.dataframe(act_df.head(15), use_container_width=True)

# --- Privalomi stulpeliai (tikslūs pavadinimai iš 8-os eil.) ---
headers = [str(c).strip() for c in act_df.columns]
def col_index_exact(col_name: str) -> int:
    try:
        return headers.index(col_name) + 1   # 1-based index (Excel col)
    except ValueError:
        return -1

col_plotas  = col_index_exact("Plotas kv m./kiekis/val")
col_ikainis = col_index_exact("įkainis")
col_period  = col_index_exact("Periodiškumas")
col_kaina   = col_index_exact("Kaina")

weekday_cols_map = {
    "Pirmadienis": col_index_exact("Pirmadienis"),
    "Antradienis": col_index_exact("Antradienis"),
    "Trečiadienis": col_index_exact("Trečiadienis"),
    "Ketvirtadienis": col_index_exact("Ketvirtadienis"),
    "Penktadienis": col_index_exact("Penktadienis"),
    "Šeštadienis": col_index_exact("Šeštadienis"),
    "Sekmadienis": col_index_exact("Sekmadienis"),
}
has_weekday_x = any(idx != -1 for idx in weekday_cols_map.values())

missing = []
for name, idx in [("Plotas kv m./kiekis/val", col_plotas),
                  ("įkainis", col_ikainis),
                  ("Periodiškumas", col_period),
                  ("Kaina", col_kaina)]:
    if idx == -1:
        missing.append(name)
if missing:
    st.error("Trūksta stulpelių (8-oje eilutėje): " + ", ".join(missing))
    st.stop()

st.info("Rasta stulpelių atitikmenys (8-oje eilutėje):")
st.code(f"""
Plotas kv m./kiekis/val -> kol. #{col_plotas}
įkainis                  -> kol. #{col_ikainis}
Periodiškumas            -> kol. #{col_period}
Kaina                    -> kol. #{col_kaina}
Savaitės dienos          -> {[name for name, idx in weekday_cols_map.items() if idx != -1]}
""")

# --- Rašymas su FORMULĖMIS ---
with st.spinner("Generuoju Excel su formulėmis..."):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 1) Rašome akto lentelę taip, kad header būtų 8-oje eilutėje (startrow=7)
        #    Pandas pats įrašys antraštes į 8-ą eil. ir duomenis nuo 9-os.
        act_df.to_excel(writer, index=False, sheet_name=act_sheet, startrow=HEADER_ROW-1)

        wb = writer.book
        ws = wb[act_sheet]

        # 2) Skaičiavimai: savaitės dienų kiekiai mėnesyje minus šventinės
        calc_ws = wb.create_sheet("Skaičiavimai")
        calc_ws["A1"] = "Metai"; calc_ws["B1"] = "Mėnuo"
        calc_ws["A2"] = int(year_choice); calc_ws["B2"] = int(month_choice)
        calc_ws["E1"] = "Šventinės (E2:E200)"
        calc_ws["D2"] = "Įveskite YYYY-MM-DD; bus atimta iš mėnesio dienų."

        # Pir..Sek į C1..I1; formulės į C2..I2 (365 LET/FILTER)
        headers_w = ["Pir", "An", "Tre", "Ket", "Pen", "Šeš", "Sek"]
        letters   = ["C","D","E","F","G","H","I"]
        for i, h in enumerate(headers_w):
            calc_ws.cell(row=1, column=3+i, value=h)  # C1..I1
            # wd_num = i+1
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

        # 3) Formulės eilutėms: nuo DATA_START_ROW (9-oji eil.) iki (9 + len(act_df) - 1)
        n_data = len(act_df)
        first_row = DATA_START_ROW
        last_row  = DATA_START_ROW + n_data - 1

        def addr(col_idx_1_based: int, row_1_based: int) -> str:
            return f"{get_column_letter(col_idx_1_based)}{row_1_based}"

        # SUM eilutė viršuje (7-oje): patogu turėti sumas
        sum_row = HEADER_ROW - 1  # 7-oji eilė
        # SUM Periodiškumas
        ws[addr(col_period, sum_row)] = f"=SUM({addr(col_period, first_row)}:{addr(col_period, last_row)})"
        ws[addr(col_period, sum_row)].number_format = "0.00"
        # SUM Kaina
        ws[addr(col_kaina, sum_row)] = f"=SUM({addr(col_kaina, first_row)}:{addr(col_kaina, last_row)})"
        ws[addr(col_kaina, sum_row)].number_format = "0.00"
        # SUM Plotas
        ws[addr(col_plotas, sum_row)] = f"=SUM({addr(col_plotas, first_row)}:{addr(col_plotas, last_row)})"
        ws[addr(col_plotas, sum_row)].number_format = "0.00"

        # Eilučių formulės
        for r in range(first_row, last_row+1):
            plot_cell   = addr(col_plotas,  r)
            rate_cell   = addr(col_ikainis, r)
            period_cell = addr(col_period,  r)
            price_cell  = addr(col_kaina,   r)

            # PERIODIŠKUMAS:
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
                        base_cell = f"Skaičiavimai!{letters[wd_num-1]}2"  # C2..I2
                        parts.append(f'IF({wd_cell}="X",{base_cell},0)')
                ws[period_cell] = "=" + ("+".join(parts) if parts else "0")
            else:
                # fallback: X -> 1; skaičius tekste -> VALUE; kita -> 0
                ws[period_cell] = f'=IF(LOWER({period_cell})="x",1,IFERROR(VALUE({period_cell}),0))'

            # KAINA = Plotas * Įkainis * Periodiškumas (be ROUND; formatas 0.00 rodymui)
            ws[price_cell] = (
                f'=IFERROR(VALUE({plot_cell}),{plot_cell})'
                f'*IFERROR(VALUE({rate_cell}),{rate_cell})'
                f'*IFERROR(VALUE({period_cell}),{period_cell})'
            )
            ws[price_cell].number_format = "0.00"

    output.seek(0)
    out_name = f"Aktas_atnaujintas_{year_choice}_{month_choice:02d}.xlsx"
    st.download_button("Atsisiųsti atnaujintą aktą (.xlsx)", data=output, file_name=out_name,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.success("✅ Paruošta: antraštės 8-oje eil., formulės nuo 9-os eil., jokios statinės „final“ reikšmės.")
st.info("Įrašyk šventines datas į Skaičiavimai!E2:E200 (YYYY-MM-DD). Formulės automatiškai jas eliminuos.")

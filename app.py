
# app.py — Akto atnaujinimas su EXCEL FORMULĖMIS
# - Antraštės suformuojamos iš sujungtų 7 ir 8 eilučių (ffill + bfill per stulpelius)
# - Kaina = Plotas * Įkainis * Periodiškumas (be ROUND; rodymas 0.00)
# - Periodiškumas iš savaitės X (jei yra) arba fallback iš teksto
# - PVM netaikomas. Jokios „final“ statinių reikšmių — tik formulės.

import streamlit as st
import pandas as pd
import io
import tempfile
import calendar
from datetime import datetime as dt
from openpyxl.utils import get_column_letter

HEADER_TOP_ROW = 7              # 1-based: viršutinė sujungtų antraščių eilutė
HEADER_BOTTOM_ROW = 8           # 1-based: apatinė sujungtų antraščių eilutė
DATA_START_ROW = HEADER_BOTTOM_ROW + 1  # 9-oji eilė: duomenys

st.set_page_config(page_title="Akto atnaujinimas (formulės)", layout="wide")
st.title("Akto atnaujinimas · formulės (antraštės iš 7–8 eilučių)")

# -------- Pagalbinės ----------
def save_to_temp(uploaded_file, suffix: str):
    t = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    t.write(uploaded_file.read()); t.flush(); t.close()
    return t.name

def build_headers_from_two_rows(df_raw: pd.DataFrame, top_row_1based: int, bottom_row_1based: int):
    """
    Suformuoja antraštes iš dviejų eilučių (pvz., 7 ir 8), taikant ffill+bfill,
    kad sujungtų/antraštiniai blokai būtų teisingai parsiųsti.
    Grąžina (df_data, headers)
    """
    # Konvertuojam į 0-based indeksą
    r_top = top_row_1based - 1
    r_bot = bottom_row_1based - 1

    # Paimam abi antraščių eilutes ir pritaikom ffill+bfill per stulpelius
    top = df_raw.iloc[[r_top, r_bot], :].copy().fillna("")
    # transponuojam, kad darytume ffill/bfill "žemyn" stulpelio viduje
    tb = top.T
    tb_ff = tb.ffill()   # jei tekstas įrašytas tik 7 eilutėje (viršuje)
    tb_full = tb_ff.bfill()  # jei tekstas įrašytas tik 8 eilutėje (apačioje)
    top_full = tb_full.T

    # Sukuriam galutinį antraštės pavadinimą kiekvienam stulpeliui:
    headers = []
    for c in range(top_full.shape[1]):
        parts = [str(top_full.iat[r, c]).strip() for r in range(top_full.shape[0])]
        parts_clean = [p for p in parts if p]
        # Jei abi eilutės turi tekstą, sujungiame su „ / “; jei viena — imam tą vieną
        col_name = " / ".join(parts_clean) if parts_clean else f"Col_{c}"
        headers.append(col_name)

    # Duomenys prasideda nuo bottom_row + 1 (t.y. 9-a eilė)
    df_data = df_raw.iloc[bottom_row_1based:, :].reset_index(drop=True)
    df_data.columns = headers
    return df_data, headers

# -------- Įkėlimas ----------
act_file = st.file_uploader("Įkelkite akto failą (.xls/.xlsx)", type=["xls","xlsx"])
year_choice = st.number_input("Metai", min_value=2000, max_value=2100, value=dt.now().year)
month_choice = st.number_input("Mėnuo (1–12)", min_value=1, max_value=12, value=dt.now().month)

if not act_file:
    st.info("Įkelkite akto failą ir nurodykite metus/mėnesį.")
    st.stop()

# -------- Nuskaitymas: header=None, o antraštes formuojam patys iš 7+8 eil. --------
path = save_to_temp(act_file, "." + act_file.name.split(".")[-1].lower())
engine = "openpyxl" if path.endswith(".xlsx") else "xlrd"
xls = pd.ExcelFile(path, engine=engine)
act_sheet = st.selectbox("Pasirinkite akto lapą", options=xls.sheet_names)

df_raw = xls.parse(act_sheet, header=None)
act_df, act_headers = build_headers_from_two_rows(df_raw, HEADER_TOP_ROW, HEADER_BOTTOM_ROW)

st.subheader("Akto peržiūra (duomenys nuo 9-os eil.)")
st.dataframe(act_df.head(15), use_container_width=True)

# -------- Po suformavimo ieškome TIKSLIŲ pavadinimų --------
# Kad atitiktų tavo laukus, tikriname tiek tikslų pavadinimą, tiek galimą sujungtą (pvz., 'Plotas kv m./kiekis/val' arba 'Plotas kv m./kiekis/val / ...')
def find_col_exact_or_prefix(headers_list, target):
    # tiksli atitiktis
    for i, h in enumerate(headers_list):
        if h.strip() == target:
            return i + 1  # 1-based
    # jei sujungta „X / Y“, ieškome pirmojo segmento
    for i, h in enumerate(headers_list):
        if h.strip().startswith(target):
            return i + 1
    return -1

col_plotas  = find_col_exact_or_prefix(act_headers, "Plotas kv m./kiekis/val")
col_ikainis = find_col_exact_or_prefix(act_headers, "įkainis")
col_period  = find_col_exact_or_prefix(act_headers, "Periodiškumas")
col_kaina   = find_col_exact_or_prefix(act_headers, "Kaina")

weekday_map_targets = ["Pirmadienis", "Antradienis", "Trečiadienis", "Ketvirtadienis", "Penktadienis", "Šeštadienis", "Sekmadienis"]
weekday_cols_map = {t: find_col_exact_or_prefix(act_headers, t) for t in weekday_map_targets}
has_weekday_x = any(idx != -1 for idx in weekday_cols_map.values())

missing = []
for name, idx in [("Plotas kv m./kiekis/val", col_plotas),
                  ("įkainis", col_ikainis),
                  ("Periodiškumas", col_period),
                  ("Kaina", col_kaina)]:
    if idx == -1:
        missing.append(name)
if missing:
    st.error("Trūksta stulpelių (po 7–8 eil. suformavimo): " + ", ".join(missing))
    st.stop()

st.info("Rasta stulpelių atitikmenys (po 7–8 eil. suformavimo):")
st.code(f"""
Plotas kv m./kiekis/val -> kol. #{col_plotas}
įkainis                  -> kol. #{col_ikainis}
Periodiškumas            -> kol. #{col_period}
Kaina                    -> kol. #{col_kaina}
Savaitės dienos          -> {[name for name, idx in weekday_cols_map.items() if idx != -1]}
""")

# -------- Rašymas su FORMULĖMIS --------
with st.spinner("Generuoju Excel su formulėmis..."):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 1) Rašome akto lentelę nuo 8-os eilės (startrow=HEADER_BOTTOM_ROW-1 → 7), kad antraštės liktų 8-oje,
        #    o duomenys — nuo 9-os.
        act_df.to_excel(writer, index=False, sheet_name=act_sheet, startrow=HEADER_BOTTOM_ROW-1)

        wb = writer.book
        ws = wb[act_sheet]

        # 2) Sukuriame "Skaičiavimai" lapą su 365 formulėmis (Pir..Sek kiekiai mėnesyje, atėmus šventines)
        calc_ws = wb.create_sheet("Skaičiavimai")
        calc_ws["A1"] = "Metai"; calc_ws["B1"] = "Mėnuo"
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

        # 3) Formulės eilutėms: nuo DATA_START_ROW iki DATA_START_ROW + len(act_df) - 1
        n_data = len(act_df)
        first_row = DATA_START_ROW
        last_row  = DATA_START_ROW + n_data - 1

        def addr(col_idx_1based: int, row_1based: int) -> str:
            return f"{get_column_letter(col_idx_1based)}{row_1based}"

        # SUM viršuje (7-oje): Periodiškumas, Kaina, Plotas
        sum_row = HEADER_BOTTOM_ROW - 1  # 7-oji eilė
        ws[addr(col_period, sum_row)] = f"=SUM({addr(col_period, first_row)}:{addr(col_period, last_row)})"
        ws[addr(col_period, sum_row)].number_format = "0.00"
        ws[addr(col_kaina, sum_row)]  = f"=SUM({addr(col_kaina,  first_row)}:{addr(col_kaina,  last_row)})"
        ws[addr(col_kaina, sum_row)].number_format  = "0.00"
        ws[addr(col_plotas, sum_row)] = f"=SUM({addr(col_plotas, first_row)}:{addr(col_plotas, last_row)})"
        ws[addr(col_plotas, sum_row)].number_format = "0.00"

        # Eilučių FORMULĖS
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

            # KAINA = Plotas * Įkainis * Periodiškumas (be ROUND; tik rodymas 0.00)
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

st.success("✅ Paruošta: antraštes sujungtos iš 7–8 eilučių, formulės nuo 9-os eil., jokios statinės „final“ reikšmės.")
st.info("Šventines datas įrašykite į Skaičiavimai!E2:E200 (YYYY-MM-DD). Formulės jas automatiškai eliminuos.")

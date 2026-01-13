
# app.py — Akto atnaujinimas su EXCEL FORMULĖMIS + grafiko (.ods/.xlsx/.xls) įkėlimas
# - Akto antraštės suformuojamos iš sujungtų 7–8 eilučių (ffill + bfill)
# - Formulės rašomos nuo 9-os eilutės; SUM po lentelės
# - Periodiškumas iš savaitės dienų X arba fallback iš teksto
# - Kaina = Plotas * Įkainis * Periodiškumas (be ROUND; rodymas 0.00)
# - PVM netaikomas

import streamlit as st
import pandas as pd
import io
import tempfile
import calendar
from datetime import datetime as dt
from openpyxl.utils import get_column_letter

# Konstanta: antraščių eilutės (Excel, 1-based)
HEADER_TOP_ROW = 7             # viršutinė sujungtų antraščių eilė
HEADER_BOTTOM_ROW = 8          # apatinė sujungtų antraščių eilė
DATA_START_ROW = HEADER_BOTTOM_ROW + 1  # 9-oji eilė: duomenys

st.set_page_config(page_title="Akto atnaujinimas (formulės)", layout="wide")
st.title("Akto atnaujinimas · Excel formulės (grafikas + aktas)")

# ---------------- Pagalbinės ----------------
def save_to_temp(uploaded_file, suffix: str):
    """Išsaugo įkeltą failą į laikiną vietą, grąžina kelią."""
    t = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    t.write(uploaded_file.read()); t.flush(); t.close()
    return t.name

def build_headers_from_two_rows(df_raw: pd.DataFrame, top_row_1based: int, bottom_row_1based: int):
    """
    Suformuoja antraštes iš dviejų eilučių (pvz., 7 ir 8), taikant ffill+bfill per stulpelį.
    Grąžina (df_data, headers).
    """
    r_top = top_row_1based - 1
    r_bot = bottom_row_1based - 1

    # paimame 7 ir 8 eil. kaip "header bloką"
    top = df_raw.iloc[[r_top, r_bot], :].copy().fillna("")
    tb = top.T
    tb_ff = tb.ffill()           # jei tekstas tik 7 eil.
    tb_full = tb_ff.bfill()      # jei tekstas tik 8 eil.
    top_full = tb_full.T

    headers = []
    for c in range(top_full.shape[1]):
        parts = [str(top_full.iat[r, c]).strip() for r in range(top_full.shape[0])]
        parts_clean = [p for p in parts if p]
        col_name = " / ".join(parts_clean) if parts_clean else f"Col_{c}"
        headers.append(col_name)

    # duomenys nuo 9-os eilės
    df_data = df_raw.iloc[bottom_row_1based:, :].reset_index(drop=True)
    df_data.columns = headers
    return df_data, headers

def find_col_exact_or_prefix(headers_list, target):
    """Ieško tikslios antraštės arba prefikso (kai sujungta „X / Y“). Grąžina 1-based indeksą arba -1."""
    for i, h in enumerate(headers_list):
        if h.strip() == target:
            return i + 1
    for i, h in enumerate(headers_list):
        if h.strip().startswith(target):
            return i + 1
    return -1

# ---------------- Įkėlimas ----------------
st.subheader("1) Įkelkite grafiko ir akto failus")
schedule_file = st.file_uploader("Grafikas (.ods/.xlsx/.xls)", type=["ods","xlsx","xls"])
act_file      = st.file_uploader("Aktas (.xls/.xlsx)",          type=["xls","xlsx"])
year_choice   = st.number_input("Metai", min_value=2000, max_value=2100, value=dt.now().year)
month_choice  = st.number_input("Mėnuo (1–12)", min_value=1, max_value=12, value=dt.now().month)

if not (schedule_file and act_file):
    st.info("Įkelkite abu failus ir nurodykite metus/mėnesį.")
    st.stop()

# ---------------- Grafiko nuskaitymas (peržiūrai) ----------------
st.subheader("2) Grafiko peržiūra")
def parse_schedule_raw(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith(".ods"):
        path = save_to_temp(uploaded_file, ".ods")
        df = pd.read_excel(path, engine="odf", header=None)
        sheet = None
        return df, sheet
    elif name.endswith(".xlsx"):
        path = save_to_temp(uploaded_file, ".xlsx")
        xls = pd.ExcelFile(path, engine="openpyxl")
        sheet = st.selectbox("Grafiko lapas", options=xls.sheet_names)
        df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl", header=None)
        return df, sheet
    elif name.endswith(".xls"):
        path = save_to_temp(uploaded_file, ".xls")
        xls = pd.ExcelFile(path, engine="xlrd")
        sheet = st.selectbox("Grafiko lapas", options=xls.sheet_names)
        df = pd.read_excel(path, sheet_name=sheet, engine="xlrd", header=None)
        return df, sheet
    else:
        raise ValueError("Nepalaikomas grafiko formatas. Naudokite .ods, .xlsx arba .xls")

raw_sched_df, sched_sheet = parse_schedule_raw(schedule_file)
st.caption(f"Grafiko forma: {raw_sched_df.shape[0]} eil. × {raw_sched_df.shape[1]} kol.")
st.dataframe(raw_sched_df.head(20), use_container_width=True)

# (jei norėtum – čia galima pridėti 7–8 eilučių header wizard grafike, bet šiuo metu grafikas nenaudojamas skaičiavimui)

# ---------------- Akto nuskaitymas su 7–8 eil. header ----------------
st.subheader("3) Akto peržiūra (antraštės 7–8 eil.)")
path_act = save_to_temp(act_file, "." + act_file.name.split(".")[-1].lower())
engine_act = "openpyxl" if path_act.endswith(".xlsx") else "xlrd"
xls_act = pd.ExcelFile(path_act, engine=engine_act)
act_sheet_name = st.selectbox("Akto lapas", options=xls_act.sheet_names)

# nuskaitymas be header; suformuojame antraštes iš 7+8
df_raw_act = xls_act.parse(act_sheet_name, header=None)
act_df, act_headers = build_headers_from_two_rows(df_raw_act, HEADER_TOP_ROW, HEADER_BOTTOM_ROW)
st.dataframe(act_df.head(15), use_container_width=True)

# Reikalingi stulpeliai (tikslūs tavo pavadinimai)
col_plotas  = find_col_exact_or_prefix(act_headers, "Plotas kv m./kiekis/val")
col_ikainis = find_col_exact_or_prefix(act_headers, "įkainis")
col_period  = find_col_exact_or_prefix(act_headers, "Periodiškumas")
col_kaina   = find_col_exact_or_prefix(act_headers, "Kaina")

weekday_targets = ["Pirmadienis","Antradienis","Trečiadienis","Ketvirtadienis","Penktadienis","Šeštadienis","Sekmadienis"]
weekday_cols_map = {t: find_col_exact_or_prefix(act_headers, t) for t in weekday_targets}
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

st.info("Rasta stulpelių atitikmenys:")
st.code(f"""
Plotas kv m./kiekis/val -> kol. #{col_plotas}
įkainis                  -> kol. #{col_ikainis}
Periodiškumas            -> kol. #{col_period}
Kaina                    -> kol. #{col_kaina}
Savaitės dienos          -> {[name for name, idx in weekday_cols_map.items() if idx != -1]}
""")

# ---------------- Eksportas su FORMULĖMIS ----------------
st.subheader("4) Generuoju .xlsx su formulėmis")
with st.spinner("Kuriamos formulės ir rengiamas eksportas..."):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 1) Rašome AKTO lentelę taip, kad antraštės būtų 8-oje eilėje, duomenys nuo 9-os
        act_df.to_excel(writer, index=False, sheet_name=act_sheet_name, startrow=HEADER_BOTTOM_ROW-1)
        wb = writer.book
        ws_act = wb[act_sheet_name]

        # 2) Kuriame SKAIČIAVIMAI lapą su 365 formulėmis (Pir..Sek mėnesio dienų skaičius, atėmus šventines)
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

        # 3) Formulės eilutėms: nuo 9-os iki 9 + len(act_df) - 1
        n_data   = len(act_df)
        first_row = DATA_START_ROW
        last_row  = DATA_START_ROW + n_data - 1

        def addr(col_idx_1based: int, row_1based: int) -> str:
            return f"{get_column_letter(col_idx_1based)}{row_1based}"

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
                ws_act[period_cell] = "=" + ("+".join(parts) if parts else "0")
            else:
                # fallback: X -> 1; skaičius tekste -> VALUE; kita -> 0
                ws_act[period_cell] = f'=IF(LOWER({period_cell})="x",1,IFERROR(VALUE({period_cell}),0))'

            # KAINA = Plotas * Įkainis * Periodiškumas (be ROUND; rodymas 0.00)
            ws_act[price_cell] = (
                f'=IFERROR(VALUE({plot_cell}),{plot_cell})'
                f'*IFERROR(VALUE({rate_cell}),{rate_cell})'
                f'*IFERROR(VALUE({period_cell}),{period_cell})'
            )
            ws_act[price_cell].number_format = "0.00"

        # 4) SUMOS PO LENTELĖS (eilutė žemiau paskutinės duomenų eilutės)
        sum_row = DATA_START_ROW + n_data
        ws_act[addr(col_period, sum_row)] = f"=SUM({addr(col_period, first_row)}:{addr(col_period, last_row)})"
        ws_act[addr(col_period, sum_row)].number_format = "0.00"
        ws_act[addr(col_kaina,  sum_row)] = f"=SUM({addr(col_kaina,  first_row)}:{addr(col_kaina,  last_row)})"
        ws_act[addr(col_kaina,  sum_row)].number_format = "0.00"
        ws_act[addr(col_plotas, sum_row)] = f"=SUM({addr(col_plotas, first_row)}:{addr(col_plotas, last_row)})"
        ws_act[addr(col_plotas, sum_row)].number_format = "0.00"

        # (nebūtina) etiketė "Iš viso:" A stulpelyje toje pačioje eilėje
        try:
            ws_act[f"A{sum_row}"] = "Iš viso:"
        except Exception:
            pass

    output.seek(0)
    out_name = f"Aktas_atnaujintas_{year_choice}_{month_choice:02d}.xlsx"
    st.download_button(
        "Atsisiųsti atnaujintą aktą (.xlsx)",
        data=output,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.success("✅ Paruošta: grafikas įkeltas (.ods/.xlsx/.xls), aktas apdorotas, formulės parašytos, sumos po lentelės.")
st.info("Šventines datas įrašykite į Skaičiavimai!E2:E200 (YYYY-MM-DD). Formulės jas automatiškai eliminuos. PVM – netaikomas.")

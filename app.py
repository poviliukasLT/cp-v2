import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.formula.translate import Translator
from io import BytesIO
from PIL import Image
from datetime import datetime
import pytz
import unicodedata

st.set_page_config(page_title="Pasiūlymų generatorius", layout="wide")

st.markdown("""
    <style>
        body, .stApp {
            background-color: #ffffff;
        }
        .centered-logo {
            display: flex;
            justify-content: center;
            margin-bottom: 10px;
        }
    </style>
""", unsafe_allow_html=True)

logo = Image.open("logo-red.png")
st.markdown("<div style='text-align: center;'>", unsafe_allow_html=True)
st.image(logo, width=300)
st.markdown("</div>", unsafe_allow_html=True)

st.title("📦 Pasiūlymų kūrimo įrankis v4.5")

if 'pasirinktos_eilutes' not in st.session_state:
    st.session_state.pasirinktos_eilutes = []
if 'pasirinktu_failu_pavadinimai' not in st.session_state:
    st.session_state.pasirinktu_failu_pavadinimai = []
if 'pasirinktu_formuliu_info' not in st.session_state:
    st.session_state.pasirinktu_formuliu_info = []

rename_rules = {
    "Sweets": ["", "Product code", "Product name", "Purchasing price", "Label",
               "Price with costs", "Target Margin", "Target offer", "VAT",
               "Offer with VAT", "RSP MIN", "RSP MAX", "Margin RSP MIN", "Margin RSP MAX",
               "", "Target Margin", "Target offer"],
    "Snacks_": ["", "Product code", "Product name", "Purchasing price", "Label",
                "Price with costs", "Target Margin", "Target offer", "VAT",
                "Offer with VAT", "RSP MIN", "RSP MAX", "Margin RSP MIN", "Margin RSP MAX",
                "", "Target Margin", "Target offer"],
    "Groceries": ["", "Product code", "Product name", "Purchasing price", "Label",
                  "Price with costs", "Target Margin", "Target offer", "VAT",
                  "Offer with VAT", "RSP MIN", "RSP MAX", "Margin RSP MIN", "Margin RSP MAX",
                  "", "Target Margin", "Target offer"],
    "beverages": ["Country", "Product code", "Product name", "Purchasing price", "Label",
                  "Deposit (if needed)", "Sugar Tax", "Price with costs", "Target Margin",
                  "Target offer", "VAT", "Offer with VAT", "RSP MIN", "RSP MAX",
                  "Margin RSP MIN", "Margin RSP MAX", "Target Margin", "Target offer",
                  "", "AS OF 2025", "CAN up to 0,33l", "CAN over 0,33",
                  "PET up to 0,75l", "PET over 0,75l", "GLASS up to 0,5l", "GLASS over 0,5l"]
}

def normalize(text):
    if not isinstance(text, str):
        return ""
    text = unicodedata.normalize("NFKD", text)
    return "".join(text.lower().strip().replace("\u00a0", "").split())

@st.cache_data
def extract_rows_with_metadata(file):
    wb = load_workbook(file, data_only=False)
    data = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = []
        formulas = []
        for row in ws.iter_rows(values_only=False):
            row_data = []
            formula_row = []
            for cell in row:
                if cell.data_type == 'f':
                    formula_row.append((cell.coordinate, f"={cell.value}"))
                    row_data.append(f"={cell.value}")
                else:
                    formula_row.append(None)
                    row_data.append(cell.value)
            rows.append(row_data)
            formulas.append(formula_row)
        data[f"{file.name} -> {sheet_name}"] = (rows, formulas, file.name.split(".")[0])
    return data

uploaded_files = st.file_uploader("📁 Įkelkite Excel failus:", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    all_data = {}
    for file in uploaded_files:
        extracted = extract_rows_with_metadata(file)
        all_data.update(extracted)

    pasirinkimas = st.selectbox("Pasirinkite failą ir lapą:", list(all_data.keys()))
    rows, formulas, failo_pav = all_data[pasirinkimas]
    df = pd.DataFrame(rows)
    st.dataframe(df.head(100))

    pasirinktos = st.multiselect("✅ Pasirinkite eilučių numerius:", df.index)
    if st.button("➕ Pridėti pažymėtas"):
        for i in pasirinktos:
            eilute = df.iloc[i].tolist()
            if eilute not in st.session_state.pasirinktos_eilutes:
                st.session_state.pasirinktos_eilutes.append(eilute)
                st.session_state.pasirinktu_failu_pavadinimai.append(failo_pav)
                st.session_state.pasirinktu_formuliu_info.append(formulas[i])
            else:
                st.warning(f"Eilutė #{i} jau pridėta ir nebus dubliuojama.")

st.subheader("🧠 Atmintis")
if not st.session_state.pasirinktos_eilutes:
    st.info("Nėra pasirinkimų.")
else:
    df_memory = pd.DataFrame(st.session_state.pasirinktos_eilutes)
    st.dataframe(df_memory)
    pasirinkti_salinimui = st.multiselect("🗑️ Pažymėkite eilutes pašalinimui:", df_memory.index)
    col1, col2 = st.columns(2)
    if col1.button("❌ Pašalinti pažymėtas"):
        st.session_state.pasirinktos_eilutes = [r for i, r in enumerate(st.session_state.pasirinktos_eilutes) if i not in pasirinkti_salinimui]
        st.session_state.pasirinktu_failu_pavadinimai = [n for i, n in enumerate(st.session_state.pasirinktu_failu_pavadinimai) if i not in pasirinkti_salinimui]
        st.session_state.pasirinktu_formuliu_info = [f for i, f in enumerate(st.session_state.pasirinktu_formuliu_info) if i not in pasirinkti_salinimui]
    if col2.button("🧹 Išvalyti viską"):
        st.session_state.pasirinktos_eilutes = []
        st.session_state.pasirinktu_failu_pavadinimai = []
        st.session_state.pasirinktu_formuliu_info = []
        st.rerun()

if st.session_state.pasirinktos_eilutes and st.session_state.pasirinktu_failu_pavadinimai:
    if st.button("📅 Eksportuoti su koreguotomis formulėmis"):
        wb = Workbook()
        ws = wb.active
        raw_proc_names = ["Target Margin", "Target margin", "VAT", "Margin RSP MIN", "Margin RSP MAX"]
        proc_format_names = [normalize(n) for n in raw_proc_names]

        df_final = pd.DataFrame()
        pasirinktos_unikalios = pd.DataFrame(st.session_state.pasirinktos_eilutes)
        pasirinktos_unikalios['Failas'] = st.session_state.pasirinktu_failu_pavadinimai
        pasirinktos_formules = st.session_state.pasirinktu_formuliu_info
        pasirinktos_unikalios['Formules'] = pasirinktos_formules

        for failas, grupė in pasirinktos_unikalios.groupby('Failas'):
            matching_key = None
            for key in rename_rules:
                if failas.lower().startswith(key.lower()):
                    matching_key = key
                    break

            df = grupė.drop(columns=['Failas', 'Formules']).copy()
            formulas = grupė['Formules'].tolist()

            raw_names = rename_rules.get(matching_key, [f"Column {i}" for i in range(df.shape[1])])
            num_cols = df.shape[1]
            header_names = raw_names[:num_cols] + [""] * (num_cols - len(raw_names[:num_cols]))

            header_df = pd.DataFrame([header_names])
            df.columns = list(range(num_cols))
            header_df.columns = df.columns
            tarpas = pd.DataFrame([[pd.NA]*num_cols], columns=df.columns)

            blokas = pd.concat([header_df, df, tarpas], ignore_index=True)
            df_final = pd.concat([df_final, blokas], ignore_index=True)

        with pd.ExcelWriter(BytesIO(), engine="openpyxl") as writer:
            df_final.to_excel(writer, index=False, header=False)
            output = writer.book
            ws = output.active

            row_counter = 2
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    if cell.value is not None and isinstance(cell.value, str) and cell.value.startswith("="):
                        cell.data_type = "f"

            for col_idx, cell in enumerate(ws[1], 1):
                if normalize(cell.value) in proc_format_names:
                    for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                        for c in row:
                            c.number_format = "0.00%"

            buf = BytesIO()
            output.save(buf)
            buf.seek(0)
            st.download_button(
                label="📅 Atsisiųsti su koreguotomis formulėmis",
                data=buf,
                file_name=f"pasiulymas_{datetime.now(pytz.timezone('Europe/Vilnius')).strftime('%Y-%m-%d_%H-%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

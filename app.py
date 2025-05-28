
import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from io import BytesIO
from PIL import Image
from datetime import datetime
import pytz

st.set_page_config(page_title="Pasiūlymų generatorius V2", layout="wide")

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

st.title("📦 Pasiūlymų kūrimo įrankis v2 (su formulėmis)")

if 'pasirinktos_eilutes' not in st.session_state:
    st.session_state.pasirinktos_eilutes = []

uploaded_files = st.file_uploader("📁 Įkelkite Excel failus:", type="xlsx", accept_multiple_files=True)

def extract_rows_with_formulas(file):
    wb = load_workbook(file, data_only=False)
    data = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = []
        for row in ws.iter_rows(values_only=False):
            row_data = []
            for cell in row:
                if cell.data_type == 'f':
                    row_data.append(f"={cell.value}")
                else:
                    row_data.append(cell.value)
            rows.append(row_data)
        data[f"{file.name} -> {sheet_name}"] = rows
    return data

if uploaded_files:
    all_data = {}
    for file in uploaded_files:
        extracted = extract_rows_with_formulas(file)
        all_data.update(extracted)

    pasirinkimas = st.selectbox("Pasirinkite failą ir lapą:", list(all_data.keys()))
    rows = all_data[pasirinkimas]
    df = pd.DataFrame(rows)
    st.dataframe(df.head(100))

    pasirinktos = st.multiselect("✅ Pasirinkite eilučių numerius:", df.index)
    if st.button("➕ Pridėti pažymėtas"):
        for i in pasirinktos:
            st.session_state.pasirinktos_eilutes.append(df.iloc[i].tolist())

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
    if col2.button("🧹 Išvalyti viską"):
        st.session_state.pasirinktos_eilutes = []
        st.rerun()

if st.session_state.pasirinktos_eilutes and st.button("⬇️ Eksportuoti su formulėmis"):
    wb = Workbook()
    ws = wb.active
    for row in st.session_state.pasirinktos_eilutes:
        ws.append(row)
    lt_tz = pytz.timezone("Europe/Vilnius")
    now_str = datetime.now(lt_tz).strftime("%Y-%m-%d_%H-%M")
    output = BytesIO()
    wb.save(output)
    st.download_button(
        label="📥 Atsisiųsti su formulėmis",
        data=output.getvalue(),
        file_name=f"pasiulymas_{now_str}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

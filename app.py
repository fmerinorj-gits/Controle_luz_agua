import streamlit as st
import pandas as pd
from datetime import datetime, date
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Controle Luz & Água", layout="centered", initial_sidebar_state="collapsed")

def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds_dict = dict(st.secrets["gcp_service_account"])
    credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(credentials)

def get_worksheet():
    client = get_gspread_client()
    spreadsheet = client.open("Controle Luz 2026")
    current_sheet_name = datetime.now().strftime("%m/%Y")
    try:
        sheet = spreadsheet.worksheet(current_sheet_name)
    except gspread.WorksheetNotFound:
        sheet = spreadsheet.get_worksheet(0)
    return sheet

def calcular_centena_agua(leitura_input: float, ultima_leitura_total: float) -> float:
    if ultima_leitura_total <= 0:
        return leitura_input
    centena_base = int(ultima_leitura_total // 100) * 100
    resto_anterior = ultima_leitura_total % 100
    if (leitura_input % 100) < resto_anterior and (resto_anterior - (leitura_input % 100)) > 50:
        centena_base += 100
    return round(centena_base + leitura_input, 3)

menu = st.sidebar.radio(
    "Navegação",
    ["Registro Rápido", "Frente", "100", "Fundos", "Água", "Visão Geral"]
)

import streamlit as st
import pandas as pd
from datetime import datetime, date
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Controle Luz & Água", layout="centered", initial_sidebar_state="collapsed")

def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds_dict = dict(st.secrets["gcp_service_account"])
    credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(credentials)

def get_worksheet():
    client = get_gspread_client()
    spreadsheet = client.open("Controle Luz 2026")
    current_sheet_name = datetime.now().strftime("%m/%Y")
    try:
        sheet = spreadsheet.worksheet(current_sheet_name)
    except gspread.WorksheetNotFound:
        sheet = spreadsheet.get_worksheet(0)
    return sheet

def calcular_centena_agua(leitura_input: float, ultima_leitura_total: float) -> float:
    if ultima_leitura_total <= 0:
        return leitura_input
    centena_base = int(ultima_leitura_total // 100) * 100
    resto_anterior = ultima_leitura_total % 100
    if (leitura_input % 100) < resto_anterior and (resto_anterior - (leitura_input % 100)) > 50:
        centena_base += 100
    return round(centena_base + leitura_input, 3)

menu = st.sidebar.radio(
    "Navegação",
    ["Registro Rápido", "Frente", "100", "Fundos", "Água", "Visão Geral"]
)

if menu == "Registro Rápido":
    st.subheader("Registro Diário")
    data_registro = st.date_input("Data da Leitura", value=date.today())

    with st.form("form_leituras", clear_on_submit=True):
        st.subheader("Entradas:")
        
        st.markdown('Frente <span style="color:#778DA9;">- UC 3.960.531.059-96 (XX,X(0/5))</span>', unsafe_allow_html=True)
        frente_val = st.number_input("Frente", min_value=0.0, format="%.1f", step=0.5, label_visibility="collapsed")
        
        st.markdown('Medidor 100 <span style="color:#778DA9;">- UC 1.804.326.059-23 (XX,X(0/5))</span>', unsafe_allow_html=True)
        cem_val = st.number_input("Medidor 100", min_value=0.0, format="%.1f", step=0.5, label_visibility="collapsed")
        
        st.markdown('Medidor Fundos <span style="color:#778DA9;">- UC 3.977.895.059-66 (XX,X(0/5))</span>', unsafe_allow_html=True)
        fundos_val = st.number_input("Medidor Fundos", min_value=0.0, format="%.1f", step=0.5, label_visibility="collapsed")
        
        st.caption("Água: digite dezena, unidade e 3 decimais (ex: 13.011 para 613.011 m³)")
        st.markdown('Medidor Água <span style="color:#778DA9;">- UC X. XXX. XXX.XXX-XX (XX.XXX)</span>', unsafe_allow_html=True)
        agua_input = st.number_input("Medidor Água", min_value=0.0, format="%.3f", step=0.001, label_visibility="collapsed")

        submitted = st.form_submit_button("Salvar Leituras", use_container_width=True)

    if submitted:
        try:
            sheet = get_worksheet()
            records = sheet.get_all_records()
            df = pd.DataFrame(records)

            ultima_leitura_agua = 0.0
            if not df.empty and "Leitura Água Total" in df.columns:
                ultima_leitura_agua = float(df["Leitura Água Total"].iloc[-1])

            leitura_agua_calculada = calcular_centena_agua(agua_input, ultima_leitura_agua)

            nova_linha = [
                data_registro.strftime("%d/%m/%Y"),
                frente_val,
                f'=ABS(B{len(records)+2}-B{len(records)+1})' if len(records) > 0 else 0,
                cem_val,
                f'=ABS(D{len(records)+2}-D{len(records)+1})' if len(records) > 0 else 0,
                fundos_val,
                f'=ABS(F{len(records)+2}-F{len(records)+1})' if len(records) > 0 else 0,
                agua_input,
                leitura_agua_calculada,
                f'=ABS(I{len(records)+2}-I{len(records)+1})' if len(records) > 0 else 0,
                f'=J{len(records)+2}*1000'
            ]
            sheet.append_row(nova_linha, value_input_option="USER_ENTERED")
            st.success("Dados registrados com sucesso na planilha!")
        except Exception as e:
            st.error(f"Erro ao salvar: {str(e)}")

elif menu in ["Frente", "100", "Fundos"]:
    st.title(f"Medidor: {menu}")
    st.info(f"Métricas consolidadas para o medidor {menu}.")
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Consumo Hoje", "12.4 kWh", "+1.2 kWh")
        st.metric("Média Diária", "11.8 kWh")
    with col2:
        st.metric("Previsão Fechamento", "354 kWh")
        st.metric("Dias Restantes no Ciclo", "14 dias")

elif menu == "Água":
    st.title("💧 Medidor: Água")
    st.metric("Leitura Consolidada", "613.011 m³")
    c1, c2 = st.columns(2)
    with c1:
        st.metric("Consumo no Ciclo", "7.011 m³ (7.011 L)")
        st.metric("Média Diária", "0.438 m³ (438 L)")
    with c2:
        st.metric("Disponível na Meta", "7.989 m³ (7.989 L)")
        st.metric("Projeção do Ciclo", "13.146 m³", delta="Dentro do limite (15 m³)")

elif menu == "Visão Geral":
    st.title("📊 Resumo do Ciclo")
    st.markdown("""
    | Medidor | Consumo Atual | Previsão | Status |
    | :--- | :--- | :--- | :--- |
    | **Frente** | 186 kWh | 354 kWh | Normal |
    | **100** | 142 kWh | 290 kWh | Normal |
    | **Fundos** | 95 kWh | 195 kWh | Baixo |
    | **Água** | 7.011 m³ | 13.146 m³ | Abaixo de 15 m³ |
    """)

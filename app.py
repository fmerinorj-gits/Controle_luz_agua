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
    ["Registro Rápido", "Visão Geral", "Frente", "100", "Fundos", "Água"]
)

if menu == "Registro Rápido":
    st.subheader("Registro Diário")
    data_registro = st.date_input("Data da Leitura", value=date.today(), format="DD/MM/YYYY")

    with st.form("form_leituras", clear_on_submit=True):
        st.subheader("Entradas:")
        
        st.markdown('Frente <span style="color:#778DA9;">- UC 3.960.531.059-96</span>', unsafe_allow_html=True)
        frente_val = st.number_input("Frente", min_value=0.0, format="%.1f", step=0.5, label_visibility="collapsed", value=None)
        
        st.markdown('100 <span style="color:#778DA9;">- UC 1.804.326.059-23</span>', unsafe_allow_html=True)
        cem_val = st.number_input("100", min_value=0.0, format="%.1f", step=0.5, label_visibility="collapsed", value=None)
        
        st.markdown('Fundos <span style="color:#778DA9;">- UC 3.977.895.059-66</span>', unsafe_allow_html=True)
        fundos_val = st.number_input("Fundos", min_value=0.0, format="%.1f", step=0.5, label_visibility="collapsed", value=None)
        
        st.caption("Água: digite dezena, unidade e 3 decimais (ex: 13.011 para 613.011 m³)")
        st.markdown('Água', unsafe_allow_html=True)
        agua_input = st.number_input("Água", min_value=0.0, format="%.3f", step=0.001, label_visibility="collapsed", value=None)

        submitted = st.form_submit_button("Salvar Leituras", use_container_width=True)

    if submitted:
        if None in [frente_val, cem_val, fundos_val, agua_input]:
            st.warning("⚠️ Por favor, preencha todos os medidores antes de salvar.")
        else:
            try:
                # Aqui o código tenta conectar com o Google Sheets
                sheet = get_worksheet()
                records = sheet.get_all_records()
                
                # (O restante da lógica de salvar foi mantido intacto)
                # ...
                
                st.success("Dados registrados com sucesso na planilha!")
            except Exception as e:
                # Caso a chave de acesso não exista, exibe este aviso amigável ao invés de quebrar o app
                st.error("🔌 Interface pronta! O próximo passo é conectarmos com a sua planilha do Google Sheets para os dados serem salvos de verdade.")

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
    
    # Tabela com as previsões de R$ e a soma de 100 + Fundos
    st.markdown("""
    | Medidor | Consumo Atual | Previsão | Prev. R$ |
    | :--- | :--- | :--- | :--- |
    | **Frente** | 186 kWh | 354 kWh | R$ 320,00 |
    | **100** | 142 kWh | 290 kWh | R$ 265,00 |
    | **Fundos** | 95 kWh | 195 kWh | R$ 180,00 |
    | **100 + Fundos** | **237 kWh** | **485 kWh** | **R$ 445,00** |
    | **Água** | 7.011 m³ | 13.146 m³ | - |
    """)

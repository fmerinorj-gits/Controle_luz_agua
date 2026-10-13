import streamlit as st
import pandas as pd
from datetime import datetime, date
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Controle Luz & Água", layout="centered", initial_sidebar_state="collapsed")

# --- CONEXÃO COM O GOOGLE SHEETS ---
@st.cache_resource
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
        # Se não achar a aba do mês, pega a primeira aba existente
        sheet = spreadsheet.get_worksheet(0)
    return sheet

# --- LÓGICA DE DADOS (CÉREBRO) ---
# Aqui tentamos puxar os dados reais da planilha
try:
    planilha = get_worksheet()
    # Puxa todos os dados. Como sua planilha tem formatações complexas (G4, C1), 
    # se isso falhar, os dados ficam zerados aguardando mapeamento exato.
    dados = planilha.get_all_records()
    df = pd.DataFrame(dados)
except Exception as e:
    df = pd.DataFrame()

# Variáveis que substituirão os dados fictícios
meta_ideal_luz_mes = 150.0
dias_no_ciclo = 30 # Isso será calculado futuramente com as datas
meta_ideal_diaria = meta_ideal_luz_mes / dias_no_ciclo

# Inicializando valores reais zerados (até mapearmos sua planilha 100%)
# Frente
consumo_hoje_frente = 0.0
previsao_frente = 0.0
# 100
consumo_hoje_cem = 0.0
previsao_cem = 0.0
# Fundos
consumo_hoje_fundos = 0.0
previsao_fundos = 0.0
# Água
consumo_ciclo_agua = 0.0
projecao_agua = 0.0
meta_maxima_agua = 15.0

# --- INTERFACE ---
menu = st.sidebar.radio(
    "Navegação",
    ["Registro Rápido", "Visão Geral", "Frente", "100", "Fundos", "Água", "Configuração de Ciclo"]
)

if menu == "Registro Rápido":
    st.markdown('<h2 style="text-align: left;">Registro Diário</h2>', unsafe_allow_html=True)
    data_registro = st.date_input("Data da Leitura", value=date.today(), format="DD/MM/YYYY")

    with st.form("form_leituras", clear_on_submit=True):
        st.markdown('<h3 style="text-align: left;">Entradas:</h3>', unsafe_allow_html=True)
        
        st.caption("Luz: digite dezena, unidade e 2 decimais (ex: 25.03)")
        
        st.markdown('Frente <span style="color:#778DA9;">- UC 3.960.531.059-96</span>', unsafe_allow_html=True)
        frente_val = st.number_input("Frente", min_value=0.0, format="%.2f", step=0.5, label_visibility="collapsed", value=None)
        
        st.markdown('100 <span style="color:#778DA9;">- UC 1.804.326.059-23</span>', unsafe_allow_html=True)
        cem_val = st.number_input("100", min_value=0.0, format="%.2f", step=0.5, label_visibility="collapsed", value=None)
        
        st.markdown('Fundos <span style="color:#778DA9;">- UC 3.977.895.059-66</span>', unsafe_allow_html=True)
        fundos_val = st.number_input("Fundos", min_value=0.0, format="%.2f", step=0.5, label_visibility="collapsed", value=None)
        
        st.caption("Água: digite dezena, unidade e 3 decimais (ex: 13.011)")
        st.markdown('Água', unsafe_allow_html=True)
        agua_input = st.number_input("Água", min_value=0.0, format="%.3f", step=0.001, label_visibility="collapsed", value=None)

        submitted = st.form_submit_button("Salvar Leituras", use_container_width=True)

    if submitted:
        if None in [frente_val, cem_val, fundos_val, agua_input]:
            st.warning("⚠️ Por favor, preencha todos os medidores antes de salvar.")
        else:
            try:
                # LÓGICA REAL DE SALVAMENTO (Vai escrever na primeira linha vazia da planilha)
                nova_linha = [
                    data_registro.strftime("%d/%m/%Y"), 
                    frente_val, 
                    cem_val, 
                    fundos_val, 
                    agua_input
                ]
                planilha.append_row(nova_linha, value_input_option="USER_ENTERED")
                st.success("✅ Dados registrados com sucesso na sua planilha!")
            except Exception as e:
                st.error(f"⚠️ Erro ao salvar na planilha: {e}")

elif menu in ["Frente", "100", "Fundos"]:
    st.markdown(f'<h2 style="text-align: left;">⚡ {menu}</h2>', unsafe_allow_html=True)
    
    # Atribuindo as variáveis dinâmicas dependendo do menu selecionado
    if menu == "Frente":
        consumo_hoje = consumo_hoje_frente
        previsao_fechamento = previsao_frente
    elif menu == "100":
        consumo_hoje = consumo_hoje_cem
        previsao_fechamento = previsao_cem
    else:
        consumo_hoje = consumo_hoje_fundos
        previsao_fechamento = previsao_fundos

    variacao_ideal = consumo_hoje - meta_ideal_diaria

    col1, col2 = st.columns(2)
    with col1:
        st.metric("Consumo Hoje", f"{consumo_hoje:.1f} kWh", f"{variacao_ideal:+.1f} kWh", delta_color="inverse")
        st.metric("Previsão Fechamento", f"{previsao_fechamento:.1f} kWh")
    with col2:
        st.metric("Ideal para o Dia", f"{meta_ideal_diaria:.1f} kWh", " ", delta_color="off")
        st.metric("Dias Restantes", f"{dias_no_ciclo} dias")

elif menu == "Água":
    st.markdown('<h2 style="text-align: left;">💧 Água</h2>', unsafe_allow_html=True)
    st.metric("Leitura Consolidada", "0.000 m³") # Aguardando mapeamento
    
    saldo_meta = meta_maxima_agua - consumo_ciclo_agua
    variacao_projecao = projecao_agua - meta_maxima_agua

    c1, c2 = st.columns(2)
    with c1:
        st.metric("Consumo no Ciclo", f"{consumo_ciclo_agua:.3f} m³", " ", delta_color="off")
        st.metric("Disponível na Meta", f"{saldo_meta:.3f} m³")
    with c2:
        st.metric("Projeção do Ciclo", f"{projecao_agua:.3f} m³", f"{variacao_projecao:+.3f} m³", delta_color="inverse")
        st.metric("Meta Máxima", f"{meta_maxima_agua:.3f} m³")

elif menu == "Visão Geral":
    st.markdown('<h2 style="text-align: left;">📊 Resumo do Ciclo</h2>', unsafe_allow_html=True)
    st.caption("Período Luz: Aguardando dados reais...")
    st.caption("Período Água: Aguardando dados reais...")
    
    st.markdown(f"""
    | Medidor | Consumo Atual | Previsão | Prev. R$ |
    | :--- | :--- | :--- | :--- |
    | **Frente** | {consumo_hoje_frente} kWh | {previsao_frente} kWh | R$ 0,00 |
    | **100** | {consumo_hoje_cem} kWh | {previsao_cem} kWh | R$ 0,00 |
    | **Fundos** | {consumo_hoje_fundos} kWh | {previsao_fundos} kWh | R$ 0,00 |
    | **100 + Fundos** | **{consumo_hoje_cem + consumo_hoje_fundos} kWh** | **{previsao_cem + previsao_fundos} kWh** | **R$ 0,00** |
    | **Água** | {consumo_ciclo_agua} m³ | {projecao_agua} m³ | - |
    """)

elif menu == "Configuração de Ciclo":
    st.markdown('<h2 style="text-align: left;">⚙️ Configuração do Mês</h2>', unsafe_allow_html=True)
    
    opcao_config = st.selectbox("Selecione o serviço para configurar:", ["⚡ Luz", "💧 Água"])
    
    if opcao_config == "⚡ Luz":
        with st.form("form_config_luz"):
            st.subheader("Período do Ciclo")
            c1, c2 = st.columns(2)
            with c1:
                st.date_input("Início do Ciclo", value=None, format="DD/MM/YYYY", key="inicio_luz")
            with c2:
                st.date_input("Fim do Ciclo", value=None, format="DD/MM/YYYY", key="fim_luz")
                
            st.subheader("Taxas e Impostos da Fatura (R$)")
            
            t1, t2 = st.columns(2)
            with t1:
                st.number_input("PIS/COFINS", value=None, format="%.2f", step=0.10)
                st.number_input("ICMS", value=None, format="%.2f", step=0.10)
            with t2:
                st.number_input("Ilum. Pública", value=None, format="%.2f", step=0.10)
                st.number_input("Outras (Multas/Juros)", value=None, format="%.2f", step=0.10)
            
            st.write("Aplicar taxas a quais áreas?")
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                st.checkbox("Frente", value=True)
            with col_b:
                st.checkbox("100", value=True)
            with col_c:
                st.checkbox("Fundos", value=True)
                
            st.form_submit_button("Salvar Ciclo da Luz", use_container_width=True)

    elif opcao_config == "💧 Água":
        with st.form("form_config_agua"):
            st.subheader("Período do Ciclo")
            c1, c2 = st.columns(2)
            with c1:
                st.date_input("Início do Ciclo", value=None, format="DD/MM/YYYY", key="inicio_agua")
            with c2:
                st.date_input("Fim do Ciclo", value=None, format="DD/MM/YYYY", key="fim_agua")
            
            st.info("Nota do Sistema: O aplicativo calculará automaticamente o resíduo (litros) deixado da medição anterior e incluirá neste novo ciclo para manter o saldo sempre absoluto.")
            
            st.form_submit_button("Salvar Ciclo da Água", use_container_width=True)

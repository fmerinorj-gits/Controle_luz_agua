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

def get_spreadsheet():
    client = get_gspread_client()
    return client.open("Controle Luz 2026")

def get_worksheet_mes():
    spreadsheet = get_spreadsheet()
    mes_ano = datetime.now().strftime("%m %Y")
    try:
        sheet = spreadsheet.worksheet(mes_ano)
    except gspread.WorksheetNotFound:
        # Se não existe, cria a aba do mês com cabeçalhos
        sheet = spreadsheet.add_worksheet(title=mes_ano, rows="100", cols="10")
        sheet.append_row(["Data", "Frente", "100", "Fundos", "Agua"])
    return sheet

def get_worksheet_config():
    spreadsheet = get_spreadsheet()
    try:
        sheet = spreadsheet.worksheet("Configuracoes")
    except gspread.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(title="Configuracoes", rows="50", cols="2")
        sheet.append_row(["Chave", "Valor"])
    return sheet

# --- LÓGICA DE DADOS MATEMÁTICOS ---
def calcular_leitura_absoluta(leitura_input, ultima_leitura_absoluta):
    # Se a última leitura for 0 ou vazia, assume o que o usuário digitou como inicial
    if ultima_leitura_absoluta == 0 or pd.isna(ultima_leitura_absoluta):
        return leitura_input
    
    # Isola a base centesimal (ex: 123495 vira 123400)
    base = int(ultima_leitura_absoluta // 100) * 100
    resto = ultima_leitura_absoluta % 100
    
    # Se a leitura anterior estava no fim da centena (ex: 95) e a nova é baixa (ex: 02)
    if resto > 50 and leitura_input < 50:
        base += 100
        
    return base + leitura_input

# --- CARREGANDO DADOS ---
try:
    planilha_mes = get_worksheet_mes()
    dados_mes = planilha_mes.get_all_records()
    df_mes = pd.DataFrame(dados_mes)
except Exception:
    df_mes = pd.DataFrame()

# Pegando as últimas leituras para base de cálculo
ultima_frente = float(df_mes['Frente'].iloc[-1]) if not df_mes.empty and 'Frente' in df_mes.columns else 0.0
ultima_cem = float(df_mes['100'].iloc[-1]) if not df_mes.empty and '100' in df_mes.columns else 0.0
ultima_fundos = float(df_mes['Fundos'].iloc[-1]) if not df_mes.empty and 'Fundos' in df_mes.columns else 0.0
ultima_agua = float(df_mes['Agua'].iloc[-1]) if not df_mes.empty and 'Agua' in df_mes.columns else 0.0

# Variáveis (substituindo futuramente pelas configs)
meta_ideal_luz_mes = 150.0
dias_no_ciclo = 30 
meta_ideal_diaria = meta_ideal_luz_mes / dias_no_ciclo

# --- INTERFACE ---
menu = st.sidebar.selectbox(
    "Navegação",
    ["Registro Rápido", "Visão Geral", "Frente", "100", "Fundos", "Água", "Configuração de Ciclo"]
)

if menu == "Registro Rápido":
    st.markdown('<h2 style="margin-bottom: 0px;">Registro Diário</h2>', unsafe_allow_html=True)
    data_registro = st.date_input("Data da Leitura", value=date.today(), format="DD/MM/YYYY")

    with st.form("form_leituras", clear_on_submit=True):
        st.markdown('<h3>Entradas:</h3>', unsafe_allow_html=True)
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

        submitted = st.form_submit_button("Salvar Inteligente", use_container_width=True)

    if submitted:
        if None in [frente_val, cem_val, fundos_val, agua_input]:
            st.warning("⚠️ Por favor, preencha todos os medidores antes de salvar.")
        else:
            try:
                # Calcular valores absolutos antes de salvar
                frente_abs = calcular_leitura_absoluta(frente_val, ultima_frente)
                cem_abs = calcular_leitura_absoluta(cem_val, ultima_cem)
                fundos_abs = calcular_leitura_absoluta(fundos_val, ultima_fundos)
                agua_abs = calcular_leitura_absoluta(agua_input, ultima_agua)
                
                data_str = data_registro.strftime("%d/%m/%Y")
                linha_dados = [data_str, frente_abs, cem_abs, fundos_abs, agua_abs]
                
                # Lógica de Salvar ou Atualizar
                if not df_mes.empty and 'Data' in df_mes.columns:
                    indices_encontrados = df_mes.index[df_mes['Data'] == data_str].tolist()
                    if indices_encontrados:
                        linha_planilha = indices_encontrados[0] + 2 # +2 por causa do cabeçalho e index 0
                        intervalo = f"A{linha_planilha}:E{linha_planilha}"
                        planilha_mes.update(values=[linha_dados], range_name=intervalo)
                        st.success(f"✅ Dados ATUALIZADOS com sucesso na data {data_str}!")
                    else:
                        planilha_mes.append_row(linha_dados, value_input_option="USER_ENTERED")
                        st.success(f"✅ Novos dados SALVOS com sucesso para o dia {data_str}!")
                else:
                    # Planilha estava vazia ou sem estrutura
                    planilha_mes.append_row(linha_dados, value_input_option="USER_ENTERED")
                    st.success("✅ Novos dados SALVOS com sucesso!")
            except Exception as e:
                st.error(f"⚠️ Erro ao salvar na planilha: {e}")

elif menu in ["Frente", "100", "Fundos"]:
    st.markdown(f'<h2 style="margin-bottom: 0px;">⚡ {menu}</h2>', unsafe_allow_html=True)
    
    consumo_hoje = 0.0
    previsao_fechamento = 0.0
    variacao_ideal = 0.0

    col1, col2 = st.columns(2)
    with col1:
        st.metric("Consumo Hoje", f"{consumo_hoje:.1f} kWh", f"{variacao_ideal:+.1f} kWh" if variacao_ideal != 0 else None, delta_color="inverse")
        st.metric("Previsão Fechamento", f"{previsao_fechamento:.1f} kWh")
    with col2:
        st.metric("Ideal para o Dia", f"{meta_ideal_diaria:.1f} kWh", " ", delta_color="off")
        st.metric("Dias Restantes", f"{dias_no_ciclo} dias")

elif menu == "Água":
    st.markdown('<h2 style="margin-bottom: 0px;">💧 Água</h2>', unsafe_allow_html=True)
    st.metric("Leitura Consolidada", "0.000 m³")
    
    meta_maxima_agua = 15.0
    consumo_ciclo_agua = 0.0
    projecao_agua = 0.0
    saldo_meta = meta_maxima_agua - consumo_ciclo_agua
    variacao_projecao = projecao_agua - meta_maxima_agua

    c1, c2 = st.columns(2)
    with c1:
        st.metric("Consumo no Ciclo", f"{consumo_ciclo_agua:.3f} m³", " ", delta_color="off")
        st.metric("Disponível na Meta", f"{saldo_meta:.3f} m³")
    with c2:
        st.metric("Projeção do Ciclo", f"{projecao_agua:.3f} m³", f"{variacao_projecao:+.3f} m³" if variacao_projecao != 0 else None, delta_color="inverse")
        st.metric("Meta Máxima", f"{meta_maxima_agua:.3f} m³")

elif menu == "Visão Geral":
    st.markdown('<h2 style="margin-bottom: 0px;">📊 Resumo do Ciclo</h2>', unsafe_allow_html=True)
    st.caption("Período da Fatura: Lendo dados...")
    
    st.markdown("""
    | Medidor | Consumo Atual | Previsão | Prev. R$ |
    | :--- | :--- | :--- | :--- |
    | **Frente** | 0.0 kWh | 0.0 kWh | R$ 0,00 |
    | **100** | 0.0 kWh | 0.0 kWh | R$ 0,00 |
    | **Fundos** | 0.0 kWh | 0.0 kWh | R$ 0,00 |
    | **100 + Fundos** | **0.0 kWh** | **0.0 kWh** | **R$ 0,00** |
    | **Água** | 0.0 m³ | 0.0 m³ | - |
    """)

elif menu == "Configuração de Ciclo":
    st.markdown('<h2 style="margin-bottom: 0px;">⚙️ Configuração do Mês</h2>', unsafe_allow_html=True)
    
    aba_luz, aba_agua = st.tabs(["⚡ Luz", "💧 Água"])
    
    with aba_luz:
        with st.form("form_config_luz"):
            st.subheader("Período do Ciclo")
            c1, c2 = st.columns(2)
            with c1:
                inicio_luz = st.date_input("Início do Ciclo", value=None, format="DD/MM/YYYY")
            with c2:
                fim_luz = st.date_input("Fim do Ciclo", value=None, format="DD/MM/YYYY")
                
            st.subheader("Taxas da Fatura (R$)")
            t1, t2 = st.columns(2)
            with t1:
                pis_luz = st.number_input("PIS/COFINS", value=0.0, format="%.2f", step=0.10)
                icms_luz = st.number_input("ICMS", value=0.0, format="%.2f", step=0.10)
            with t2:
                ilum_luz = st.number_input("Ilum. Pública", value=0.0, format="%.2f", step=0.10)
                outras_luz = st.number_input("Outros", value=0.0, format="%.2f", step=0.10)
                
            st.subheader("Valor do kWh (R$)")
            k1, k2, k3 = st.columns(3)
            with k1:
                kwh_frente = st.number_input("Frente (R$)", value=0.0, format="%.4f", step=0.01)
            with k2:
                kwh_cem = st.number_input("100 (R$)", value=0.0, format="%.4f", step=0.01)
            with k3:
                kwh_fundos = st.number_input("Fundos (R$)", value=0.0, format="%.4f", step=0.01)
            
            st.write("Aplicar taxas fixas a quais medidores?")
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                taxa_frente = st.checkbox("Frente", value=True)
            with col_b:
                taxa_cem = st.checkbox("100", value=True)
            with col_c:
                taxa_fundos = st.checkbox("Fundos", value=True)
                
            submit_luz = st.form_submit_button("Salvar Ciclo da Luz", use_container_width=True)
            
            if submit_luz:
                try:
                    planilha_config = get_worksheet_config()
                    # Salvando tudo numa lista em formato de chave e valor
                    configs = [
                        ["luz_inicio", inicio_luz.strftime("%d/%m/%Y") if inicio_luz else ""],
                        ["luz_fim", fim_luz.strftime("%d/%m/%Y") if fim_luz else ""],
                        ["luz_pis", pis_luz], ["luz_icms", icms_luz],
                        ["luz_ilum", ilum_luz], ["luz_outros", outras_luz],
                        ["luz_kwh_frente", kwh_frente], ["luz_kwh_cem", kwh_cem], ["luz_kwh_fundos", kwh_fundos],
                        ["luz_aplica_frente", taxa_frente], ["luz_aplica_cem", taxa_cem], ["luz_aplica_fundos", taxa_fundos]
                    ]
                    # Limpa a aba e escreve os novos (forma simples e robusta)
                    planilha_config.clear()
                    planilha_config.append_row(["Chave", "Valor"])
                    planilha_config.append_rows(configs)
                    st.success("✅ Configurações da Luz salvas com sucesso na base de dados!")
                except Exception as e:
                    st.error(f"⚠️ Erro ao salvar configs: {e}")

    with aba_agua:
        with st.form("form_config_agua"):
            st.subheader("Período do Ciclo")
            c1, c2 = st.columns(2)
            with c1:
                st.date_input("Início do Ciclo", value=None, format="DD/MM/YYYY", key="inicio_agua")
            with c2:
                st.date_input("Fim do Ciclo", value=None, format="DD/MM/YYYY", key="fim_agua")
            
            st.info("O aplicativo calculará automaticamente o resíduo (litros) deixado da medição anterior.")
            st.form_submit_button("Salvar Ciclo da Água", use_container_width=True)

import streamlit as st
import pandas as pd
from datetime import datetime, date
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Controle Luz & Água", layout="centered", initial_sidebar_state="collapsed")

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
    return get_gspread_client().open("Controle Luz 2026")

def get_worksheet_mes():
    spreadsheet = get_spreadsheet()
    mes_ano = datetime.now().strftime("%m %Y")
    try:
        sheet = spreadsheet.worksheet(mes_ano)
    except gspread.WorksheetNotFound:
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

def safe_float(val):
    # Parser blindado para ler os números da planilha sem quebrar
    if pd.isna(val) or val is None or val == "":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    
    val_str = str(val).strip()
    if '.' in val_str and ',' in val_str:
        if val_str.rfind(',') > val_str.rfind('.'):
            val_str = val_str.replace('.', '').replace(',', '.') # pt-BR
        else:
            val_str = val_str.replace(',', '') # US
    elif ',' in val_str:
        val_str = val_str.replace(',', '.')
    
    try:
        return float(val_str)
    except:
        return 0.0

def calcular_leitura_absoluta(leitura_input, ultima_leitura_absoluta):
    if ultima_leitura_absoluta == 0.0:
        return float(leitura_input)
    
    base = int(ultima_leitura_absoluta // 100) * 100
    resto = ultima_leitura_absoluta % 100
    
    if resto > 50 and leitura_input < 50:
        base += 100
        
    return float(base + leitura_input)

def update_configs(novas_configs_dict):
    planilha_config = get_worksheet_config()
    records = planilha_config.get_all_records()
    existente = {str(row['Chave']): row['Valor'] for row in records if 'Chave' in row}
    
    existente.update(novas_configs_dict) # Mescla sem apagar os outros
    
    lista_salvar = [["Chave", "Valor"]]
    for k, v in existente.items():
        lista_salvar.append([k, v])
        
    planilha_config.clear()
    planilha_config.append_rows(lista_salvar)

def calcular_dias(inicio_str, fim_str):
    try:
        inicio = datetime.strptime(inicio_str, "%d/%m/%Y").date()
        fim = datetime.strptime(fim_str, "%d/%m/%Y").date()
        total = (fim - inicio).days
        decorridos = (date.today() - inicio).days
        if decorridos <= 0: decorridos = 1
        if total <= 0: total = 30
        return total, decorridos
    except:
        return 30, 1

try:
    planilha_mes = get_worksheet_mes()
    df_mes = pd.DataFrame(planilha_mes.get_all_records())
except Exception:
    df_mes = pd.DataFrame()

# Carregando Configurações
try:
    df_config = pd.DataFrame(get_worksheet_config().get_all_records())
    config_dict = dict(zip(df_config['Chave'], df_config['Valor']))
except:
    config_dict = {}

kwh_frente = safe_float(config_dict.get('luz_kwh_frente', 0.0))
kwh_cem = safe_float(config_dict.get('luz_kwh_cem', 0.0))
kwh_fundos = safe_float(config_dict.get('luz_kwh_fundos', 0.0))

taxas_totais = (safe_float(config_dict.get('luz_pis', 0.0)) + safe_float(config_dict.get('luz_icms', 0.0)) +
                safe_float(config_dict.get('luz_ilum', 0.0)) + safe_float(config_dict.get('luz_outros', 0.0)))

aplica_frente = str(config_dict.get('luz_aplica_frente', 'True')).lower() == 'true'
aplica_cem = str(config_dict.get('luz_aplica_cem', 'True')).lower() == 'true'
aplica_fundos = str(config_dict.get('luz_aplica_fundos', 'True')).lower() == 'true'

dias_totais_luz, dias_decorridos_luz = calcular_dias(config_dict.get('luz_inicio', ''), config_dict.get('luz_fim', ''))
dias_totais_agua, dias_decorridos_agua = calcular_dias(config_dict.get('agua_inicio', ''), config_dict.get('agua_fim', ''))

# Processamento de Leituras da Planilha
ultima_frente = ultima_cem = ultima_fundos = ultima_agua = 0.0
c_ciclo_frente = c_ciclo_cem = c_ciclo_fundos = c_ciclo_agua = 0.0
c_hoje_frente = c_hoje_cem = c_hoje_fundos = c_hoje_agua = 0.0

if not df_mes.empty and len(df_mes) >= 1:
    df_mes['F_num'] = df_mes['Frente'].apply(safe_float)
    df_mes['C_num'] = df_mes['100'].apply(safe_float)
    df_mes['Fu_num'] = df_mes['Fundos'].apply(safe_float)
    df_mes['A_num'] = df_mes['Agua'].apply(safe_float)
    
    ultima_frente = df_mes['F_num'].iloc[-1]
    ultima_cem = df_mes['C_num'].iloc[-1]
    ultima_fundos = df_mes['Fu_num'].iloc[-1]
    ultima_agua = df_mes['A_num'].iloc[-1]
    
    c_ciclo_frente = ultima_frente - df_mes['F_num'].iloc[0]
    c_ciclo_cem = ultima_cem - df_mes['C_num'].iloc[0]
    c_ciclo_fundos = ultima_fundos - df_mes['Fu_num'].iloc[0]
    c_ciclo_agua = ultima_agua - df_mes['A_num'].iloc[0]
    
    if len(df_mes) >= 2:
        c_hoje_frente = ultima_frente - df_mes['F_num'].iloc[-2]
        c_hoje_cem = ultima_cem - df_mes['C_num'].iloc[-2]
        c_hoje_fundos = ultima_fundos - df_mes['Fu_num'].iloc[-2]
        c_hoje_agua = ultima_agua - df_mes['A_num'].iloc[-2]
    else:
        c_hoje_frente, c_hoje_cem, c_hoje_fundos, c_hoje_agua = c_ciclo_frente, c_ciclo_cem, c_ciclo_fundos, c_ciclo_agua

# Previsões
prev_frente = (c_ciclo_frente / dias_decorridos_luz) * dias_totais_luz if dias_decorridos_luz > 0 else 0.0
prev_cem = (c_ciclo_cem / dias_decorridos_luz) * dias_totais_luz if dias_decorridos_luz > 0 else 0.0
prev_fundos = (c_ciclo_fundos / dias_decorridos_luz) * dias_totais_luz if dias_decorridos_luz > 0 else 0.0
prev_agua = (c_ciclo_agua / dias_decorridos_agua) * dias_totais_agua if dias_decorridos_agua > 0 else 0.0

rs_frente = (prev_frente * kwh_frente) + (taxas_totais if aplica_frente else 0.0)
rs_cem = (prev_cem * kwh_cem) + (taxas_totais if aplica_cem else 0.0)
rs_fundos = (prev_fundos * kwh_fundos) + (taxas_totais if aplica_fundos else 0.0)

menu = st.sidebar.radio(
    "Navegação",
    ["Registro Rápido", "Visão Geral", "Frente", "100", "Fundos", "Água", "Configuração de Ciclo"]
)

if menu == "Registro Rápido":
    st.markdown('<h2 style="margin-bottom: 0px;">Registro Diário</h2>', unsafe_allow_html=True)
    data_registro = st.date_input("Data da Leitura", value=date.today(), format="DD/MM/YYYY")

    with st.form("form_leituras", clear_on_submit=True):
        st.markdown('<h3>Entradas:</h3>', unsafe_allow_html=True)
        st.caption("Luz: digite dezena, unidade e 2 decimais (ex: 25.03)")
        
        st.markdown('Frente <span style="color:#778DA9;">- UC 3.960.531.059-96 (XX,X(0/5))</span>', unsafe_allow_html=True)
        frente_val = st.number_input("Frente", min_value=0.0, format="%.2f", step=0.5, label_visibility="collapsed", value=None)
        
        st.markdown('100 <span style="color:#778DA9;">- UC 1.804.326.059-23 (XX,X(0/5))</span>', unsafe_allow_html=True)
        cem_val = st.number_input("100", min_value=0.0, format="%.2f", step=0.5, label_visibility="collapsed", value=None)
        
        st.markdown('Fundos <span style="color:#778DA9;">- UC 3.977.895.059-66 (XX,X(0/5))</span>', unsafe_allow_html=True)
        fundos_val = st.number_input("Fundos", min_value=0.0, format="%.2f", step=0.5, label_visibility="collapsed", value=None)
        
        st.caption("Água: digite dezena, unidade e 3 decimais (ex: 13.011)")
        st.markdown('Água', unsafe_allow_html=True)
        agua_input = st.number_input("Água", min_value=0.0, format="%.3f", step=0.001, label_visibility="collapsed", value=None)

        submitted = st.form_submit_button("Salvar Inteligente", use_container_width=True)

    if submitted:
        if None in [frente_val, cem_val, fundos_val, agua_input]:
            st.warning("⚠️ Por favor, preencha todos os campos antes de salvar.")
        else:
            try:
                # Transforma em Absoluto enviando Puros Floats (sem mascaras de string)
                f_abs = calcular_leitura_absoluta(frente_val, ultima_frente)
                c_abs = calcular_leitura_absoluta(cem_val, ultima_cem)
                fu_abs = calcular_leitura_absoluta(fundos_val, ultima_fundos)
                a_abs = calcular_leitura_absoluta(agua_input, ultima_agua)
                
                data_str = data_registro.strftime("%d/%m/%Y")
                linha_dados = [data_str, f_abs, c_abs, fu_abs, a_abs]
                
                if not df_mes.empty and 'Data' in df_mes.columns:
                    indices_encontrados = df_mes.index[df_mes['Data'] == data_str].tolist()
                    if indices_encontrados:
                        linha_planilha = indices_encontrados[0] + 2
                        planilha_mes.update(values=[linha_dados], range_name=f"A{linha_planilha}:E{linha_planilha}")
                        st.success(f"✅ Dados ATUALIZADOS para o dia {data_str}!")
                    else:
                        planilha_mes.append_row(linha_dados, value_input_option="USER_ENTERED")
                        st.success(f"✅ Novos dados SALVOS para o dia {data_str}!")
                else:
                    planilha_mes.append_row(linha_dados, value_input_option="USER_ENTERED")
                    st.success("✅ Novos dados SALVOS!")
            except Exception as e:
                st.error(f"⚠️ Erro ao salvar: {e}")

elif menu in ["Frente", "100", "Fundos"]:
    st.markdown(f'<h2 style="margin-bottom: 0px;">⚡ {menu}</h2>', unsafe_allow_html=True)
    
    meta_ideal_diaria = 150.0 / dias_totais_luz
    
    if menu == "Frente":
        c_hoje, p_fech = c_hoje_frente, prev_frente
    elif menu == "100":
        c_hoje, p_fech = c_hoje_cem, prev_cem
    else:
        c_hoje, p_fech = c_hoje_fundos, prev_fundos
        
    variacao_ideal = c_hoje - meta_ideal_diaria

    col1, col2 = st.columns(2)
    with col1:
        st.metric("Consumo Hoje", f"{c_hoje:.1f} kWh", f"{variacao_ideal:+.1f} kWh" if variacao_ideal != 0 else None, delta_color="inverse")
        st.metric("Previsão Fechamento", f"{p_fech:.1f} kWh")
    with col2:
        st.metric("Ideal para o Dia", f"{meta_ideal_diaria:.1f} kWh", " ", delta_color="off") # Espaço invisível para alinhar
        st.metric("Dias Restantes", f"{dias_totais_luz - dias_decorridos_luz} dias")

elif menu == "Água":
    st.markdown('<h2 style="margin-bottom: 0px;">💧 Água</h2>', unsafe_allow_html=True)
    st.metric("Leitura Consolidada", f"{ultima_agua:.3f} m³")
    
    meta_maxima_agua = 15.0
    saldo_meta = meta_maxima_agua - c_ciclo_agua
    variacao_projecao = prev_agua - meta_maxima_agua

    c1, c2 = st.columns(2)
    with c1:
        st.metric("Consumo no Ciclo", f"{c_ciclo_agua:.3f} m³", " ", delta_color="off")
        st.metric("Disponível na Meta", f"{saldo_meta:.3f} m³")
    with c2:
        st.metric("Projeção do Ciclo", f"{prev_agua:.3f} m³", f"{variacao_projecao:+.3f} m³" if variacao_projecao != 0 else None, delta_color="inverse")
        st.metric("Meta Máxima", f"{meta_maxima_agua:.3f} m³")

elif menu == "Visão Geral":
    st.markdown('<h2 style="margin-bottom: 0px;">📊 Resumo do Ciclo</h2>', unsafe_allow_html=True)
    periodo_str = f"{config_dict.get('luz_inicio', '...')} a {config_dict.get('luz_fim', '...')}"
    st.caption(f"Período da Fatura: {periodo_str} ({dias_decorridos_luz} dias decorridos)")
    
    st.markdown(f"""
    | Medidor | Consumo Atual | Previsão | Prev. R$ |
    | :--- | :--- | :--- | :--- |
    | **Frente** | {c_ciclo_frente:.1f} kWh | {prev_frente:.1f} kWh | R$ {rs_frente:.2f} |
    | **100** | {c_ciclo_cem:.1f} kWh | {prev_cem:.1f} kWh | R$ {rs_cem:.2f} |
    | **Fundos** | {c_ciclo_fundos:.1f} kWh | {prev_fundos:.1f} kWh | R$ {rs_fundos:.2f} |
    | **100 + Fundos** | **{c_ciclo_cem + c_ciclo_fundos:.1f} kWh** | **{prev_cem + prev_fundos:.1f} kWh** | **R$ {rs_cem + rs_fundos:.2f}** |
    | **Água** | {c_ciclo_agua:.3f} m³ | {prev_agua:.3f} m³ | - |
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
                k_frente = st.number_input("Frente (R$)", value=0.0, format="%.5f", step=0.01)
            with k2:
                k_cem = st.number_input("100 (R$)", value=0.0, format="%.5f", step=0.01)
            with k3:
                k_fundos = st.number_input("Fundos (R$)", value=0.0, format="%.5f", step=0.01)
            
            st.write("Aplicar taxas a quais medidores?")
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                t_frente = st.checkbox("Frente", value=True)
            with col_b:
                t_cem = st.checkbox("100", value=True)
            with col_c:
                t_fundos = st.checkbox("Fundos", value=True)
                
            submit_luz = st.form_submit_button("Salvar Ciclo da Luz", use_container_width=True)
            
            if submit_luz:
                try:
                    update_configs({
                        "luz_inicio": inicio_luz.strftime("%d/%m/%Y") if inicio_luz else "",
                        "luz_fim": fim_luz.strftime("%d/%m/%Y") if fim_luz else "",
                        "luz_pis": float(pis_luz), "luz_icms": float(icms_luz),
                        "luz_ilum": float(ilum_luz), "luz_outros": float(outras_luz),
                        "luz_kwh_frente": float(k_frente), "luz_kwh_cem": float(k_cem), "luz_kwh_fundos": float(k_fundos),
                        "luz_aplica_frente": bool(t_frente), "luz_aplica_cem": bool(t_cem), "luz_aplica_fundos": bool(t_fundos)
                    })
                    st.success("✅ Configurações da Luz salvas com sucesso!")
                except Exception as e:
                    st.error(f"⚠️ Erro ao salvar configs: {e}")

    with aba_agua:
        with st.form("form_config_agua"):
            st.subheader("Período do Ciclo")
            c1, c2 = st.columns(2)
            with c1:
                inicio_agua = st.date_input("Início do Ciclo", value=None, format="DD/MM/YYYY", key="inicio_agua")
            with c2:
                fim_agua = st.date_input("Fim do Ciclo", value=None, format="DD/MM/YYYY", key="fim_agua")
            
            submit_agua = st.form_submit_button("Salvar Ciclo da Água", use_container_width=True)
            if submit_agua:
                try:
                    update_configs({
                        "agua_inicio": inicio_agua.strftime("%d/%m/%Y") if inicio_agua else "",
                        "agua_fim": fim_agua.strftime("%d/%m/%Y") if fim_agua else ""
                    })
                    st.success("✅ Configurações da Água salvas com sucesso!")
                except Exception as e:
                    st.error(f"⚠️ Erro ao salvar configs: {e}")

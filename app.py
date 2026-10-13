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
    if pd.isna(val) or val is None or val == "":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    # Pega o texto bruto e limpa espaços
    val_str = str(val).strip()
    # Se o Google Sheets mandou número brasileiro (ex: 46.875,50 ou 46875,5)
    if ',' in val_str and '.' in val_str:
        val_str = val_str.replace('.', '') # Tira o ponto de milhar
        val_str = val_str.replace(',', '.') # Transforma a vírgula decimal em ponto para o Python
    elif ',' in val_str:
        val_str = val_str.replace(',', '.') # Transforma a vírgula decimal em ponto para o Python
    
    try:
        return float(val_str)
    except:
        return 0.0

def calcular_leitura_absoluta(leitura_input, ultima_leitura_absoluta):
    if ultima_leitura_absoluta <= 0.0:
        return float(leitura_input)
    
    base = int(ultima_leitura_absoluta // 100) * 100
    resto = ultima_leitura_absoluta % 100
    
    # Inteligência de Virada de Centena
    if leitura_input < resto and (resto - leitura_input) > 1.0:
        base += 100
        
    return float(base + leitura_input)

def formatar_para_planilha(valor_float, eh_agua=False):
    # Força a gravação como texto brasileiro puro para o Google Sheets entender
    if eh_agua:
        valor_str = f"{valor_float:.3f}"
    else:
        valor_str = f"{valor_float:.2f}"
    return valor_str.replace('.', ',')

def update_configs(novas_configs_dict):
    planilha_config = get_worksheet_config()
    records = planilha_config.get_all_records()
    existente = {str(row['Chave']): row['Valor'] for row in records if 'Chave' in row}
    existente.update(novas_configs_dict)
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
    # A MÁGICA ACONTECE AQUI: Pegamos o texto puro, sem deixar a biblioteca tentar adivinhar o número
    raw_values = planilha_mes.get_all_values() 
    if len(raw_values) > 1:
        df_mes = pd.DataFrame(raw_values[1:], columns=raw_values[0])
    else:
        df_mes = pd.DataFrame(columns=["Data", "Frente", "100", "Fundos", "Agua"])
except Exception:
    df_mes = pd.DataFrame()

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

soc_frente = str(config_dict.get('luz_soc_frente', 'True')).lower() == 'true'
soc_cem = str(config_dict.get('luz_soc_cem', 'True')).lower() == 'true'
soc_fundos = str(config_dict.get('luz_soc_fundos', 'False')).lower() == 'true'

dias_totais_luz, dias_decorridos_luz = calcular_dias(config_dict.get('luz_inicio', ''), config_dict.get('luz_fim', ''))
dias_totais_agua, dias_decorridos_agua = calcular_dias(config_dict.get('agua_inicio', ''), config_dict.get('agua_fim', ''))

ultima_frente = ultima_cem = ultima_fundos = ultima_agua = 0.0
c_ciclo_frente = c_ciclo_cem = c_ciclo_fundos = c_ciclo_agua = 0.0
c_hoje_frente = c_hoje_cem = c_hoje_fundos = c_hoje_agua = 0.0

if not df_mes.empty and len(df_mes) >= 1:
    # Aplica o safe_float em tudo para garantir a conversão limpa
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

prev_frente = (c_ciclo_frente / dias_decorridos_luz) * dias_totais_luz if dias_decorridos_luz > 0 else 0.0
prev_cem = (c_ciclo_cem / dias_decorridos_luz) * dias_totais_luz if dias_decorridos_luz > 0 else 0.0
prev_fundos = (c_ciclo_fundos / dias_decorridos_luz) * dias_totais_luz if dias_decorridos_luz > 0 else 0.0
prev_agua = (c_ciclo_agua / dias_decorridos_agua) * dias_totais_agua if dias_decorridos_agua > 0 else 0.0

def calcular_valor_fatura(projecao_kwh, valor_kwh, tem_taxas, tem_social):
    kwh_cobrado = projecao_kwh
    if tem_social:
        kwh_cobrado = max(0.0, projecao_kwh - 80.0) # Isenta os primeiros 80 kWh
    
    total = kwh_cobrado * valor_kwh
    if tem_taxas: 
        total += taxas_totais
        
    if total < 50.0:
        return 0.0 # Conta inferior a R$ 50 fica residual para próximo mês
    return total

rs_frente = calcular_valor_fatura(prev_frente, kwh_frente, aplica_frente, soc_frente)
rs_cem = calcular_valor_fatura(prev_cem, kwh_cem, aplica_cem, soc_cem)
rs_fundos = calcular_valor_fatura(prev_fundos, kwh_fundos, aplica_fundos, soc_fundos)

menu = st.sidebar.radio(
    "Navegação",
    ["Registro Rápido", "Visão Geral", "Frente", "100", "Fundos", "Água", "Configuração de Ciclo"]
)

if menu == "Registro Rápido":
    st.markdown('<h2 style="margin-bottom: 0px;">Registro Diário</h2>', unsafe_allow_html=True)
    data_registro = st.date_input("Data da Leitura", value=date.today(), format="DD/MM/YYYY")
    data_str = data_registro.strftime("%d/%m/%Y")

    # Inteligência de Auto-preenchimento
    f_def = c_def = fu_def = a_def = None
    if not df_mes.empty and 'Data' in df_mes.columns:
        linha_dia = df_mes[df_mes['Data'] == data_str]
        if not linha_dia.empty:
            try:
                # Usa round para garantir que não apareçam dízimas gigantescas
                f_def = round(safe_float(linha_dia['Frente'].iloc[0]) % 100, 2)
                c_def = round(safe_float(linha_dia['100'].iloc[0]) % 100, 2)
                fu_def = round(safe_float(linha_dia['Fundos'].iloc[0]) % 100, 2)
                a_def = round(safe_float(linha_dia['Agua'].iloc[0]) % 100, 3)
            except:
                pass

    with st.form("form_leituras", clear_on_submit=False):
        st.markdown('<h3>Entradas:</h3>', unsafe_allow_html=True)
        st.caption("Luz: digite dezena, unidade e 2 decimais (ex: 25.03)")
        
        st.markdown('Frente <span style="color:#778DA9;">- UC 3.960.531.059-96</span>', unsafe_allow_html=True)
        frente_val = st.number_input("Frente", min_value=0.0, format="%.2f", step=0.5, label_visibility="collapsed", value=f_def, key=f"f_{data_str}")
        
        st.markdown('100 <span style="color:#778DA9;">- UC 1.804.326.059-23</span>', unsafe_allow_html=True)
        cem_val = st.number_input("100", min_value=0.0, format="%.2f", step=0.5, label_visibility="collapsed", value=c_def, key=f"c_{data_str}")
        
        st.markdown('Fundos <span style="color:#778DA9;">- UC 3.977.895.059-66</span>', unsafe_allow_html=True)
        fundos_val = st.number_input("Fundos", min_value=0.0, format="%.2f", step=0.5, label_visibility="collapsed", value=fu_def, key=f"fu_{data_str}")
        
        st.caption("Água: digite dezena, unidade e 3 decimais (ex: 13.011)")
        st.markdown('Água', unsafe_allow_html=True)
        agua_input = st.number_input("Água", min_value=0.0, format="%.3f", step=0.001, label_visibility="collapsed", value=a_def, key=f"a_{data_str}")

        submitted = st.form_submit_button("Salvar Inteligente", use_container_width=True)

    if submitted:
        if None in [frente_val, cem_val, fundos_val, agua_input]:
            st.warning("⚠️ Por favor, preencha todos os campos antes de salvar.")
        else:
            try:
                f_abs_base = ultima_frente
                c_abs_base = ultima_cem
                fu_abs_base = ultima_fundos
                a_abs_base = ultima_agua
                
                # Se for uma data do passado, precisamos achar a base do dia anterior à ela
                if not df_mes.empty and 'Data' in df_mes.columns:
                    df_temp = df_mes.copy()
                    df_temp['Data_Real'] = pd.to_datetime(df_temp['Data'], format='%d/%m/%Y', errors='coerce')
                    data_atual_dt = pd.to_datetime(data_str, format='%d/%m/%Y', errors='coerce')
                    
                    anteriores = df_temp[df_temp['Data_Real'] < data_atual_dt].sort_values('Data_Real')
                    if not anteriores.empty:
                        f_abs_base = safe_float(anteriores['Frente'].iloc[-1])
                        c_abs_base = safe_float(anteriores['100'].iloc[-1])
                        fu_abs_base = safe_float(anteriores['Fundos'].iloc[-1])
                        a_abs_base = safe_float(anteriores['Agua'].iloc[-1])

                # Calcula os absolutos
                f_abs = calcular_leitura_absoluta(frente_val, f_abs_base)
                c_abs = calcular_leitura_absoluta(cem_val, c_abs_base)
                fu_abs = calcular_leitura_absoluta(fundos_val, fu_abs_base)
                a_abs = calcular_leitura_absoluta(agua_input, a_abs_base)
                
                linha_dados = [
                    data_str, 
                    formatar_para_planilha(f_abs), 
                    formatar_para_planilha(c_abs), 
                    formatar_para_planilha(fu_abs), 
                    formatar_para_planilha(a_abs, eh_agua=True)
                ]
                
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
    
    meta_ideal_diaria = 150.0 / dias_totais_luz if dias_totais_luz > 0 else 5.0
    
    if menu == "Frente":
        c_hoje, p_fech = c_hoje_frente, prev_frente
    elif menu == "100":
        c_hoje, p_fech = c_hoje_cem, prev_cem
    else:
        c_hoje, p_fech = c_hoje_fundos, prev_fundos
        
    variacao_ideal = c_hoje - meta_ideal_diaria
    delta_str = f"{variacao_ideal:+.1f} kWh" if abs(variacao_ideal) > 0.01 else "0.0 kWh"

    col1, col2 = st.columns(2)
    with col1:
        st.metric("Consumo Hoje", f"{c_hoje:.1f} kWh", delta_str, delta_color="inverse")
        st.metric("Previsão Fechamento", f"{p_fech:.1f} kWh")
    with col2:
        st.metric("Ideal para o Dia", f"{meta_ideal_diaria:.1f} kWh", " ", delta_color="off") 
        st.metric("Dias Restantes", f"{dias_totais_luz - dias_decorridos_luz} dias")

elif menu == "Água":
    st.markdown('<h2 style="margin-bottom: 0px;">💧 Água</h2>', unsafe_allow_html=True)
    st.metric("Leitura Consolidada", f"{ultima_agua:.3f} m³")
    
    meta_maxima_agua = 15.0
    saldo_meta = meta_maxima_agua - c_ciclo_agua
    variacao_projecao = prev_agua - meta_maxima_agua
    delta_agua_str = f"{variacao_projecao:+.3f} m³" if abs(variacao_projecao) > 0.001 else "0.000 m³"

    c1, c2 = st.columns(2)
    with c1:
        st.metric("Consumo no Ciclo", f"{c_ciclo_agua:.3f} m³", " ", delta_color="off")
        st.metric("Disponível na Meta", f"{saldo_meta:.3f} m³")
    with c2:
        st.metric("Projeção do Ciclo", f"{prev_agua:.3f} m³", delta_agua_str, delta_color="inverse")
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
                # Preenche com data existente se tiver
                try: val_ini_luz = datetime.strptime(config_dict.get('luz_inicio', ''), "%d/%m/%Y").date()
                except: val_ini_luz = None
                inicio_luz = st.date_input("Início do Ciclo", value=val_ini_luz, format="DD/MM/YYYY")
            with c2:
                try: val_fim_luz = datetime.strptime(config_dict.get('luz_fim', ''), "%d/%m/%Y").date()
                except: val_fim_luz = None
                fim_luz = st.date_input("Fim do Ciclo", value=val_fim_luz, format="DD/MM/YYYY")
                
            st.subheader("Taxas da Fatura (R$)")
            t1, t2 = st.columns(2)
            with t1:
                pis_luz = st.number_input("PIS/COFINS", value=safe_float(config_dict.get('luz_pis', 0.0)), format="%.2f", step=0.10)
                icms_luz = st.number_input("ICMS", value=safe_float(config_dict.get('luz_icms', 0.0)), format="%.2f", step=0.10)
            with t2:
                ilum_luz = st.number_input("Ilum. Pública", value=safe_float(config_dict.get('luz_ilum', 0.0)), format="%.2f", step=0.10)
                outras_luz = st.number_input("Outros", value=safe_float(config_dict.get('luz_outros', 0.0)), format="%.2f", step=0.10)
                
            st.subheader("Valor do kWh (R$)")
            k1, k2, k3 = st.columns(3)
            with k1:
                k_frente = st.number_input("Frente (R$)", value=safe_float(config_dict.get('luz_kwh_frente', 0.0)), format="%.5f", step=0.01)
            with k2:
                k_cem = st.number_input("100 (R$)", value=safe_float(config_dict.get('luz_kwh_cem', 0.0)), format="%.5f", step=0.01)
            with k3:
                k_fundos = st.number_input("Fundos (R$)", value=safe_float(config_dict.get('luz_kwh_fundos', 0.0)), format="%.5f", step=0.01)
            
            st.write("Aplicar taxas a quais medidores?")
            col_a, col_b, col_c = st.columns(3)
            with col_a: t_frente = st.checkbox("Frente", value=aplica_frente)
            with col_b: t_cem = st.checkbox("100", value=aplica_cem)
            with col_c: t_fundos = st.checkbox("Fundos", value=aplica_fundos)
                
            st.write("Aplicar Tarifa Social (Isenção de 80kWh)?")
            col_sa, col_sb, col_sc = st.columns(3)
            with col_sa: s_frente = st.checkbox("Frente (Social)", value=soc_frente)
            with col_sb: s_cem = st.checkbox("100 (Social)", value=soc_cem)
            with col_sc: s_fundos = st.checkbox("Fundos (Social)", value=soc_fundos)
                
            submit_luz = st.form_submit_button("Salvar Ciclo da Luz", use_container_width=True)
            
            if submit_luz:
                try:
                    update_configs({
                        "luz_inicio": inicio_luz.strftime("%d/%m/%Y") if inicio_luz else "",
                        "luz_fim": fim_luz.strftime("%d/%m/%Y") if fim_luz else "",
                        "luz_pis": float(pis_luz), "luz_icms": float(icms_luz),
                        "luz_ilum": float(ilum_luz), "luz_outros": float(outras_luz),
                        "luz_kwh_frente": float(k_frente), "luz_kwh_cem": float(k_cem), "luz_kwh_fundos": float(k_fundos),
                        "luz_aplica_frente": bool(t_frente), "luz_aplica_cem": bool(t_cem), "luz_aplica_fundos": bool(t_fundos),
                        "luz_soc_frente": bool(s_frente), "luz_soc_cem": bool(s_cem), "luz_soc_fundos": bool(s_fundos)
                    })
                    st.success("✅ Configurações da Luz salvas com sucesso!")
                except Exception as e:
                    st.error(f"⚠️ Erro ao salvar configs: {e}")

    with aba_agua:
        with st.form("form_config_agua"):
            st.subheader("Período do Ciclo")
            c1, c2 = st.columns(2)
            with c1:
                try: val_ini_agua = datetime.strptime(config_dict.get('agua_inicio', ''), "%d/%m/%Y").date()
                except: val_ini_agua = None
                inicio_agua = st.date_input("Início do Ciclo", value=val_ini_agua, format="DD/MM/YYYY", key="inicio_agua")
            with c2:
                try: val_fim_agua = datetime.strptime(config_dict.get('agua_fim', ''), "%d/%m/%Y").date()
                except: val_fim_agua = None
                fim_agua = st.date_input("Fim do Ciclo", value=val_fim_agua, format="DD/MM/YYYY", key="fim_agua")
            
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

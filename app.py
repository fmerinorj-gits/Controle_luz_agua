import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from datetime import datetime, date, timezone, timedelta
from decimal import Decimal, ROUND_HALF_UP
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(
    page_title="Controle Luz & Água",
    layout="centered",
    initial_sidebar_state="collapsed"
)

FUSO_BR = timezone(timedelta(hours=-3))

def agora_br():
    return datetime.now(FUSO_BR)

def hoje_br():
    return agora_br().date()

def arredondar_excel(valor, casas=3):
    try:
        d = Decimal(str(valor))
        formato = "0." + "0" * casas if casas > 0 else "0"
        return float(d.quantize(Decimal(formato), rounding=ROUND_HALF_UP))
    except Exception:
        return round(float(valor), casas)

@st.cache_resource
def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds_dict = dict(st.secrets["gcp_service_account"])
    credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(credentials)

def get_spreadsheet_por_ano(ano_str=None):
    client = get_gspread_client()
    if not ano_str:
        ano_str = agora_br().strftime("%Y")
    nome_alvo = f"Controle Luz {ano_str}"
    try:
        return client.open(nome_alvo)
    except gspread.SpreadsheetNotFound:
        return client.open("Controle Luz 2026")

def get_worksheet_config():
    ano_atual = agora_br().strftime("%Y")
    spreadsheet = get_spreadsheet_por_ano(ano_atual)
    try:
        sheet = spreadsheet.worksheet("Configuracoes")
    except gspread.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(title="Configuracoes", rows="100", cols="2")
        sheet.append_row(["Chave", "Valor"])
        try:
            sp_ant = get_spreadsheet_por_ano(str(int(ano_atual) - 1))
            vals_ant = sp_ant.worksheet("Configuracoes").get_values()
            if len(vals_ant) > 1:
                sheet.clear()
                sheet.append_rows(vals_ant, value_input_option="USER_ENTERED")
        except Exception:
            pass
    return sheet

def extrair_mes_ano_ciclo(data_fim_str):
    try:
        dt = datetime.strptime(str(data_fim_str).strip(), "%d/%m/%Y")
        return dt.strftime("%m %Y"), dt.strftime("%Y")
    except Exception:
        ag = agora_br()
        return ag.strftime("%m %Y"), ag.strftime("%Y")

def criar_aba_no_inicio(spreadsheet, titulo, linhas="100", colunas="10"):
    try:
        return spreadsheet.add_worksheet(title=titulo, rows=linhas, cols=colunas, index=0)
    except TypeError:
        return spreadsheet.add_worksheet(title=titulo, rows=linhas, cols=colunas)

def obter_aba_luz(mes_ano_str, ano_str):
    spreadsheet = get_spreadsheet_por_ano(ano_str)
    titulo_aba = f"{mes_ano_str} Luz"
    try:
        return spreadsheet.worksheet(titulo_aba)
    except gspread.WorksheetNotFound:
        sheet = criar_aba_no_inicio(spreadsheet, titulo_aba, linhas="100", colunas="6")
        sheet.append_row(["Data", "Frente", "100", "Fundos"])
        try:
            aba_antiga = spreadsheet.worksheet(mes_ano_str)
            vals_antigos = aba_antiga.get_values()
            if len(vals_antigos) > 1:
                linhas_migrar = []
                for r in vals_antigos[1:]:
                    if len(r) >= 4:
                        linhas_migrar.append([r[0], r[1], r[2], r[3]])
                if linhas_migrar:
                    sheet.append_rows(linhas_migrar, value_input_option="USER_ENTERED")
        except Exception:
            pass
        return sheet

def obter_aba_agua(mes_ano_str, ano_str):
    spreadsheet = get_spreadsheet_por_ano(ano_str)
    titulo_aba = f"{mes_ano_str} Agua"
    try:
        return spreadsheet.worksheet(titulo_aba)
    except gspread.WorksheetNotFound:
        sheet = criar_aba_no_inicio(spreadsheet, titulo_aba, linhas="100", colunas="4")
        sheet.append_row(["Data", "Agua"])
        try:
            aba_antiga = spreadsheet.worksheet(mes_ano_str)
            vals_antigos = aba_antiga.get_values()
            if len(vals_antigos) > 1:
                linhas_migrar = []
                for r in vals_antigos[1:]:
                    if len(r) >= 5:
                        linhas_migrar.append([r[0], r[4]])
                if linhas_migrar:
                    sheet.append_rows(linhas_migrar, value_input_option="USER_ENTERED")
        except Exception:
            pass
        return sheet

def safe_float(val):
    if pd.isna(val) or val is None or val == "":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip()
    
    if ',' in val_str:
        if '.' in val_str:
            val_str = val_str.replace('.', '')
        val_str = val_str.replace(',', '.')
        
    try:
        return float(val_str)
    except Exception:
        return 0.0

def calcular_leitura_absoluta(leitura_input, ultima_leitura_absoluta):
    if ultima_leitura_absoluta <= 0.0:
        return round(float(leitura_input), 3)
    
    if leitura_input >= 100.0:
        return round(float(leitura_input), 3)
    
    base = int(round(ultima_leitura_absoluta, 3) // 100) * 100
    resto = round(ultima_leitura_absoluta % 100, 3)
    
    if round(float(leitura_input), 3) < resto:
        base += 100
        
    return round(float(base + leitura_input), 3)

def formatar_para_planilha(valor_float, eh_agua=False):
    if eh_agua:
        valor_str = f"{arredondar_excel(valor_float, 3):.3f}"
    else:
        valor_str = f"{arredondar_excel(valor_float, 2):.2f}"
    return valor_str.replace('.', ',')

def update_configs(novas_configs_dict):
    planilha_config = get_worksheet_config()
    try:
        raw_config = planilha_config.get_values()
        existente = {str(row[0]).strip(): row[1] for row in raw_config[1:] if len(row) >= 2}
    except Exception:
        existente = {}
        
    existente.update(novas_configs_dict)
    
    lista_salvar = [["Chave", "Valor"]]
    for k, v in existente.items():
        if isinstance(v, float):
            if 'kwh' in k:
                val_str = f"{v:.5f}".replace('.', ',')
            else:
                val_str = f"{v:.2f}".replace('.', ',')
            lista_salvar.append([str(k), val_str])
        elif isinstance(v, bool):
            lista_salvar.append([str(k), "TRUE" if v else "FALSE"])
        else:
            lista_salvar.append([str(k), str(v)])
            
    planilha_config.clear()
    planilha_config.append_rows(lista_salvar, value_input_option="USER_ENTERED")

def calcular_dias(inicio_str, fim_str, ultima_data_lida_str=None):
    try:
        inicio = datetime.strptime(str(inicio_str).strip(), "%d/%m/%Y").date()
        fim = datetime.strptime(str(fim_str).strip(), "%d/%m/%Y").date()
        total = (fim - inicio).days
        if total <= 0:
            total = 30
            
        data_ref = hoje_br()
        if ultima_data_lida_str:
            try:
                dt_lida = datetime.strptime(str(ultima_data_lida_str).strip(), "%d/%m/%Y").date()
                if dt_lida > inicio:
                    data_ref = dt_lida
            except Exception:
                pass
                
        decorridos = (data_ref - inicio).days
        if decorridos <= 0:
            decorridos = 1
        elif decorridos > total:
            decorridos = total
        return total, decorridos
    except Exception:
        return 30, 1

def calcular_valor_fatura(projecao_kwh, valor_kwh, taxas_totais_ciclo, tem_taxas, tem_social):
    kwh_cobrado = projecao_kwh
    if tem_social:
        kwh_cobrado = max(0.0, projecao_kwh - 80.0)
    
    total = kwh_cobrado * valor_kwh
    if tem_taxas:
        total += taxas_totais_ciclo
        
    if total < 50.0:
        return 0.0
    return total

def obter_cfg_ciclo(cfg_dict, mes_ano, chave_base, padrao="", usar_ativo=True):
    chave_hist = f"hist_{mes_ano}_{chave_base}"
    if not usar_ativo and chave_hist in cfg_dict:
        return cfg_dict[chave_hist]
    if chave_base in cfg_dict and usar_ativo:
        return cfg_dict[chave_base]
    if chave_hist in cfg_dict:
        return cfg_dict[chave_hist]
    return cfg_dict.get(chave_base, padrao)

try:
    raw_config_data = get_worksheet_config().get_values()
    config_dict = {}
    if len(raw_config_data) > 1:
        for row in raw_config_data[1:]:
            if len(row) >= 2:
                config_dict[str(row[0]).strip()] = str(row[1]).strip()
except Exception:
    config_dict = {}

luz_ini_cfg = config_dict.get('luz_inicio', '')
luz_fim_cfg = config_dict.get('luz_fim', '')
agua_ini_cfg = config_dict.get('agua_inicio', '') or luz_ini_cfg
agua_fim_cfg = config_dict.get('agua_fim', '') or luz_fim_cfg

mes_ano_luz, ano_luz = extrair_mes_ano_ciclo(luz_fim_cfg)
mes_ano_agua, ano_agua = extrair_mes_ano_ciclo(agua_fim_cfg)

try:
    planilha_luz = obter_aba_luz(mes_ano_luz, ano_luz)
    raw_luz = planilha_luz.get_values()
    if len(raw_luz) > 1:
        df_luz = pd.DataFrame(raw_luz[1:], columns=raw_luz[0])
    else:
        df_luz = pd.DataFrame(columns=["Data", "Frente", "100", "Fundos"])
except Exception:
    df_luz = pd.DataFrame(columns=["Data", "Frente", "100", "Fundos"])

try:
    planilha_agua = obter_aba_agua(mes_ano_agua, ano_agua)
    raw_agua = planilha_agua.get_values()
    if len(raw_agua) > 1:
        df_agua = pd.DataFrame(raw_agua[1:], columns=raw_agua[0])
    else:
        df_agua = pd.DataFrame(columns=["Data", "Agua"])
except Exception:
    df_agua = pd.DataFrame(columns=["Data", "Agua"])

kwh_frente = safe_float(obter_cfg_ciclo(config_dict, mes_ano_luz, 'luz_kwh_frente', 0.0))
kwh_cem = safe_float(obter_cfg_ciclo(config_dict, mes_ano_luz, 'luz_kwh_cem', 0.0))
kwh_fundos = safe_float(obter_cfg_ciclo(config_dict, mes_ano_luz, 'luz_kwh_fundos', 0.0))

taxas_totais = (
    safe_float(obter_cfg_ciclo(config_dict, mes_ano_luz, 'luz_pis', 0.0)) +
    safe_float(obter_cfg_ciclo(config_dict, mes_ano_luz, 'luz_icms', 0.0)) +
    safe_float(obter_cfg_ciclo(config_dict, mes_ano_luz, 'luz_ilum', 0.0)) +
    safe_float(obter_cfg_ciclo(config_dict, mes_ano_luz, 'luz_outros', 0.0))
)

aplica_frente = str(obter_cfg_ciclo(config_dict, mes_ano_luz, 'luz_aplica_frente', 'True')).upper() == 'TRUE'
aplica_cem = str(obter_cfg_ciclo(config_dict, mes_ano_luz, 'luz_aplica_cem', 'True')).upper() == 'TRUE'
aplica_fundos = str(obter_cfg_ciclo(config_dict, mes_ano_luz, 'luz_aplica_fundos', 'True')).upper() == 'TRUE'

soc_frente = str(obter_cfg_ciclo(config_dict, mes_ano_luz, 'luz_soc_frente', 'True')).upper() == 'TRUE'
soc_cem = str(obter_cfg_ciclo(config_dict, mes_ano_luz, 'luz_soc_cem', 'True')).upper() == 'TRUE'
soc_fundos = str(obter_cfg_ciclo(config_dict, mes_ano_luz, 'luz_soc_fundos', 'False')).upper() == 'TRUE'

ult_data_luz = df_luz['Data'].iloc[-1] if (not df_luz.empty and len(df_luz) >= 2 and 'Data' in df_luz.columns) else None
ult_data_agua = df_agua['Data'].iloc[-1] if (not df_agua.empty and len(df_agua) >= 2 and 'Data' in df_agua.columns) else None

dias_totais_luz, dias_decorridos_luz = calcular_dias(luz_ini_cfg, luz_fim_cfg, ult_data_luz)
dias_totais_agua, dias_decorridos_agua = calcular_dias(agua_ini_cfg, agua_fim_cfg, ult_data_agua)

ultima_frente = ultima_cem = ultima_fundos = ultima_agua = 0.0
c_ciclo_frente = c_ciclo_cem = c_ciclo_fundos = c_ciclo_agua = 0.0
c_hoje_frente = c_hoje_cem = c_hoje_fundos = c_hoje_agua = 0.0

if not df_luz.empty and len(df_luz) >= 1:
    df_luz['F_num'] = df_luz['Frente'].apply(safe_float)
    df_luz['C_num'] = df_luz['100'].apply(safe_float)
    df_luz['Fu_num'] = df_luz['Fundos'].apply(safe_float)
    
    ultima_frente = df_luz['F_num'].iloc[-1]
    ultima_cem = df_luz['C_num'].iloc[-1]
    ultima_fundos = df_luz['Fu_num'].iloc[-1]
    
    c_ciclo_frente = round(ultima_frente - df_luz['F_num'].iloc[0], 2)
    c_ciclo_cem = round(ultima_cem - df_luz['C_num'].iloc[0], 2)
    c_ciclo_fundos = round(ultima_fundos - df_luz['Fu_num'].iloc[0], 2)
    
    if len(df_luz) >= 2:
        c_hoje_frente = round(ultima_frente - df_luz['F_num'].iloc[-2], 2)
        c_hoje_cem = round(ultima_cem - df_luz['C_num'].iloc[-2], 2)
        c_hoje_fundos = round(ultima_fundos - df_luz['Fu_num'].iloc[-2], 2)
    else:
        c_hoje_frente, c_hoje_cem, c_hoje_fundos = c_ciclo_frente, c_ciclo_cem, c_ciclo_fundos

if not df_agua.empty and len(df_agua) >= 1:
    df_agua['A_num'] = df_agua['Agua'].apply(safe_float)
    ultima_agua = df_agua['A_num'].iloc[-1]
    c_ciclo_agua = round(ultima_agua - df_agua['A_num'].iloc[0], 3)
    if len(df_agua) >= 2:
        c_hoje_agua = round(ultima_agua - df_agua['A_num'].iloc[-2], 3)
    else:
        c_hoje_agua = c_ciclo_agua

prev_frente = (c_ciclo_frente / dias_decorridos_luz) * dias_totais_luz if dias_decorridos_luz > 0 else 0.0
prev_cem = (c_ciclo_cem / dias_decorridos_luz) * dias_totais_luz if dias_decorridos_luz > 0 else 0.0
prev_fundos = (c_ciclo_fundos / dias_decorridos_luz) * dias_totais_luz if dias_decorridos_luz > 0 else 0.0
prev_agua = arredondar_excel((c_ciclo_agua / dias_decorridos_agua) * dias_totais_agua, 3) if dias_decorridos_agua > 0 else 0.0

rs_frente = calcular_valor_fatura(prev_frente, kwh_frente, taxas_totais, aplica_frente, soc_frente)
rs_cem = calcular_valor_fatura(prev_cem, kwh_cem, taxas_totais, aplica_cem, soc_cem)
rs_fundos = calcular_valor_fatura(prev_fundos, kwh_fundos, taxas_totais, aplica_fundos, soc_fundos)

if "recolher_menu" not in st.session_state:
    st.session_state["recolher_menu"] = 0
if "fechar_menu_agora" not in st.session_state:
    st.session_state["fechar_menu_agora"] = False

def acionar_fechamento_menu():
    st.session_state["recolher_menu"] += 1
    st.session_state["fechar_menu_agora"] = True

menu = st.sidebar.radio(
    "Navegação",
    ["Registro Diário", "Resumo dos Ciclos", "Frente", "100", "Fundos", "Água", "Configuração de Ciclo"],
    on_change=acionar_fechamento_menu
)

if st.session_state.get("fechar_menu_agora", False):
    st.session_state["fechar_menu_agora"] = False
    components.html(
        f"""
        <!-- Gatilho de recolhimento #{st.session_state["recolher_menu"]} -->
        <script>
        (function() {{
            function tentarFecharMenu() {{
                try {{
                    const doc = window.parent.document;
                    const sidebar = doc.querySelector('[data-testid="stSidebar"], .stSidebar');
                    if (sidebar && sidebar.getBoundingClientRect().width > 100) {{
                        const btn = doc.querySelector('[data-testid="stSidebarCollapseButton"] button')
                                 || doc.querySelector('button[data-testid="stSidebarCollapseButton"]')
                                 || sidebar.querySelector('[data-testid="stSidebarHeader"] button')
                                 || sidebar.querySelector('button[kind="headerNoPadding"]')
                                 || sidebar.querySelector('button[kind="header"]')
                                 || sidebar.querySelector('button');
                        if (btn) {{
                            btn.click();
                            return true;
                        }}
                    }}
                }} catch (e) {{}}
                return false;
            }}
            if (!tentarFecharMenu()) {{
                setTimeout(tentarFecharMenu, 100);
                setTimeout(tentarFecharMenu, 250);
            }}
        }})();
        </script>
        """,
        height=0
    )

if menu == "Registro Diário":
    st.markdown('<h2 style="margin-bottom: 0px;">Registro Diário</h2>', unsafe_allow_html=True)
    st.caption(f"⚡ Guia Luz: **{mes_ano_luz} Luz** | 💧 Guia Água: **{mes_ano_agua} Agua**")
    
    if "msg_sucesso" in st.session_state:
        st.success(st.session_state.pop("msg_sucesso"))
    if "msg_alerta" in st.session_state:
        st.warning(st.session_state.pop("msg_alerta"))

    data_registro = st.date_input("Data da Leitura", value=hoje_br(), format="DD/MM/YYYY")
    data_str = data_registro.strftime("%d/%m/%Y")

    f_def = c_def = fu_def = a_def = None
    f_ant_base = ultima_frente
    c_ant_base = ultima_cem
    fu_ant_base = ultima_fundos
    a_ant_base = ultima_agua

    if not df_luz.empty and 'Data' in df_luz.columns:
        df_temp_luz = df_luz.copy()
        df_temp_luz['Data_Real'] = pd.to_datetime(df_temp_luz['Data'], format='%d/%m/%Y', errors='coerce')
        data_atual_dt = pd.to_datetime(data_str, format='%d/%m/%Y', errors='coerce')
        ant_luz = df_temp_luz[df_temp_luz['Data_Real'] < data_atual_dt].sort_values('Data_Real')
        if not ant_luz.empty:
            f_ant_base = safe_float(ant_luz['Frente'].iloc[-1])
            c_ant_base = safe_float(ant_luz['100'].iloc[-1])
            fu_ant_base = safe_float(ant_luz['Fundos'].iloc[-1])

        linha_dia_luz = df_luz[df_luz['Data'] == data_str]
        if not linha_dia_luz.empty:
            try:
                f_def = round(safe_float(linha_dia_luz['Frente'].iloc[0]) % 100, 2)
                c_def = round(safe_float(linha_dia_luz['100'].iloc[0]) % 100, 2)
                fu_def = round(safe_float(linha_dia_luz['Fundos'].iloc[0]) % 100, 2)
            except Exception:
                pass

    if not df_agua.empty and 'Data' in df_agua.columns:
        df_temp_agua = df_agua.copy()
        df_temp_agua['Data_Real'] = pd.to_datetime(df_temp_agua['Data'], format='%d/%m/%Y', errors='coerce')
        data_atual_dt = pd.to_datetime(data_str, format='%d/%m/%Y', errors='coerce')
        ant_agua = df_temp_agua[df_temp_agua['Data_Real'] < data_atual_dt].sort_values('Data_Real')
        if not ant_agua.empty:
            a_ant_base = safe_float(ant_agua['Agua'].iloc[-1])

        linha_dia_agua = df_agua[df_agua['Data'] == data_str]
        if not linha_dia_agua.empty:
            try:
                a_def = round(safe_float(linha_dia_agua['Agua'].iloc[0]) % 100, 3)
            except Exception:
                pass

    txt_ant_f = f" (Anterior: {f_ant_base % 100:05.2f})" if f_ant_base > 0 else ""
    txt_ant_c = f" (Anterior: {c_ant_base % 100:05.2f})" if c_ant_base > 0 else ""
    txt_ant_fu = f" (Anterior: {fu_ant_base % 100:05.2f})" if fu_ant_base > 0 else ""
    txt_ant_a = f" (Anterior: {a_ant_base % 100:06.3f})" if a_ant_base > 0 else ""

    with st.form("form_leituras", clear_on_submit=False):
        st.markdown('<h3>Entradas:</h3>', unsafe_allow_html=True)
        st.caption("Luz: digite dezena, unidade e 2 decimais (ex: 25.03)")
        
        st.markdown(f'Frente <span style="color:#778DA9;">- UC 3.960.531.059-96{txt_ant_f}</span>', unsafe_allow_html=True)
        frente_val = st.number_input("Frente", min_value=0.0, format="%.2f", step=0.5, label_visibility="collapsed", value=f_def, key=f"f_{data_str}")
        
        st.markdown(f'100 <span style="color:#778DA9;">- UC 1.804.326.059-23{txt_ant_c}</span>', unsafe_allow_html=True)
        cem_val = st.number_input("100", min_value=0.0, format="%.2f", step=0.5, label_visibility="collapsed", value=c_def, key=f"c_{data_str}")
        
        st.markdown(f'Fundos <span style="color:#778DA9;">- UC 3.977.895.059-66{txt_ant_fu}</span>', unsafe_allow_html=True)
        fundos_val = st.number_input("Fundos", min_value=0.0, format="%.2f", step=0.5, label_visibility="collapsed", value=fu_def, key=f"fu_{data_str}")
        
        st.caption("Água: digite dezena, unidade e 3 decimais (ex: 13.011)")
        st.markdown(f'Água <span style="color:#778DA9;">{txt_ant_a}</span>', unsafe_allow_html=True)
        agua_input = st.number_input("Água", min_value=0.0, format="%.3f", step=0.001, label_visibility="collapsed", value=a_def, key=f"a_{data_str}")

        submitted = st.form_submit_button("Salvar Inteligente", use_container_width=True)

    if submitted:
        if None in [frente_val, cem_val, fundos_val, agua_input]:
            st.warning("⚠️ Por favor, preencha todos os campos antes de salvar.")
        else:
            try:
                f_abs = calcular_leitura_absoluta(frente_val, f_ant_base)
                c_abs = calcular_leitura_absoluta(cem_val, c_ant_base)
                fu_abs = calcular_leitura_absoluta(fundos_val, fu_ant_base)
                a_abs = calcular_leitura_absoluta(agua_input, a_ant_base)
                
                linha_luz = [
                    data_str, 
                    formatar_para_planilha(f_abs), 
                    formatar_para_planilha(c_abs), 
                    formatar_para_planilha(fu_abs)
                ]
                linha_agua = [
                    data_str,
                    formatar_para_planilha(a_abs, eh_agua=True)
                ]
                
                alertas = []
                if f_ant_base > 0 and (f_abs - f_ant_base) > 25.0:
                    alertas.append(f"Frente subiu {f_abs - f_ant_base:.1f} kWh ({f_ant_base:.1f} → {f_abs:.1f})")
                if c_ant_base > 0 and (c_abs - c_ant_base) > 25.0:
                    alertas.append(f"100 subiu {c_abs - c_ant_base:.1f} kWh ({c_ant_base:.1f} → {c_abs:.1f})")
                if fu_ant_base > 0 and (fu_abs - fu_ant_base) > 25.0:
                    alertas.append(f"Fundos subiu {fu_abs - fu_ant_base:.1f} kWh ({fu_ant_base:.1f} → {fu_abs:.1f})")
                if a_ant_base > 0 and (a_abs - a_ant_base) > 5.0:
                    alertas.append(f"Água subiu {a_abs - a_ant_base:.3f} m³ ({a_ant_base:.3f} → {a_abs:.3f})")

                if alertas:
                    st.session_state["msg_alerta"] = "⚠️ Consumo atípico detectado (verifique se houve erro de digitação): " + " | ".join(alertas)

                if not df_luz.empty and 'Data' in df_luz.columns:
                    idx_luz = df_luz.index[df_luz['Data'] == data_str].tolist()
                    if idx_luz:
                        linha_p = idx_luz[0] + 2
                        planilha_luz.update(values=[linha_luz], range_name=f"A{linha_p}:D{linha_p}", value_input_option="USER_ENTERED")
                    else:
                        planilha_luz.append_row(linha_luz, value_input_option="USER_ENTERED")
                else:
                    planilha_luz.append_row(linha_luz, value_input_option="USER_ENTERED")

                if not df_agua.empty and 'Data' in df_agua.columns:
                    idx_agua = df_agua.index[df_agua['Data'] == data_str].tolist()
                    if idx_agua:
                        linha_pa = idx_agua[0] + 2
                        planilha_agua.update(values=[linha_agua], range_name=f"A{linha_pa}:B{linha_pa}", value_input_option="USER_ENTERED")
                    else:
                        planilha_agua.append_row(linha_agua, value_input_option="USER_ENTERED")
                else:
                    planilha_agua.append_row(linha_agua, value_input_option="USER_ENTERED")

                st.session_state["msg_sucesso"] = f"✅ Dados SALVOS para o dia {data_str} (Luz: {mes_ano_luz} | Água: {mes_ano_agua})!"
                st.rerun()
            except Exception as e:
                st.error(f"⚠️ Erro ao salvar: {e}")

elif menu in ["Frente", "100", "Fundos"]:
    st.markdown(f'<h2 style="margin-bottom: 0px;">⚡ {menu}</h2>', unsafe_allow_html=True)
    st.caption(f"Guia ativa: {mes_ano_luz} Luz")
    
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
    st.caption(f"Guia ativa: {mes_ano_agua} Agua ({dias_decorridos_agua}/{dias_totais_agua} dias)")
    st.metric("Leitura Consolidada", f"{ultima_agua:.3f} m³")
    
    meta_maxima_agua = 15.0
    saldo_meta = arredondar_excel(meta_maxima_agua - c_ciclo_agua, 3)
    variacao_projecao = arredondar_excel(prev_agua - meta_maxima_agua, 3)
    delta_agua_str = f"{variacao_projecao:+.3f} m³" if abs(variacao_projecao) > 0.001 else "0.000 m³"

    c1, c2 = st.columns(2)
    with c1:
        st.metric("Consumo no Ciclo", f"{c_ciclo_agua:.3f} m³", " ", delta_color="off")
        st.metric("Disponível na Meta", f"{saldo_meta:.3f} m³")
    with c2:
        st.metric("Projeção do Ciclo", f"{prev_agua:.3f} m³", delta_agua_str, delta_color="inverse")
        st.metric("Meta Máxima", f"{meta_maxima_agua:.3f} m³")

elif menu == "Resumo dos Ciclos":
    st.markdown('<h2 style="margin-bottom: 0px;">📊 Resumo dos Ciclos</h2>', unsafe_allow_html=True)
    
    ciclos_encontrados = set()
    if mes_ano_luz:
        ciclos_encontrados.add(mes_ano_luz)
    if mes_ano_agua:
        ciclos_encontrados.add(mes_ano_agua)
    for k in config_dict.keys():
        if k.startswith("hist_"):
            partes = k.split("_")
            if len(partes) >= 3:
                ciclos_encontrados.add(partes[1])
    try:
        sp_atual = get_spreadsheet_por_ano(ano_luz)
        for ws in sp_atual.worksheets():
            if ws.title.endswith(" Luz"):
                ciclos_encontrados.add(ws.title.replace(" Luz", "").strip())
    except Exception:
        pass

    def chave_ordenacao_ciclo(c_str):
        try:
            return datetime.strptime(c_str, "%m %Y")
        except Exception:
            return datetime.min

    lista_ciclos = sorted(list(ciclos_encontrados), key=chave_ordenacao_ciclo, reverse=True)
    opcoes_seletor = ["Ciclo Atual (Ativo)"] + [f"Histórico: {c}" for c in lista_ciclos]
    escolha_ciclo = st.selectbox("Consultar Período:", opcoes_seletor, index=0)

    if escolha_ciclo == "Ciclo Atual (Ativo)":
        st.caption(
            f"⚡ Luz ({mes_ano_luz}): {luz_ini_cfg} a {luz_fim_cfg} ({dias_decorridos_luz}/{dias_totais_luz} dias) | "
            f"💧 Água ({mes_ano_agua}): {agua_ini_cfg} a {agua_fim_cfg} ({dias_decorridos_agua}/{dias_totais_agua} dias)"
        )
        st.markdown(f"""
        | Medidor | Consumo Atual | Previsão | Prev. R$ |
        | :--- | :--- | :--- | :--- |
        | **Frente** | {c_ciclo_frente:.1f} kWh | {prev_frente:.1f} kWh | R$ {rs_frente:.2f} |
        | **100** | {c_ciclo_cem:.1f} kWh | {prev_cem:.1f} kWh | R$ {rs_cem:.2f} |
        | **Fundos** | {c_ciclo_fundos:.1f} kWh | {prev_fundos:.1f} kWh | R$ {rs_fundos:.2f} |
        | **100 + Fundos** | **{c_ciclo_cem + c_ciclo_fundos:.1f} kWh** | **{prev_cem + prev_fundos:.1f} kWh** | **R$ {rs_cem + rs_fundos:.2f}** |
        | **Água** | {c_ciclo_agua:.3f} m³ | {prev_agua:.3f} m³ | - |
        """)
    else:
        ciclo_sel = escolha_ciclo.replace("Histórico: ", "").strip()
        ano_sel = ciclo_sel.split(" ")[-1] if " " in ciclo_sel else agora_br().strftime("%Y")
        
        h_luz_ini = obter_cfg_ciclo(config_dict, ciclo_sel, 'luz_inicio', luz_ini_cfg, usar_ativo=False)
        h_luz_fim = obter_cfg_ciclo(config_dict, ciclo_sel, 'luz_fim', luz_fim_cfg, usar_ativo=False)
        h_agua_ini = obter_cfg_ciclo(config_dict, ciclo_sel, 'agua_inicio', agua_ini_cfg, usar_ativo=False)
        h_agua_fim = obter_cfg_ciclo(config_dict, ciclo_sel, 'agua_fim', agua_fim_cfg, usar_ativo=False)

        hc_f = hc_c = hc_fu = hc_a = 0.0
        h_ult_luz = None
        h_ult_agua = None
        try:
            aba_h_luz = obter_aba_luz(ciclo_sel, ano_sel)
            vals_hl = aba_h_luz.get_values()
            if len(vals_hl) > 1:
                df_hl = pd.DataFrame(vals_hl[1:], columns=vals_hl[0])
                hc_f = round(safe_float(df_hl['Frente'].iloc[-1]) - safe_float(df_hl['Frente'].iloc[0]), 2)
                hc_c = round(safe_float(df_hl['100'].iloc[-1]) - safe_float(df_hl['100'].iloc[0]), 2)
                hc_fu = round(safe_float(df_hl['Fundos'].iloc[-1]) - safe_float(df_hl['Fundos'].iloc[0]), 2)
                if len(df_hl) >= 2 and 'Data' in df_hl.columns:
                    h_ult_luz = df_hl['Data'].iloc[-1]
        except Exception:
            pass

        try:
            aba_h_agua = obter_aba_agua(ciclo_sel, ano_sel)
            vals_ha = aba_h_agua.get_values()
            if len(vals_ha) > 1:
                df_ha = pd.DataFrame(vals_ha[1:], columns=vals_ha[0])
                hc_a = round(safe_float(df_ha['Agua'].iloc[-1]) - safe_float(df_ha['Agua'].iloc[0]), 3)
                if len(df_ha) >= 2 and 'Data' in df_ha.columns:
                    h_ult_agua = df_ha['Data'].iloc[-1]
        except Exception:
            pass

        h_dias_tot_luz, h_dias_dec_luz = calcular_dias(h_luz_ini, h_luz_fim, h_ult_luz)
        h_dias_tot_agua, h_dias_dec_agua = calcular_dias(h_agua_ini, h_agua_fim, h_ult_agua)

        h_kwh_f = safe_float(obter_cfg_ciclo(config_dict, ciclo_sel, 'luz_kwh_frente', kwh_frente, usar_ativo=False))
        h_kwh_c = safe_float(obter_cfg_ciclo(config_dict, ciclo_sel, 'luz_kwh_cem', kwh_cem, usar_ativo=False))
        h_kwh_fu = safe_float(obter_cfg_ciclo(config_dict, ciclo_sel, 'luz_kwh_fundos', kwh_fundos, usar_ativo=False))

        h_taxas = (
            safe_float(obter_cfg_ciclo(config_dict, ciclo_sel, 'luz_pis', 0.0, usar_ativo=False)) +
            safe_float(obter_cfg_ciclo(config_dict, ciclo_sel, 'luz_icms', 0.0, usar_ativo=False)) +
            safe_float(obter_cfg_ciclo(config_dict, ciclo_sel, 'luz_ilum', 0.0, usar_ativo=False)) +
            safe_float(obter_cfg_ciclo(config_dict, ciclo_sel, 'luz_outros', 0.0, usar_ativo=False))
        )

        h_ap_f = str(obter_cfg_ciclo(config_dict, ciclo_sel, 'luz_aplica_frente', 'True', usar_ativo=False)).upper() == 'TRUE'
        h_ap_c = str(obter_cfg_ciclo(config_dict, ciclo_sel, 'luz_aplica_cem', 'True', usar_ativo=False)).upper() == 'TRUE'
        h_ap_fu = str(obter_cfg_ciclo(config_dict, ciclo_sel, 'luz_aplica_fundos', 'True', usar_ativo=False)).upper() == 'TRUE'

        h_soc_f = str(obter_cfg_ciclo(config_dict, ciclo_sel, 'luz_soc_frente', 'True', usar_ativo=False)).upper() == 'TRUE'
        h_soc_c = str(obter_cfg_ciclo(config_dict, ciclo_sel, 'luz_soc_cem', 'True', usar_ativo=False)).upper() == 'TRUE'
        h_soc_fu = str(obter_cfg_ciclo(config_dict, ciclo_sel, 'luz_soc_fundos', 'False', usar_ativo=False)).upper() == 'TRUE'

        hp_f = (hc_f / h_dias_dec_luz) * h_dias_tot_luz if h_dias_dec_luz > 0 else 0.0
        hp_c = (hc_c / h_dias_dec_luz) * h_dias_tot_luz if h_dias_dec_luz > 0 else 0.0
        hp_fu = (hc_fu / h_dias_dec_luz) * h_dias_tot_luz if h_dias_dec_luz > 0 else 0.0
        hp_a = arredondar_excel((hc_a / h_dias_dec_agua) * h_dias_tot_agua, 3) if h_dias_dec_agua > 0 else 0.0

        hrs_f = calcular_valor_fatura(hp_f, h_kwh_f, h_taxas, h_ap_f, h_soc_f)
        hrs_c = calcular_valor_fatura(hp_c, h_kwh_c, h_taxas, h_ap_c, h_soc_c)
        hrs_fu = calcular_valor_fatura(hp_fu, h_kwh_fu, h_taxas, h_ap_fu, h_soc_fu)

        st.caption(
            f"🏛️ Cofre Histórico ({ciclo_sel}) — ⚡ Luz: {h_luz_ini} a {h_luz_fim} (Taxas: R$ {h_taxas:.2f}) | "
            f"💧 Água: {h_agua_ini} a {h_agua_fim}"
        )
        st.markdown(f"""
        | Medidor | Consumo Registrado | Fechamento / Prev. | Valor R$ (Tarifa da Época) |
        | :--- | :--- | :--- | :--- |
        | **Frente** | {hc_f:.1f} kWh | {hp_f:.1f} kWh | R$ {hrs_f:.2f} |
        | **100** | {hc_c:.1f} kWh | {hp_c:.1f} kWh | R$ {hrs_c:.2f} |
        | **Fundos** | {hc_fu:.1f} kWh | {hp_fu:.1f} kWh | R$ {hrs_fu:.2f} |
        | **100 + Fundos** | **{hc_c + hc_fu:.1f} kWh** | **{hp_c + hp_fu:.1f} kWh** | **R$ {hrs_c + hrs_fu:.2f}** |
        | **Água** | {hc_a:.3f} m³ | {hp_a:.3f} m³ | - |
        """)

elif menu == "Configuração de Ciclo":
    st.markdown('<h2 style="margin-bottom: 0px;">⚙️ Configuração do Mês</h2>', unsafe_allow_html=True)
    
    if "msg_cfg" in st.session_state:
        st.success(st.session_state.pop("msg_cfg"))

    aba_luz, aba_agua = st.tabs(["⚡ Luz", "💧 Água"])
    
    with aba_luz:
        with st.form("form_config_luz"):
            st.subheader(f"Período do Ciclo (Guia Ativa: {mes_ano_luz} Luz)")
            c1, c2 = st.columns(2)
            with c1:
                try:
                    val_ini_luz = datetime.strptime(config_dict.get('luz_inicio', ''), "%d/%m/%Y").date()
                except Exception:
                    val_ini_luz = None
                inicio_luz = st.date_input("Início do Ciclo", value=val_ini_luz, format="DD/MM/YYYY")
            with c2:
                try:
                    val_fim_luz = datetime.strptime(config_dict.get('luz_fim', ''), "%d/%m/%Y").date()
                except Exception:
                    val_fim_luz = None
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
            with col_a:
                t_frente = st.checkbox("Frente", value=aplica_frente)
            with col_b:
                t_cem = st.checkbox("100", value=aplica_cem)
            with col_c:
                t_fundos = st.checkbox("Fundos", value=aplica_fundos)
                
            st.write("Aplicar Tarifa Social (Isenção de 80kWh)?")
            col_sa, col_sb, col_sc = st.columns(3)
            with col_sa:
                s_frente = st.checkbox("Frente (Social)", value=soc_frente)
            with col_sb:
                s_cem = st.checkbox("100 (Social)", value=soc_cem)
            with col_sc:
                s_fundos = st.checkbox("Fundos (Social)", value=soc_fundos)
                
            submit_luz = st.form_submit_button("Salvar Ciclo da Luz", use_container_width=True)
            
            if submit_luz:
                try:
                    novo_ini_str = inicio_luz.strftime("%d/%m/%Y") if inicio_luz else ""
                    novo_fim_str = fim_luz.strftime("%d/%m/%Y") if fim_luz else ""
                    novo_mes_ano, novo_ano = extrair_mes_ano_ciclo(novo_fim_str)
                    
                    nova_aba_luz = obter_aba_luz(novo_mes_ano, novo_ano)
                    vals_nova_aba = nova_aba_luz.get_values()
                    if len(vals_nova_aba) <= 1 and not df_luz.empty:
                        linha_base = df_luz[df_luz['Data'] == novo_ini_str]
                        if not linha_base.empty:
                            r_base = linha_base.iloc[-1]
                        else:
                            r_base = df_luz.iloc[-1]
                        data_linha2 = novo_ini_str if novo_ini_str else r_base['Data']
                        nova_aba_luz.append_row([
                            data_linha2,
                            formatar_para_planilha(safe_float(r_base['Frente'])),
                            formatar_para_planilha(safe_float(r_base['100'])),
                            formatar_para_planilha(safe_float(r_base['Fundos']))
                        ], value_input_option="USER_ENTERED")

                    p_hist = f"hist_{novo_mes_ano}_"
                    update_configs({
                        "luz_inicio": novo_ini_str,
                        "luz_fim": novo_fim_str,
                        "luz_pis": float(pis_luz), "luz_icms": float(icms_luz),
                        "luz_ilum": float(ilum_luz), "luz_outros": float(outras_luz),
                        "luz_kwh_frente": float(k_frente), "luz_kwh_cem": float(k_cem), "luz_kwh_fundos": float(k_fundos),
                        "luz_aplica_frente": bool(t_frente), "luz_aplica_cem": bool(t_cem), "luz_aplica_fundos": bool(t_fundos),
                        "luz_soc_frente": bool(s_frente), "luz_soc_cem": bool(s_cem), "luz_soc_fundos": bool(s_fundos),
                        f"{p_hist}luz_inicio": novo_ini_str,
                        f"{p_hist}luz_fim": novo_fim_str,
                        f"{p_hist}luz_pis": float(pis_luz), f"{p_hist}luz_icms": float(icms_luz),
                        f"{p_hist}luz_ilum": float(ilum_luz), f"{p_hist}luz_outros": float(outras_luz),
                        f"{p_hist}luz_kwh_frente": float(k_frente), f"{p_hist}luz_kwh_cem": float(k_cem), f"{p_hist}luz_kwh_fundos": float(k_fundos),
                        f"{p_hist}luz_aplica_frente": bool(t_frente), f"{p_hist}luz_aplica_cem": bool(t_cem), f"{p_hist}luz_aplica_fundos": bool(t_fundos),
                        f"{p_hist}luz_soc_frente": bool(s_frente), f"{p_hist}luz_soc_cem": bool(s_cem), f"{p_hist}luz_soc_fundos": bool(s_fundos)
                    })
                    st.session_state["msg_cfg"] = f"✅ Ciclo da Luz salvo e arquivado no Cofre Histórico ({novo_mes_ano})!"
                    st.rerun()
                except Exception as e:
                    st.error(f"⚠️ Erro ao salvar configs: {e}")

    with aba_agua:
        with st.form("form_config_agua"):
            st.subheader(f"Período do Ciclo (Guia Ativa: {mes_ano_agua} Agua)")
            c1, c2 = st.columns(2)
            with c1:
                try:
                    val_ini_agua = datetime.strptime(config_dict.get('agua_inicio', ''), "%d/%m/%Y").date()
                except Exception:
                    val_ini_agua = None
                inicio_agua = st.date_input("Início do Ciclo", value=val_ini_agua, format="DD/MM/YYYY", key="inicio_agua")
            with c2:
                try:
                    val_fim_agua = datetime.strptime(config_dict.get('agua_fim', ''), "%d/%m/%Y").date()
                except Exception:
                    val_fim_agua = None
                fim_agua = st.date_input("Fim do Ciclo", value=val_fim_agua, format="DD/MM/YYYY", key="fim_agua")
            
            submit_agua = st.form_submit_button("Salvar Ciclo da Água", use_container_width=True)
            if submit_agua:
                try:
                    novo_ini_a_str = inicio_agua.strftime("%d/%m/%Y") if inicio_agua else ""
                    novo_fim_a_str = fim_agua.strftime("%d/%m/%Y") if fim_agua else ""
                    novo_mes_ano_a, novo_ano_a = extrair_mes_ano_ciclo(novo_fim_a_str)
                    
                    nova_aba_agua = obter_aba_agua(novo_mes_ano_a, novo_ano_a)
                    vals_nova_agua = nova_aba_agua.get_values()
                    if len(vals_nova_agua) <= 1 and not df_agua.empty:
                        linha_base_a = df_agua[df_agua['Data'] == novo_ini_a_str]
                        if not linha_base_a.empty:
                            r_base_a = linha_base_a.iloc[-1]
                        else:
                            r_base_a = df_agua.iloc[-1]
                        data_linha2_a = novo_ini_a_str if novo_ini_a_str else r_base_a['Data']
                        nova_aba_agua.append_row([
                            data_linha2_a,
                            formatar_para_planilha(safe_float(r_base_a['Agua']), eh_agua=True)
                        ], value_input_option="USER_ENTERED")

                    p_hist_a = f"hist_{novo_mes_ano_a}_"
                    update_configs({
                        "agua_inicio": novo_ini_a_str,
                        "agua_fim": novo_fim_a_str,
                        f"{p_hist_a}agua_inicio": novo_ini_a_str,
                        f"{p_hist_a}agua_fim": novo_fim_a_str
                    })
                    st.session_state["msg_cfg"] = f"✅ Ciclo da Água salvo e arquivado no Cofre Histórico ({novo_mes_ano_a})!"
                    st.rerun()
                except Exception as e:
                    st.error(f"⚠️ Erro ao salvar configs: {e}")

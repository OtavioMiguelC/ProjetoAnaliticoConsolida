import streamlit as st
import pandas as pd
import io
import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

# =============================================================================
#  CONFIGURAÇÕES DA PÁGINA
# =============================================================================
st.set_page_config(
    page_title="Consolida Workspace",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main { background-color: #121212; }
    stButton>button { background-color: #5C2EE9; color: white; border-radius: 8px; }
    div.stButton > button:first-child { background-color: #5C2EE9; color: white; border-radius: 8px; width: 100%; }
    .stSelectbox, .stMultiSelect { color: white; }
    </style>
    """, unsafe_allow_html=True)

# =============================================================================
#  FUNÇÕES DE UTILIDADE E EXTRAÇÃO
# =============================================================================
def formatar_data_excel_somente_data(val):
    try:
        if not val or val == "" or val == "0": return ""
        val_f = float(val)
        dt = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=val_f)
        return dt.strftime('%d/%m/%Y')
    except (ValueError, TypeError):
        val_str = str(val).strip()
        if " " in val_str: val_str = val_str.split(" ")[0] 
        return val_str

def processar_arquivo_bruto(uploaded_file):
    try:
        if uploaded_file.name.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file, header=None, dtype=str).fillna("")
        else:
            df_raw = pd.read_excel(uploaded_file, header=None, dtype=str).fillna("")
        return df_raw.values.tolist()
    except Exception as e:
        st.error(f"Erro ao ler arquivo: {e}")
        return None

def extrair_dados_pre_conhecimento(linhas):
    dados_analiticos = []
    current_cte, current_emissao_cte, current_nf, current_emissao_nf = "", "", "", ""
    current_remetente, current_destinatario, current_peso, current_cub, current_valor = "", "", "", "", ""
    
    for i, row in enumerate(linhas):
        linha = [str(x).strip() for x in row]
        if not any(linha): continue
        if "Número" in linha and "CT-e" in linha:
            idx_cte = linha.index("CT-e")
            idx_emis_cte = linha.index("Emissão") if "Emissão" in linha else -1
            if i + 1 < len(linhas):
                sub = [str(x).strip() for x in linhas[i+1]]
                if len(sub) > idx_cte:
                    cands = [str(x).strip() for x in sub[idx_cte:idx_cte+5] if str(x).strip()]
                    if cands: current_cte = cands[0]
                if idx_emis_cte != -1 and len(sub) > idx_emis_cte:
                    cands_em = [str(x).strip() for x in sub[idx_emis_cte:idx_emis_cte+5] if str(x).strip()]
                    if cands_em: current_emissao_cte = formatar_data_excel_somente_data(cands_em[0])
                current_nf, current_emissao_nf, current_peso, current_cub, current_valor = "", "", "", "", ""

        for cell in linha:
            if "Remetente:" in cell:
                cands = [x for x in linha if x and "Remetente" not in x]
                if cands: current_remetente = cands[0].replace('\n', ' ')
            if "Destinatário:" in cell:
                cands = [x for x in linha if x and "Destinatário" not in x]
                if cands: current_destinatario = cands[0].replace('\n', ' ')
                
        if "Peso" in linha and "Cub." in linha and "Valor" in linha:
            idx_peso, idx_cub, idx_valor = linha.index("Peso"), linha.index("Cub."), linha.index("Valor")
            idx_emis_nf = linha.index("Emissão") if "Emissão" in linha else -1
            for j in range(i+1, min(i+5, len(linhas))):
                sub = [str(x).strip() for x in linhas[j]]
                if "NF" in sub:
                    idx_nf = sub.index("NF")
                    c_nf = [x for x in sub[idx_nf+1:] if x]
                    if c_nf: current_nf = c_nf[0]
                    if idx_emis_nf != -1 and len(sub) > idx_emis_nf:
                        c_em_nf = [x for x in sub[idx_emis_nf:idx_emis_nf+5] if x]
                        if c_em_nf: current_emissao_nf = formatar_data_excel_somente_data(c_em_nf[0])
                    try: current_peso = [x for x in sub[max(0, idx_peso-1):idx_peso+3] if x][0]
                    except: pass
                    try: current_cub = [x for x in sub[max(0, idx_cub-1):idx_cub+3] if x][0]
                    except: pass
                    try: current_valor = [x for x in sub[max(0, idx_valor-1):idx_valor+3] if x][0]
                    except: pass
                    break

        if "Frete calculado" in linha and "Frete realizado" in linha:
            idx_calc, idx_real = linha.index("Frete calculado"), linha.index("Frete realizado")
            j = i + 1
            while j < len(linhas):
                sub = [str(x).strip() for x in linhas[j]]
                if "Total do Frete" in sub or "Total de documentos" in sub or ("Número" in sub and "CT-e" in sub):
                    break
                itens = [x for x in sub[:10] if x]
                if itens and len(sub) > max(idx_calc, idx_real):
                    nome = itens[0]
                    try: calc = float(sub[idx_calc]); real = float(sub[idx_real])
                    except: calc, real = 0.0, 0.0
                    diff = real - calc
                    status = "OK"
                    if diff > 0.01: status = "DIVERGÊNCIA (A MAIOR)"
                    elif diff < -0.01: status = "DIVERGÊNCIA (A MENOR)"
                    if "PEDÁGIO" in nome.upper() and diff > 0.01: status = "ALERTA: PEDÁGIO A MAIOR!"
                    dados_analiticos.append({
                        "CT-e": current_cte, "Emissão CT-e": current_emissao_cte,
                        "NF": current_nf, "Emissão NF": current_emissao_nf,
                        "Remetente": current_remetente, "Destinatário": current_destinatario,
                        "Peso": current_peso, "Cub": current_cub, "Valor NF": current_valor,
                        "Componente": nome, "Calculado": calc, "Realizado": real, "Diferença": diff, "Status": status
                    })
                j += 1
    return dados_analiticos

def extrair_dados_embarque(linhas):
    dados_embarque = []
    cur_emb, cur_dt_criacao, cur_transp, cur_origem, cur_destino = "", "", "", "", ""
    for i, row in enumerate(linhas):
        linha = [str(x).strip() for x in row]
        if not any(linha): continue
        if "Embarque" in linha and "Transportadora" in linha:
            idx_emb, idx_dt, idx_transp = linha.index("Embarque"), (linha.index("Dt. criação") if "Dt. criação" in linha else -1), linha.index("Transportadora")
            if i + 1 < len(linhas):
                sub = [str(x).strip() for x in linhas[i+1]]
                if len(sub) > idx_emb: cur_emb = sub[idx_emb].replace(".0", "")
                if idx_dt != -1 and len(sub) > idx_dt:
                    cands_dt = [str(x).strip() for x in sub[idx_dt:idx_dt+5] if str(x).strip()]
                    if cands_dt: cur_dt_criacao = formatar_data_excel_somente_data(cands_dt[0])
                if len(sub) > idx_transp:
                    cands_tr = [str(x).strip() for x in sub[idx_transp:idx_transp+8] if str(x).strip()]
                    if cands_tr: cur_transp = cands_tr[0]
        for cell in linha:
            if "Origem:" in cell:
                cands = [x for x in linha if x and "Origem" not in x]
                if cands: cur_origem = cands[0].replace('\n', ' ')
            if "Destino:" in cell:
                cands = [x for x in linha if x and "Destino" not in x]
                if cands: cur_destino = cands[0].replace('\n', ' ')
        if linha[0] == "Nome" and "Frete calculado" in linha and "Frete realizado" in linha:
            idx_calc, idx_real = linha.index("Frete calculado"), linha.index("Frete realizado")
            j = i + 1
            while j < len(linhas):
                sub = [str(x).strip() for x in linhas[j]]
                if not any(sub) or "Total" in sub[0] or "Pré-conhecimentos" in sub or "Embarque" in sub: break
                nome = sub[0]
                if nome and len(sub) > max(idx_calc, idx_real):
                    try: calc, real = float(sub[idx_calc]), float(sub[idx_real])
                    except: calc, real = 0.0, 0.0
                    diff = real - calc
                    status = "OK"
                    if diff > 0.01: status = "DIVERGÊNCIA (A MAIOR)"
                    elif diff < -0.01: status = "DIVERGÊNCIA (A MENOR)"
                    dados_embarque.append({
                        "Embarque ID": cur_emb, "Data Criação": cur_dt_criacao,
                        "Transportadora": cur_transp, "Origem": cur_origem, "Destino": cur_destino,
                        "Componente": nome, "Calculado": calc, "Realizado": real, "Diferença": diff, "Status": status
                    })
                j += 1
    return dados_embarque

# =============================================================================
#  GERADORES DE EXCEL
# =============================================================================
def gerar_excel_colorido(df_local):
    """Gera o arquivo ANALÍTICO PADRÃO (Colorido)"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_local.to_excel(writer, index=False, sheet_name='Auditoria')
        worksheet = writer.sheets['Auditoria']
        
        # Estilo Cabeçalho
        fill_cab = PatternFill(start_color="5C2EE9", end_color="5C2EE9", fill_type="solid")
        font_cab = Font(color="FFFFFF", bold=True)
        for col_num in range(1, len(df_local.columns) + 1):
            cell = worksheet.cell(row=1, column=col_num)
            cell.fill = fill_cab
            cell.font = font_cab

        # Cores Status
        fill_verde = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        fill_vermelho = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        fill_amarelo = PatternFill(start_color="F4D03F", end_color="F4D03F", fill_type="solid")

        cols = {col: i + 1 for i, col in enumerate(df_local.columns)}
        start_col = cols.get('Componente', 1)
        end_col = cols.get('Status', len(df_local.columns))
        
        for row_idx, row_data in enumerate(df_local.itertuples(), start=2):
            status = str(row_data.Status).upper()
            target_fill = fill_verde if "OK" in status else (fill_vermelho if "A MAIOR" in status else fill_amarelo)
            for col_idx in range(start_col, end_col + 1):
                worksheet.cell(row=row_idx, column=col_idx).fill = target_fill

        worksheet.auto_filter.ref = worksheet.dimensions
        for col in worksheet.columns:
            max_len = max([len(str(cell.value)) for cell in col])
            worksheet.column_dimensions[col[0].column_letter].width = min(max_len + 2, 40)
    return output.getvalue()

def gerar_excel_observacoes(df_filtrado):
    """Gera o NOVO arquivo CONSOLIDADO (Embarque | Observação)"""
    output = io.BytesIO()
    # Agrupa e concatena componentes
    df_obs = df_filtrado.groupby('Embarque ID')['Componente'].apply(lambda x: " - ".join(x)).reset_index()
    df_obs.columns = ['Embarque', 'Observação']
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_obs.to_excel(writer, index=False, sheet_name='Resumo')
        worksheet = writer.sheets['Resumo']
        fill_cab = PatternFill(start_color="5C2EE9", end_color="5C2EE9", fill_type="solid")
        font_cab = Font(color="FFFFFF", bold=True)
        for cell in worksheet[1]:
            cell.fill = fill_cab; cell.font = font_cab
        worksheet.column_dimensions['A'].width = 20
        worksheet.column_dimensions['B'].width = 80
    return output.getvalue()

# =============================================================================
#  INTERFACE STREAMLIT
# =============================================================================
st.sidebar.markdown("# 📊 CONSOLIDA\n### WORKSPACE")
st.sidebar.divider()
modulo = st.sidebar.radio("Navegação", ["Auditoria de Frete"])

if modulo == "Auditoria de Frete":
    st.title("Módulo de Extração Analítica")
    tab_cte, tab_emb = st.tabs(["📦 Pré-Conhecimentos (CT-e / NF)", "🚢 Embarques Globais"])

    with tab_cte:
        arquivo = st.file_uploader("Upload CT-e", type=['xlsx', 'csv'], key="u1")
        if arquivo and st.button("🚀 Analisar Arquivo CT-e"):
            linhas = processar_arquivo_bruto(arquivo)
            if linhas:
                dados = extrair_dados_pre_conhecimento(linhas)
                if dados:
                    df = pd.DataFrame(dados)
                    st.success(f"Encontrados {len(dados)} itens!")
                    st.dataframe(df, use_container_width=True)
                    excel = gerar_excel_colorido(df)
                    st.download_button("⬇️ Baixar Excel Analítico", data=excel, file_name="Auditoria_CTE.xlsx")

    with tab_emb:
        arquivo_emb = st.file_uploader("Upload Embarques", type=['xlsx', 'csv'], key="u2")
        
        if arquivo_emb:
            if 'arquivo_emb_atual' not in st.session_state or st.session_state['arquivo_emb_atual'] != arquivo_emb.name:
                st.session_state['arquivo_emb_atual'] = arquivo_emb.name
                if 'dados_emb_brutos' in st.session_state: del st.session_state['dados_emb_brutos']

            if st.button("🚀 Analisar Arquivo de Embarques"):
                linhas = processar_arquivo_bruto(arquivo_emb)
                if linhas:
                    dados = extrair_dados_embarque(linhas)
                    if dados: st.session_state['dados_emb_brutos'] = dados
            
            if 'dados_emb_brutos' in st.session_state:
                df_orig = pd.DataFrame(st.session_state['dados_emb_brutos'])
                st.divider()
                st.subheader("⚙️ Filtros")
                c1, c2 = st.columns(2)
                with c1:
                    sel_comp = st.multiselect("Componentes:", options=df_orig['Componente'].unique().tolist(), default=df_orig['Componente'].unique().tolist())
                with c2:
                    sel_div = st.selectbox("Divergência:", ["Todas", "Diferente de OK", "Somente A Maior", "Somente A Menor"])

                df_f = df_orig[df_orig['Componente'].isin(sel_comp)].copy()
                if sel_div == "Diferente de OK": df_f = df_f[df_f['Status'] != "OK"]
                elif sel_div == "Somente A Maior": df_f = df_f[df_f['Status'] == "DIVERGÊNCIA (A MAIOR)"]
                elif sel_div == "Somente A Menor": df_f = df_f[df_f['Status'] == "DIVERGÊNCIA (A MENOR)"]

                st.dataframe(df_f, use_container_width=True)

                if not df_f.empty:
                    st.divider()
                    st.subheader("📥 Downloads Disponíveis")
                    col_dl1, col_dl2 = st.columns(2)
                    with col_dl1:
                        # BOTÃO 1: PADRÃO COLORIDO
                        excel_padrao = gerar_excel_colorido(df_f)
                        st.download_button("⬇️ Baixar Relatório Analítico (Padrão)", data=excel_padrao, file_name="Auditoria_Embarques_Detalhado.xlsx")
                    with col_dl2:
                        # BOTÃO 2: NOVO CONSOLIDADO
                        excel_resumo = gerar_excel_observacoes(df_f)
                        st.download_button("⬇️ Baixar Resumo (Embarque | Obs)", data=excel_resumo, file_name="Divergencias_Consolidadas.xlsx")

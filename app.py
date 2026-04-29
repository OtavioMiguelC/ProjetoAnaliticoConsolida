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
#  FUNÇÕES DE UTILIDADE
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

# =============================================================================
#  LÓGICA DE EXTRAÇÃO - CT-e (ORIGINAL)
# =============================================================================
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

# =============================================================================
#  LÓGICA DE EXTRAÇÃO - EMBARQUE
# =============================================================================
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
#  GERAÇÃO DE EXCEL (CT-E - ORIGINAL)
# =============================================================================
def gerar_excel_colorido(dados, headers_param=None):
    output = io.BytesIO()
    df = pd.DataFrame(dados)
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Auditoria')
        workbook  = writer.book
        worksheet = writer.sheets['Auditoria']
        
        fill_cab = PatternFill(start_color="5C2EE9", end_color="5C2EE9", fill_type="solid")
        font_cab = Font(color="FFFFFF", bold=True)
        for col_num in range(1, len(df.columns) + 1):
            cell = worksheet.cell(row=1, column=col_num)
            cell.fill = fill_cab
            cell.font = font_cab

        fill_verde = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        fill_vermelho = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        fill_amarelo = PatternFill(start_color="F4D03F", end_color="F4D03F", fill_type="solid")

        cols = {col: i + 1 for i, col in enumerate(df.columns)}
        start_col = cols.get('Componente', 1)
        end_col = cols.get('Status', len(df.columns))
        
        for row_idx, row_data in enumerate(df.itertuples(), start=2):
            status = str(row_data.Status).upper()
            target_fill = None
            if "OK" in status: target_fill = fill_verde
            elif "A MAIOR" in status: target_fill = fill_vermelho
            elif "A MENOR" in status: target_fill = fill_amarelo
            
            if target_fill:
                for col_idx in range(start_col, end_col + 1):
                    worksheet.cell(row=row_idx, column=col_idx).fill = target_fill

        worksheet.auto_filter.ref = worksheet.dimensions
        for col in worksheet.columns:
            max_len = 0
            column_letter = col[0].column_letter
            for cell in col:
                try: 
                    if len(str(cell.value)) > max_len: max_len = len(str(cell.value))
                except: pass
            worksheet.column_dimensions[column_letter].width = min(max_len + 2, 40)

    return output.getvalue()

# =============================================================================
#  GERAÇÃO DE EXCEL CONSOLIDADO (EMBARQUES - NOVO)
# =============================================================================
def gerar_excel_observacoes(df_filtrado):
    output = io.BytesIO()
    
    # Agrupa por embarque e junta os nomes dos componentes em uma string separada por " - "
    df_obs = df_filtrado.groupby('Embarque ID')['Componente'].apply(lambda x: " - ".join(x)).reset_index()
    df_obs.columns = ['Embarque', 'Observação']
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_obs.to_excel(writer, index=False, sheet_name='Divergencias')
        worksheet = writer.sheets['Divergencias']
        
        fill_cab = PatternFill(start_color="5C2EE9", end_color="5C2EE9", fill_type="solid")
        font_cab = Font(color="FFFFFF", bold=True)
        for cell in worksheet[1]:
            cell.fill = fill_cab
            cell.font = font_cab
            
        worksheet.column_dimensions['A'].width = 20
        worksheet.column_dimensions['B'].width = 80

    return output.getvalue()

# =============================================================================
#  INTERFACE DO USUÁRIO (STREAMLIT)
# =============================================================================
st.sidebar.markdown("# 📊 CONSOLIDA")
st.sidebar.markdown("### WORKSPACE")
st.sidebar.divider()
modulo = st.sidebar.radio("Navegação", ["Auditoria de Frete"])

if modulo == "Auditoria de Frete":
    st.title("Módulo de Extração Analítica")
    st.caption("Faça o upload do seu relatório logístico para gerar alertas de divergências.")

    tab_cte, tab_emb = st.tabs(["📦 Pré-Conhecimentos (CT-e / NF)", "🚢 Embarques Globais"])

    # -----------------------------------------------------------------------------
    # ABA 1: CT-E (Comportamento Original Mantido)
    # -----------------------------------------------------------------------------
    with tab_cte:
        arquivo = st.file_uploader("Arraste o arquivo bruto aqui (CT-e)", type=['xlsx', 'csv'], key="u1")
        if arquivo:
            if st.button("🚀 Analisar Arquivo CT-e"):
                linhas = processar_arquivo_bruto(arquivo)
                if linhas:
                    dados = extrair_dados_pre_conhecimento(linhas)
                    if dados:
                        st.success(f"Encontrados {len(dados)} itens analíticos!")
                        df = pd.DataFrame(dados)
                        st.dataframe(df, use_container_width=True)
                        
                        headers = ["CT-e", "Emissão CT-e", "NF", "Emissão NF", "Remetente", "Destinatário", "Peso", "Cub", "Valor NF", "Componente", "Previsto (R$)", "Realizado (R$)", "Diferença (R$)", "Status"]
                        excel = gerar_excel_colorido(dados, headers)
                        st.download_button("⬇️ Baixar Excel Analítico", data=excel, file_name="Auditoria_CTE.xlsx")

    # -----------------------------------------------------------------------------
    # ABA 2: EMBARQUES (Nova lógica de Filtro e Agrupamento)
    # -----------------------------------------------------------------------------
    with tab_emb:
        arquivo_emb = st.file_uploader("Arraste o arquivo bruto aqui (Embarques)", type=['xlsx', 'csv'], key="u2")

        if arquivo_emb:
            # Controle de estado para garantir que recarregue se o arquivo mudar
            if 'arquivo_emb_atual' not in st.session_state or st.session_state['arquivo_emb_atual'] != arquivo_emb.name:
                st.session_state['arquivo_emb_atual'] = arquivo_emb.name
                if 'dados_emb_brutos' in st.session_state:
                    del st.session_state['dados_emb_brutos']

            if st.button("🚀 Analisar Arquivo de Embarques"):
                with st.spinner("Processando..."):
                    linhas = processar_arquivo_bruto(arquivo_emb)
                    if linhas:
                        dados = extrair_dados_embarque(linhas)
                        if dados:
                            st.session_state['dados_emb_brutos'] = dados
                            st.success(f"Arquivo lido! {len(dados)} componentes encontrados.")

            # Se já analisou e tem dados na memória, mostra a interface de filtros
            if 'dados_emb_brutos' in st.session_state:
                df_original = pd.DataFrame(st.session_state['dados_emb_brutos'])
                
                st.divider()
                st.subheader("⚙️ Filtros de Componentes e Divergências")
                col1, col2 = st.columns(2)
                
                with col1:
                    todos_componentes = df_original['Componente'].unique().tolist()
                    comps_selecionados = st.multiselect(
                        "Selecione os Componentes para manter no relatório:", 
                        options=todos_componentes, 
                        default=todos_componentes
                    )
                
                with col2:
                    tipo_div = st.selectbox(
                        "Tipo de Divergência:", 
                        ["Todas (A Maior, A Menor, Zero)", "Somente Divergências (Diferente de OK)", "Somente A Maior", "Somente A Menor"]
                    )

                # Aplicação dos filtros
                df_filtrado = df_original[df_original['Componente'].isin(comps_selecionados)].copy()
                
                if tipo_div == "Somente Divergências (Diferente de OK)":
                    df_filtrado = df_filtrado[df_filtrado['Status'] != "OK"]
                elif tipo_div == "Somente A Maior":
                    df_filtrado = df_filtrado[df_filtrado['Status'] == "DIVERGÊNCIA (A MAIOR)"]
                elif tipo_div == "Somente A Menor":
                    df_filtrado = df_filtrado[df_filtrado['Status'] == "DIVERGÊNCIA (A MENOR)"]

                st.write(f"**Registros mantidos após filtros:** {len(df_filtrado)}")
                st.dataframe(df_filtrado, use_container_width=True)

                if not df_filtrado.empty:
                    st.divider()
                    excel_obs = gerar_excel_observacoes(df_filtrado)
                    st.download_button(
                        label="📄 Baixar Planilha Consolidada (Embarque | Observação)",
                        data=excel_obs,
                        file_name=f"Divergencias_Agrupadas_{datetime.date.today()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

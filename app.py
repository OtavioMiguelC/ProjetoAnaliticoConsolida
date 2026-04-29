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
    page_title="Consolida Workspace - Logística",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main { background-color: #121212; }
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
#  LÓGICA DE EXTRAÇÃO DE EMBARQUE
# =============================================================================
def extrair_dados_embarque(linhas):
    dados_embarque = []
    cur_emb, cur_dt_criacao, cur_transp, cur_origem, cur_destino = "", "", "", "", ""
    
    for i, row in enumerate(linhas):
        linha = [str(x).strip() for x in row]
        if not any(linha): continue
        
        if "Embarque" in linha and "Transportadora" in linha:
            idx_emb = linha.index("Embarque")
            idx_dt = linha.index("Dt. criação") if "Dt. criação" in linha else -1
            idx_transp = linha.index("Transportadora")
            
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
#  GERAÇÃO DE EXCEL CONSOLIDADO (DIVERGÊNCIAS)
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
#  INTERFACE STREAMLIT
# =============================================================================
st.sidebar.title("📊 CONSOLIDA TMS")
op_modulo = st.sidebar.radio("Módulos", ["Auditoria de Embarque"])

if op_modulo == "Auditoria de Embarque":
    st.title("🚢 Análise por Embarque")
    st.info("Faça o upload do arquivo, clique em analisar e refine os dados antes de baixar.")

    arquivo_emb = st.file_uploader("Upload do espelho de embarque", type=['xlsx', 'csv'])

    # Controle de estado: se o usuário trocar o arquivo, limpa o cache anterior
    if arquivo_emb:
        if 'nome_arquivo_atual' not in st.session_state or st.session_state['nome_arquivo_atual'] != arquivo_emb.name:
            st.session_state['nome_arquivo_atual'] = arquivo_emb.name
            if 'dados_brutos' in st.session_state:
                del st.session_state['dados_brutos']

        # Botão explícito para forçar a extração de dados
        if st.button("🚀 Analisar Arquivo"):
            with st.spinner("Processando dados..."):
                linhas = processar_arquivo_bruto(arquivo_emb)
                if linhas:
                    dados = extrair_dados_embarque(linhas)
                    if dados:
                        st.session_state['dados_brutos'] = dados
                        st.success(f"Arquivo lido com sucesso! {len(dados)} componentes encontrados.")
                    else:
                        st.warning("Nenhum dado de componente de frete foi encontrado neste arquivo.")
                else:
                    st.error("Falha ao ler as linhas do arquivo Excel.")

        # ==========================================
        # SEÇÃO DE FILTROS (Só aparece após a análise)
        # ==========================================
        if 'dados_brutos' in st.session_state:
            df_original = pd.DataFrame(st.session_state['dados_brutos'])
            
            st.divider()
            st.subheader("⚙️ Filtros de Componentes e Divergências")
            col1, col2 = st.columns(2)
            
            with col1:
                # Caixa de componentes disponíveis para flegar/desflegar
                todos_componentes = df_original['Componente'].unique().tolist()
                comps_selecionados = st.multiselect(
                    "Selecione os Componentes para manter no relatório:", 
                    options=todos_componentes, 
                    default=todos_componentes
                )
            
            with col2:
                # Caixa de tipo de divergência
                tipo_div = st.selectbox(
                    "Tipo de Divergência:", 
                    ["Todas (A Maior, A Menor, Zero)", "Somente Divergências (Diferente de OK)", "Somente A Maior", "Somente A Menor"]
                )

            # Aplica os filtros na base
            df_filtrado = df_original[df_original['Componente'].isin(comps_selecionados)].copy()
            
            if tipo_div == "Somente Divergências (Diferente de OK)":
                df_filtrado = df_filtrado[df_filtrado['Status'] != "OK"]
            elif tipo_div == "Somente A Maior":
                df_filtrado = df_filtrado[df_filtrado['Status'] == "DIVERGÊNCIA (A MAIOR)"]
            elif tipo_div == "Somente A Menor":
                df_filtrado = df_filtrado[df_filtrado['Status'] == "DIVERGÊNCIA (A MENOR)"]

            st.write(f"**Registros mantidos após os filtros:** {len(df_filtrado)}")
            st.dataframe(df_filtrado, use_container_width=True)

            # ==========================================
            # SEÇÃO DE DOWNLOAD
            # ==========================================
            if not df_filtrado.empty:
                st.divider()
                st.subheader("📥 Exportação")
                excel_obs = gerar_excel_observacoes(df_filtrado)
                
                st.download_button(
                    label="📄 Baixar Planilha Consolidada (Embarque | Observação)",
                    data=excel_obs,
                    file_name=f"Divergencias_Agrupadas_{datetime.date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.caption("Nota: A planilha gerada unifica as divergências filtradas em uma única linha por Embarque ID, separando os componentes por traço (' - ').")
            else:
                st.warning("Atenção: Os filtros aplicados removeram todos os dados. Não há o que exportar.")

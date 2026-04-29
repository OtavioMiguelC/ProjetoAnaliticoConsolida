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
        
        # Identificação do cabeçalho do Embarque
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

        # Identificação dos componentes de frete
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
    """Gera planilha com Embarque | Observação consolidada"""
    output = io.BytesIO()
    
    # Agrupar por Embarque ID e concatenar componentes divergentes
    df_obs = df_filtrado.groupby('Embarque ID')['Componente'].apply(lambda x: " - ".join(x)).reset_index()
    df_obs.columns = ['Embarque', 'Observação']
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_obs.to_excel(writer, index=False, sheet_name='Divergências')
        worksheet = writer.sheets['Divergências']
        
        # Estilo do cabeçalho
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
    st.info("Filtre os componentes e tipos de divergência antes de gerar o relatório final.")

    arquivo_emb = st.file_uploader("Upload do espelho de embarque", type=['xlsx', 'csv'])

    if arquivo_emb:
        # Processamento inicial e armazenamento no Session State
        if 'dados_brutos' not in st.session_state or st.sidebar.button("🔄 Recarregar Arquivo"):
            linhas = processar_arquivo_bruto(arquivo_emb)
            if linhas:
                st.session_state['dados_brutos'] = extrair_dados_embarque(linhas)

        if 'dados_brutos' in st.session_state:
            df_original = pd.DataFrame(st.session_state['dados_brutos'])
            
            # --- ÁREA DE FILTROS ---
            st.subheader("⚙️ Filtros de Refinamento")
            col1, col2 = st.columns(2)
            
            with col1:
                todos_componentes = df_original['Componente'].unique().tolist()
                comps_selecionados = st.multiselect("Selecione os Componentes para manter:", 
                                                   options=todos_componentes, 
                                                   default=todos_componentes)
            
            with col2:
                tipo_div = st.selectbox("Tipo de Divergência:", 
                                       ["Todas (A Maior, A Menor, Zero)", "Somente Divergências (Diferente de OK)", "Somente A Maior"])

            # Aplicação dos filtros no DataFrame
            df_filtrado = df_original[df_original['Componente'].isin(comps_selecionados)].copy()
            
            if tipo_div == "Somente Divergências (Diferente de OK)":
                df_filtrado = df_filtrado[df_filtrado['Status'] != "OK"]
            elif tipo_div == "Somente A Maior":
                df_filtrado = df_filtrado[df_filtrado['Status'] == "DIVERGÊNCIA (A MAIOR)"]

            # --- VISUALIZAÇÃO ---
            st.divider()
            st.write(f"**Registros encontrados após filtros:** {len(df_filtrado)}")
            st.dataframe(df_filtrado, use_container_width=True)

            # --- GERAÇÃO DE RELATÓRIOS ---
            if not df_filtrado.empty:
                col_btn1, col_btn2 = st.columns(2)
                
                with col_btn1:
                    # Relatório de Observações (O que você pediu especificamente)
                    excel_obs = gerar_excel_observacoes(df_filtrado)
                    st.download_button(
                        label="📄 Baixar Planilha de Observações (Consolidada)",
                        data=excel_obs,
                        file_name=f"Observacoes_Divergencias_{datetime.date.today()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                st.caption("A planilha de observações contém duas colunas: **Embarque** e **Observação**, consolidando os itens sem duplicar o ID do embarque.")
            else:
                st.warning("Nenhum dado disponível com os filtros selecionados.")

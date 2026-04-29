import streamlit as st
import pandas as pd
import io
import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

# =============================================================================
# CONFIGURAÇÕES DA PÁGINA
# =============================================================================
st.set_page_config(
    page_title="Consolida Workspace",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main { background-color: #121212; }
    div.stButton > button:first-child { background-color: #5C2EE9; color: white; border-radius: 8px; width: 100%; }
    .stSelectbox, .stMultiSelect, .stNumberInput { color: white; }
    </style>
    """, unsafe_allow_html=True)

# =============================================================================
# FUNÇÕES DE UTILIDADE E EXTRAÇÃO
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
        
        if "Origem:" in str(row):
            cands = [x for x in linha if x and "Origem" not in x]
            if cands: cur_origem = cands[0].replace('\n', ' ')
        if "Destino:" in str(row):
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
                    dados_embarque.append({
                        "Embarque ID": cur_emb, "Data Criação": cur_dt_criacao,
                        "Transportadora": cur_transp, "Origem": cur_origem, "Destino": cur_destino,
                        "Componente": nome, "Calculado": calc, "Realizado": real, "Diferença": real - calc
                    })
                j += 1
    return dados_embarque

# =============================================================================
# LÓGICA DE STATUS COM TOLERÂNCIA
# =============================================================================
def definir_status(diff, tolerancia):
    if abs(diff) <= tolerancia:
        return "OK"
    elif diff > tolerancia:
        return "DIVERGÊNCIA (A MAIOR)"
    else:
        return "DIVERGÊNCIA (A MENOR)"

# =============================================================================
# GERADOR DE EXCEL UNIFICADO (DUAS ABAS)
# =============================================================================
def gerar_excel_unificado_embarque(df_analitico):
    output = io.BytesIO()
    
    # Preparar dados da segunda aba (Observações)
    # Somente o que não for OK entra na observação
    df_obs_input = df_analitico[df_analitico['Status'] != "OK"]
    df_resumo = df_obs_input.groupby('Embarque ID')['Componente'].apply(lambda x: " - ".join(x)).reset_index()
    df_resumo.columns = ['Embarque', 'Observação']

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # ABA 1: ANALÍTICO
        df_analitico.to_excel(writer, index=False, sheet_name='Analítico Detalhado')
        ws1 = writer.sheets['Analítico Detalhado']
        
        # ABA 2: OBSERVAÇÕES
        df_resumo.to_excel(writer, index=False, sheet_name='Resumo Observações')
        ws2 = writer.sheets['Resumo Observações']

        # Estilização
        fill_cab = PatternFill(start_color="5C2EE9", end_color="5C2EE9", fill_type="solid")
        font_cab = Font(color="FFFFFF", bold=True)
        fill_vermelho = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        fill_amarelo = PatternFill(start_color="F4D03F", end_color="F4D03F", fill_type="solid")
        fill_verde = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

        # Formatar Aba 1
        for cell in ws1[1]:
            cell.fill = fill_cab; cell.font = font_cab
        
        for row_idx, row_data in enumerate(df_analitico.itertuples(), start=2):
            status = str(row_data.Status)
            target_fill = fill_verde if status == "OK" else (fill_vermelho if "A MAIOR" in status else fill_amarelo)
            # Pintar colunas de valor e status
            for col_idx in range(6, 11): # Componente até Status
                ws1.cell(row=row_idx, column=col_idx).fill = target_fill

        # Formatar Aba 2
        for cell in ws2[1]:
            cell.fill = fill_cab; cell.font = font_cab
        ws2.column_dimensions['A'].width = 20
        ws2.column_dimensions['B'].width = 80

    return output.getvalue()

# =============================================================================
# INTERFACE STREAMLIT
# =============================================================================
st.sidebar.markdown("# 📊 CONSOLIDA\n### WORKSPACE")
st.sidebar.divider()
modulo = st.sidebar.radio("Navegação", ["Auditoria de Frete"])

if modulo == "Auditoria de Frete":
    st.title("Módulo de Extração Analítica")
    tab_cte, tab_emb = st.tabs(["📦 Pré-Conhecimentos", "🚢 Embarques Globais"])

    with tab_cte:
        st.info("Funcionalidade original mantida para CT-e.")
        # (Código de CT-e omitido aqui para brevidade, mas permanece igual ao seu original)

    with tab_emb:
        arquivo_emb = st.file_uploader("Upload Embarques", type=['xlsx', 'csv'], key="u_emb")
        
        if arquivo_emb:
            if st.button("🚀 Analisar Arquivo de Embarques"):
                linhas = processar_arquivo_bruto(arquivo_emb)
                if linhas:
                    st.session_state['dados_emb_brutos'] = extrair_dados_embarque(linhas)

            if 'dados_emb_brutos' in st.session_state:
                df_base = pd.DataFrame(st.session_state['dados_emb_brutos'])
                
                st.divider()
                st.subheader("⚙️ Parâmetros de Refino")
                c1, c2, c3 = st.columns([2, 2, 1])
                
                with c1:
                    sel_comp = st.multiselect("Flegar Componentes:", options=df_base['Componente'].unique().tolist(), default=df_base['Componente'].unique().tolist())
                with c2:
                    sel_div = st.selectbox("Mostrar na Tela:", ["Todas", "Divergências", "A Maior", "A Menor"])
                with c3:
                    tolerancia = st.number_input("Tolerância (R$):", min_value=0.0, value=0.01, step=0.01, help="Diferenças abaixo deste valor serão consideradas OK.")

                # Aplicar Tolerância e Status dinamicamente
                df_f = df_base[df_base['Componente'].isin(sel_comp)].copy()
                df_f['Status'] = df_f['Diferença'].apply(lambda x: definir_status(x, tolerancia))

                # Filtrar visualização da tela
                if sel_div == "Divergências": df_f_view = df_f[df_f['Status'] != "OK"]
                elif sel_div == "A Maior": df_f_view = df_f[df_f['Status'] == "DIVERGÊNCIA (A MAIOR)"]
                elif sel_div == "A Menor": df_f_view = df_f[df_f['Status'] == "DIVERGÊNCIA (A MENOR)"]
                else: df_f_view = df_f

                st.write(f"**Itens Analisados:** {len(df_f_view)}")
                st.dataframe(df_f_view, use_container_width=True)

                if not df_f.empty:
                    st.divider()
                    # Gerar arquivo único com as duas abas
                    excel_final = gerar_excel_unificado_embarque(df_f)
                    
                    st.download_button(
                        label="⬇️ Baixar Relatório Unificado (Analítico + Observações)",
                        data=excel_final,
                        file_name=f"Auditoria_Embarque_Unificada_{datetime.date.today()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.caption("O arquivo contém a aba 'Analítico Detalhado' e a aba 'Resumo Observações' (agrupada por embarque).")

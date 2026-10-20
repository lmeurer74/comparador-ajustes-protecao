import streamlit as st
import pandas as pd
import openpyxl
from mapeamentos import carregar_biblioteca_prioritaria
from leitor_excel import obter_abas_validas
from leitor_rele import DriverArtecheAdaTech
from comparador import executar_comparacao, style_dataframe_resultados

st.set_page_config(page_title="Comparador de Ajustes", layout="wide")
st.title("⚡ Comparador de Ajustes de Proteção (Modular)")

if 'arquivos_lidos' not in st.session_state:
    st.session_state.arquivos_lidos = False
if 'comparacao_concluida' not in st.session_state:
    st.session_state.comparacao_concluida = False

col1, col2, col3, col4 = st.columns(4)
with col1:
    fabricante = st.selectbox("1. Equipamento:", ["ARTECHE_ADATECH"])
with col2:
    file_xlsx = st.file_uploader("2. Ordem de Ajuste (.xlsx):", type=["xlsx"])
with col3:
    file_cid = st.file_uploader("3. Arquivo (.cid):", type=["cid", "xml", "icd"])
with col4:
    file_bib = st.file_uploader("4. De-Para Opcional:", type=["xlsx"])

if file_xlsx and file_cid:
    st.session_state.arquivos_lidos = True
    wb_in = openpyxl.load_workbook(file_xlsx, data_only=True)
    abas_visiveis = obter_abas_validas(wb_in)
    
    st.markdown("### 🔗 Selecione os Grupos para Comparar")
    
    dados_iniciais = [{"Processar": True, "Aba": nome, "Grupo 1": True, "Grupo 2": True, "Grupo 3": True} for nome in abas_visiveis]
    df_selecao = st.data_editor(pd.DataFrame(dados_iniciais), hide_index=True)

    if st.button("🔍 Iniciar Comparação", type="primary"):
        with st.spinner("Comparando ajustes..."):
            # Para extrair a biblioteca da memória, precisamos passar os nós lidos
            _, dois_por_ln, _ = DriverArtecheAdaTech.carregar_equipamento(file_cid.getvalue())
            
            de_para_cod, de_para_campos = carregar_biblioteca_prioritaria(file_bib, fabricante, dois_por_ln)
            
            sucesso, excel_io, resultados_web = executar_comparacao(
                df_selecao, wb_in, fabricante, file_cid.getvalue(), 
                de_para_cod, de_para_campos, DriverArtecheAdaTech
            )
            
            if sucesso:
                st.session_state.comparacao_concluida = True
                st.session_state.excel_io = excel_io
                st.session_state.resultados_web = resultados_web
                st.success("✅ Comparação finalizada com sucesso!")
            else:
                st.warning("⚠️ Nenhum dado válido foi encontrado para comparar (verifique se as abas contêm parâmetros visíveis).")

if st.session_state.comparacao_concluida:
    st.download_button(
        "📥 Exportar Planilha de Resultado (.xlsx)", 
        data=st.session_state.excel_io, 
        file_name="Resultado_Conferencia.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.markdown("---")
    st.subheader("📊 Relatório de Conferência")
    abas_nomes = [res["nome"] for res in st.session_state.resultados_web]
    tabs = st.tabs(abas_nomes)
    
    for i, tab in enumerate(tabs):
        res = st.session_state.resultados_web[i]
        with tab:
            st.markdown(f"**Total Processado:** {res['total']} parâmetros")
            c1, c2, c3 = st.columns(3)
            c1.info(f"**Grupo 1:** {res['corretos'][1]} OK / {res['diferentes'][1]} Diferentes")
            c2.info(f"**Grupo 2:** {res['corretos'][2]} OK / {res['diferentes'][2]} Diferentes")
            c3.info(f"**Grupo 3:** {res['corretos'][3]} OK / {res['diferentes'][3]} Diferentes")
            
            st.dataframe(
                res["df"].style.apply(style_dataframe_resultados, axis=1),
                use_container_width=True,
                hide_index=True
            )
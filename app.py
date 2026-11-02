import streamlit as st
import pandas as pd
import openpyxl
from mapeamentos import carregar_biblioteca_prioritaria
from leitor_excel import obter_abas_validas
from leitor_rele import DriverArtecheAdaTech, DriverSEL
from comparador import executar_comparacao, style_dataframe_resultados

st.set_page_config(page_title="Comparador de Ajustes", layout="wide")

# ==============================================================================
# PAINEL LATERAL (SIDEBAR) PARA CONFIGURAÇÃO DE COLUNAS DO EXCEL
# ==============================================================================
with st.sidebar:
    st.header("⚙️ Configurações de Leitura")
    st.markdown("Indique em quais colunas da Ordem de Ajuste (Excel) estão os valores:")
    
    col_codigo = st.text_input("Coluna do Código CID/SEL:", value="B", max_chars=1)
    col_desc = st.text_input("Coluna da Descrição:", value="C", max_chars=1)
    st.markdown("---")
    st.markdown("**Colunas de Ajuste por Grupo:**")
    col_g1 = st.text_input("Ajuste Grupo 1 (ex: F):", value="F", max_chars=1)
    col_g2 = st.text_input("Ajuste Grupo 2 (ex: G):", value="G", max_chars=1)
    col_g3 = st.text_input("Ajuste Grupo 3 (ex: H):", value="H", max_chars=1)
    
    col_codigo = col_codigo.upper() if col_codigo.isalpha() else "B"
    col_desc = col_desc.upper() if col_desc.isalpha() else "C"
    col_g1 = col_g1.upper() if col_g1.isalpha() else "F"
    col_g2 = col_g2.upper() if col_g2.isalpha() else "G"
    col_g3 = col_g3.upper() if col_g3.isalpha() else "H"

# ==============================================================================
# INTERFACE PRINCIPAL
# ==============================================================================
st.title("⚡ Comparador de Ajustes de Proteção (Modular)")

if 'arquivos_lidos' not in st.session_state:
    st.session_state.arquivos_lidos = False
if 'comparacao_concluida' not in st.session_state:
    st.session_state.comparacao_concluida = False

col1, col2, col3, col4 = st.columns(4)
with col1:
    fabricante = st.selectbox("1. Fabricante:", ["SEL (Schweitzer)", "ARTECHE_ADATECH"])
with col2:
    file_xlsx = st.file_uploader("2. Ordem de Ajuste (.xlsx):", type=["xlsx"])
with col3:
    file_cid = st.file_uploader("3. Arq. Equipamento (.cid, .rdb, .zip):", type=["cid", "xml", "icd", "rdb", "zip"])
with col4:
    file_bib = st.file_uploader("4. De-Para Opcional:", type=["xlsx"])

if file_xlsx and file_cid:
    st.session_state.arquivos_lidos = True
    wb_in = openpyxl.load_workbook(file_xlsx, data_only=True)
    abas_visiveis = obter_abas_validas(wb_in)
    
    bytes_rele = file_cid.getvalue()
    nome_arquivo_rele = file_cid.name
    
    rele_escolhido = None
    if fabricante == "SEL (Schweitzer)":
        driver_ativo = DriverSEL
        reles_encontrados = DriverSEL.listar_reles(bytes_rele, nome_arquivo_rele)
        
        if not reles_encontrados:
            st.error("Nenhum relé/vão encontrado dentro do arquivo fornecido. Verifique se o formato está correto.")
        else:
            rele_escolhido = st.selectbox("🎯 Múltiplos relés encontrados no arquivo. Selecione qual deseja comparar:", reles_encontrados)
    else:
        driver_ativo = DriverArtecheAdaTech

    st.markdown("### 🔗 Selecione os Grupos para Comparar")
    dados_iniciais = [{"Processar": True, "Aba": nome, "Grupo 1": True, "Grupo 2": True, "Grupo 3": True} for nome in abas_visiveis]
    df_selecao = st.data_editor(pd.DataFrame(dados_iniciais), hide_index=True)

    if st.button("🔍 Iniciar Comparação", type="primary"):
        if fabricante == "SEL (Schweitzer)" and not rele_escolhido:
            st.warning("Por favor, selecione o relé na lista antes de continuar.")
        else:
            with st.spinner("Comparando ajustes..."):
                _, dois_por_ln, _ = driver_ativo.carregar_equipamento(bytes_rele, nome_arquivo_rele, rele_escolhido)
                de_para_cod, de_para_campos = carregar_biblioteca_prioritaria(file_bib, fabricante, dois_por_ln)
                
                conf_colunas = {
                    "codigo": openpyxl.utils.column_index_from_string(col_codigo),
                    "desc": openpyxl.utils.column_index_from_string(col_desc),
                    "ajuste_g1": openpyxl.utils.column_index_from_string(col_g1),
                    "ajuste_g2": openpyxl.utils.column_index_from_string(col_g2),
                    "ajuste_g3": openpyxl.utils.column_index_from_string(col_g3)
                }
                
                sucesso, excel_io, resultados_web = executar_comparacao(
                    df_selecao, wb_in, fabricante, bytes_rele, nome_arquivo_rele, rele_escolhido, 
                    de_para_cod, de_para_campos, driver_ativo, conf_colunas
                )
                
                if sucesso:
                    st.session_state.comparacao_concluida = True
                    st.session_state.excel_io = excel_io
                    st.session_state.resultados_web = resultados_web
                    st.success("✅ Comparação finalizada com sucesso!")
                else:
                    st.warning("⚠️ Nenhum dado válido foi encontrado para comparar (verifique as colunas selecionadas).")

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
            if 1 in res['corretos']: c1.info(f"**Grupo 1:** {res['corretos'][1]} OK / {res['diferentes'][1]} Diferentes")
            if 2 in res['corretos']: c2.info(f"**Grupo 2:** {res['corretos'][2]} OK / {res['diferentes'][2]} Diferentes")
            if 3 in res['corretos']: c3.info(f"**Grupo 3:** {res['corretos'][3]} OK / {res['diferentes'][3]} Diferentes")
            
            st.dataframe(
                res["df"].style.apply(style_dataframe_resultados, axis=1),
                use_container_width=True,
                hide_index=True
            )

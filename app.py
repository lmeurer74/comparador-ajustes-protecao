import streamlit as st
import pandas as pd
import openpyxl
# Importações dos seus novos módulos
from mapeamentos import carregar_biblioteca_prioritaria
from leitor_excel import obter_abas_validas
from leitor_rele import DriverArtecheAdaTech
from comparador import executar_comparacao

st.set_page_config(page_title="Comparador de Ajustes", layout="wide")
st.title("⚡ Comparador de Ajustes de Proteção (Modular)")

# Controladores de Estado (Session State) para habilitar/desabilitar botões
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

# Lógica 1: Só libera configuração das abas se Excel e CID estiverem anexados
if file_xlsx and file_cid:
    st.session_state.arquivos_lidos = True
    
    wb_in = openpyxl.load_workbook(file_xlsx, data_only=True)
    abas_visiveis = obter_abas_validas(wb_in)
    
    st.markdown("### 🔗 Selecione os Grupos para Comparar")
    
    # Prepara a tabela de seleção para o usuário
    dados_iniciais = [{"Processar": True, "Aba": nome, "Grupo 1": True, "Grupo 2": True, "Grupo 3": True} for nome in abas_visiveis]
    df_selecao = st.data_editor(pd.DataFrame(dados_iniciais), hide_index=True)

    # Lógica 2: O botão de executar só aparece depois que os arquivos foram lidos
    if st.button("🔍 Iniciar Comparação", type="primary"):
        with st.spinner("Comparando ajustes..."):
            # 1. Carrega os mapeamentos (Prioriza arquivo local)
            de_para_cod, de_para_campos = carregar_biblioteca_prioritaria(file_bib, fabricante)
            
            # 2. Roda o motor
            sucesso, excel_io, resultados_web = executar_comparacao(
                df_selecao, wb_in, fabricante, file_cid.read(), 
                de_para_cod, de_para_campos, DriverArtecheAdaTech
            )
            
            if sucesso:
                st.session_state.comparacao_concluida = True
                st.session_state.excel_io = excel_io
                st.session_state.resultados_web = resultados_web
                st.success("✅ Comparação finalizada com sucesso!")
            else:
                st.warning("Nenhum dado válido para comparar.")

# Lógica 3: Botão de Download e Resultados só aparecem se a comparação concluiu
if st.session_state.comparacao_concluida:
    st.download_button(
        "📥 Exportar Planilha de Resultado (.xlsx)", 
        data=st.session_state.excel_io, 
        file_name="Resultado_Conferencia.xlsx"
    )
    
    # Renderiza a Tabela na Tela
    abas_nomes = [res["nome"] for res in st.session_state.resultados_web]
    tabs = st.tabs(abas_nomes)
    for i, tab in enumerate(tabs):
        with tab:
            st.dataframe(st.session_state.resultados_web[i]["df"], use_container_width=True)

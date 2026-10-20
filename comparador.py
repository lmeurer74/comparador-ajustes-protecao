import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import pandas as pd
import io

def normalizar_para_comparacao(valor):
    # (Sua função de normalização: YES/NO, floats, etc)
    pass

def executar_comparacao(df_selecao_usuario, wb_in, fabricante, bytes_cid, de_para_cod, de_para_campos, driver_rele):
    # Carrega dados do Relé
    doi_por_codigo, dois_por_ln, ns = driver_rele.carregar_equipamento(bytes_cid)
    
    wb_out = openpyxl.Workbook()
    aba_padrao = wb_out.active
    resultados_web = []
    processou_algo = False
    
    # Itera apenas sobre as abas que o usuário selecionou no UI
    for index, row in df_selecao_usuario.iterrows():
        if not row["Processar"]: continue
        
        nome_aba = row["Aba"]
        grupos_ativos = [g for g in [1, 2, 3] if row[f"Grupo {g}"]]
        if not grupos_ativos: continue

        ws_in = wb_in[nome_aba]
        # Aqui o leitor_excel já filtrou e mandou só as linhas visíveis!
        from leitor_excel import extrair_parametros_visiveis
        itens_oa = extrair_parametros_visiveis(ws_in, de_para_cod, de_para_campos, dois_por_ln)
        
        if not itens_oa: continue
        processou_algo = True
        
        # (AQUI VAI O SEU LOOP DE CRIAÇÃO DA PLANILHA wb_out E DO DATAFRAME DA TELA)
        # Compara OA vs CID
        # Se cod_cid não existe, marca como "-" (não traz coisas extras do relé)
        
    if processou_algo:
        wb_out.remove(aba_padrao)
        excel_io = io.BytesIO()
        wb_out.save(excel_io)
        excel_io.seek(0)
        return True, excel_io, resultados_web
    return False, None, []

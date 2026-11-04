import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import pandas as pd
import io
from leitor_excel import extrair_parametros_visiveis

def normalizar_valor(valor):
    """Apenas limpa o valor de espaços ou textos específicos de engenharia sem alterar números."""
    if pd.isna(valor) or valor is None: return ""
    s = str(valor).strip().upper()
    if s == "IEC DEFINE TIME": return "IEC DEFINITE TIME"
    if s.endswith(" A"):
        num_part = s[:-2].strip().replace(',', '.')
        if num_part.replace('.', '', 1).isdigit():
            s = num_part
    return s

def valores_sao_iguais(val_oa, val_cid):
    """Compara dois valores de forma inteligente, lidando com decimais, vírgulas, booleanos e português."""
    s1 = normalizar_valor(val_oa)
    s2 = normalizar_valor(val_cid)
    
    # 1. Comparação exata de texto
    if s1 == s2:
        return True
        
    # 2. Equivalências Lógicas (Inglês e Português)
    bool_true = ["YES", "Y", "SIM", "S", "TRUE", "1", "1.0", "1.00", "ON", "ENABLE", "ENABLED"]
    bool_false = ["NO", "N", "NÃO", "NAO", "FALSE", "0", "0.0", "0.00", "OFF", "DISABLE", "DISABLED"]
    
    if (s1 in bool_true and s2 in bool_true) or (s1 in bool_false and s2 in bool_false):
        return True
        
    # 3. Equivalência Matemática (Lida com ponto vs vírgula e zeros à direita. Ex: 0,1 == 0.10)
    try:
        f1 = float(s1.replace(',', '.'))
        f2 = float(s2.replace(',', '.'))
        # Se a diferença matemática for praticamente zero, são o mesmo número
        return abs(f1 - f2) < 1e-6
    except ValueError:
        pass
        
    return False

def style_dataframe_resultados(row):
    """Aplica o estilo visual verde/vermelho na tabela do Streamlit."""
    styles = [''] * len(row)
    for i, col in enumerate(row.index):
        if 'Relé' in col:
            # Obtém o número do grupo a partir do nome da coluna (ex: "Relé G1" -> "1")
            grp_num = col.split(' ')[-1].replace('G', '')
            col_oa = f"OA G{grp_num}"
            val_cid = row[col]
            val_oa = row.get(col_oa, "")
            
            if val_cid == "-":
                styles[i] = 'color: gray; text-align: center;'
            elif valores_sao_iguais(val_oa, val_cid):
                styles[i] = 'color: #006100; font-weight: bold;'
            else:
                styles[i] = 'background-color: #FFC7CE; color: #9C0006; font-weight: bold;'
    return styles

def executar_comparacao(df_selecao_usuario, wb_in, fabricante, bytes_cid, filename_rele, rele_escolhido, de_para_cod, de_para_campos, driver_rele, conf_colunas=None):
    doi_por_codigo, dois_por_ln, ns_or_dados_sel = driver_rele.carregar_equipamento(bytes_cid, filename_rele, rele_escolhido)
    
    wb_out = openpyxl.Workbook()
    aba_padrao = wb_out.active
    resultados_web = []
    processou_algo = False
    
    fill_cab = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    font_cab = Font(color="FFFFFF", bold=True)
    fill_vermelho = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    font_vermelho = Font(color="9C0006", bold=True)
    fill_relatorio = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    borda = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                   top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

    for index, row in df_selecao_usuario.iterrows():
        if not row["Processar"]: continue
        
        nome_aba = row["Aba"]
        grupos_ativos = [g for g in [1, 2, 3] if row.get(f"Grupo {g}")]
        if not grupos_ativos or nome_aba not in wb_in.sheetnames: continue

        ws_in = wb_in[nome_aba]
        itens_oa = extrair_parametros_visiveis(ws_in, fabricante, de_para_cod, de_para_campos, dois_por_ln, conf_colunas)
        
        if not itens_oa: continue
        processou_algo = True
        
        ws_out = wb_out.create_sheet(title=f"CONF_{nome_aba}"[:31])
        linha_inicio = 8
        titulos = ["CÓDIGO (Referência)", "DESCRIÇÃO DOS PARÂMETROS"]
        for g in grupos_ativos:
            titulos.extend([f"Ajuste OA (G{g})", f"Equipamento Relé (G{g})"])
        
        for col_idx, t in enumerate(titulos, start=1):
            c = ws_out.cell(linha_inicio, col_idx, t)
            c.fill, c.font, c.alignment = fill_cab, font_cab, Alignment(horizontal="center", vertical="center")

        total = len(itens_oa)
        corretos = {g: 0 for g in grupos_ativos}
        diferentes = {g: 0 for g in grupos_ativos}
        linhas_tabela_web = []

        for i, item in enumerate(itens_oa, start=linha_inicio + 1):
            cod_alvo = item["cod_cid"]
            if cod_alvo == "SEM_DE_PARA" and item["col_a"]:
                cod_alvo = item["col_a"]

            if fabricante == "SEL (Schweitzer)":
                doi = cod_alvo
            else:
                doi = doi_por_codigo.get(cod_alvo) or (doi_por_codigo.get(cod_alvo.split(".", 1)[-1]) if "." in cod_alvo else None)
            
            codigo_exibicao = str(item["col_a"]).strip() if str(item["col_a"]).strip() else str(cod_alvo)
            if codigo_exibicao == "SEM_DE_PARA": codigo_exibicao = "-"
            
            ws_out.cell(i, 1, codigo_exibicao).border = borda
            ws_out.cell(i, 2, item["col_b"]).border = borda

            linha_web = {"Código": codigo_exibicao, "Descrição": str(item["col_b"])}

            col_excel = 3
            for tab_idx in grupos_ativos:
                val_oa = item.get(f"val_g{tab_idx}")
                val_cid = driver_rele.obter_valor(doi, ns_or_dados_sel, str(tab_idx))
                
                cel_oa = ws_out.cell(i, col_excel)
                cel_rele = ws_out.cell(i, col_excel + 1)
                
                cel_oa.border = borda
                cel_rele.border = borda
                
                cel_oa.value = "" if val_oa is None else val_oa
                cel_rele.value = val_cid
                
                linha_web[f"OA G{tab_idx}"] = "" if val_oa is None else str(val_oa)
                linha_web[f"Relé G{tab_idx}"] = str(val_cid)
                
                if valores_sao_iguais(val_oa, val_cid):
                    corretos[tab_idx] += 1
                else:
                    diferentes[tab_idx] += 1
                    cel_rele.fill, cel_rele.font = fill_vermelho, font_vermelho
                
                col_excel += 2
            
            linhas_tabela_web.append(linha_web)

        grupos_txt = ", ".join([f"Grupo {g}" for g in grupos_ativos])
        ws_out.cell(1, 1, f"RELATÓRIO GERAL - ABA '{nome_aba}' vs RELÉ ({grupos_txt})").font = Font(bold=True, size=12)
        
        resumo = [(3, "Parâmetros Processados", [total] * len(grupos_ativos))]
        resumo.append((4, "Parâmetros Corretos", [corretos[g] for g in grupos_ativos]))
        resumo.append((5, "Parâmetros Diferentes (Em Vermelho)", [diferentes[g] for g in grupos_ativos]))
        
        ws_out.cell(2, 2, "Indicador").font = Font(bold=True)
        col_resumo = 4
        for g in grupos_ativos:
            ws_out.cell(2, col_resumo, f"Resultados Grupo {g}").font = Font(bold=True)
            col_resumo += 2

        for r_idx, lbl, vals in resumo:
            ws_out.cell(r_idx, 2, lbl).fill = fill_relatorio
            col_val = 4
            for v in vals:
                c = ws_out.cell(r_idx, col_val, v)
                c.fill, c.font, c.alignment = fill_relatorio, Font(bold=True), Alignment(horizontal="center")
                col_val += 2

        ws_out.column_dimensions['A'].width = 30
        ws_out.column_dimensions['B'].width = 58
        for col_idx in range(3, 3 + (len(grupos_ativos) * 2)):
            ws_out.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 26

        resultados_web.append({"nome": nome_aba, "total": total, "corretos": corretos, "diferentes": diferentes, "df": pd.DataFrame(linhas_tabela_web), "grupos_ativos": grupos_ativos})

    if processou_algo:
        wb_out.remove(aba_padrao)
        excel_io = io.BytesIO()
        wb_out.save(excel_io)
        excel_io.seek(0)
        return True, excel_io, resultados_web
    return False, None, []

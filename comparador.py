import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import pandas as pd
import io
from leitor_excel import extrair_parametros_visiveis

def normalizar_para_comparacao(valor):
    if valor is None: return ""
    s = str(valor).strip()
    if s.upper() in ["YES", "1", "TRUE"]: return "YES"
    if s.upper() in ["NO", "0", "FALSE"]: return "NO"
    if s.endswith(" A") and s[:-2].strip().replace('.', '', 1).isdigit():
        s = s[:-2].strip()
    if s.upper() == "IEC DEFINE TIME": return "IEC DEFINITE TIME"
    try:
        f = float(s.replace(',', '.'))
        return str(int(f)) if f.is_integer() else f"{f:.6f}".rstrip('0').rstrip('.')
    except ValueError:
        return s.upper()

def style_dataframe_resultados(row):
    styles = [''] * len(row)
    for i, col in enumerate(row.index):
        if col.startswith('Grupo'):
            val_cid = row[col]
            val_oa = row['Ajuste OA']
            if val_cid == "-":
                styles[i] = 'color: gray; text-align: center;'
            elif normalizar_para_comparacao(val_oa) == normalizar_para_comparacao(val_cid):
                styles[i] = 'color: #006100; font-weight: bold;' # Verde OK
            else:
                styles[i] = 'background-color: #FFC7CE; color: #9C0006; font-weight: bold;' # Vermelho Diferente
    return styles

def executar_comparacao(df_selecao_usuario, wb_in, fabricante, bytes_cid, de_para_cod, de_para_campos, driver_rele):
    doi_por_codigo, dois_por_ln, ns = driver_rele.carregar_equipamento(bytes_cid)
    
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
        itens_oa = extrair_parametros_visiveis(ws_in, de_para_cod, de_para_campos, dois_por_ln)
        
        if not itens_oa: continue
        processou_algo = True
        
        ws_out = wb_out.create_sheet(title=f"CONF_{nome_aba}"[:31])
        linha_inicio = 8
        titulos = ["CÓDIGO (Referência CID)", "DESCRIÇÃO DOS PARÂMETROS", f"AJUSTE OA ({nome_aba})", "CID - Table 1 (G1)", "CID - Table 2 (G2)", "CID - Table 3 (G3)"]
        
        for col_idx, t in enumerate(titulos, start=1):
            c = ws_out.cell(linha_inicio, col_idx, t)
            c.fill, c.font, c.alignment = fill_cab, font_cab, Alignment(horizontal="center", vertical="center")

        total = len(itens_oa)
        corretos = {1: 0, 2: 0, 3: 0}
        diferentes = {1: 0, 2: 0, 3: 0}
        linhas_tabela_web = []

        for i, item in enumerate(itens_oa, start=linha_inicio + 1):
            cod_alvo = item["cod_cid"]
            doi = doi_por_codigo.get(cod_alvo) or (doi_por_codigo.get(cod_alvo.split(".", 1)[-1]) if "." in cod_alvo else None)
            val_oa = item["col_c"]
            
            # --- NOVIDADE AQUI ---
            # Se a coluna de código no Excel estiver vazia (Modelo sem código), ele exibe o código CID que o sistema encontrou!
            codigo_exibicao = str(item["col_a"]).strip() if str(item["col_a"]).strip() else str(item["cod_cid"])
            if codigo_exibicao == "SEM_DE_PARA":
                codigo_exibicao = "-" # Estético: se não achar o mapeamento, coloca um traço em vez da palavra
            
            ws_out.cell(i, 1, codigo_exibicao).border = borda
            ws_out.cell(i, 2, item["col_b"]).border = borda
            ws_out.cell(i, 3, "" if val_oa is None else val_oa).border = borda

            linha_web = {
                "Código": codigo_exibicao, 
                "Descrição": str(item["col_b"]), 
                "Ajuste OA": "" if val_oa is None else str(val_oa)
            }

            for tab_idx, col_excel in [(1, 4), (2, 5), (3, 6)]:
                cel = ws_out.cell(i, col_excel)
                cel.border = borda
                if tab_idx in grupos_ativos:
                    val_cid = driver_rele.obter_valor(doi, ns, str(tab_idx))
                    cel.value = val_cid
                    linha_web[f"Grupo {tab_idx}"] = str(val_cid)
                    
                    if normalizar_para_comparacao(val_oa) == normalizar_para_comparacao(val_cid):
                        corretos[tab_idx] += 1
                    else:
                        diferentes[tab_idx] += 1
                        cel.fill, cel.font = fill_vermelho, font_vermelho
                else:
                    cel.value = "-"
                    cel.alignment = Alignment(horizontal="center")
                    linha_web[f"Grupo {tab_idx}"] = "-"
            
            linhas_tabela_web.append(linha_web)

        grupos_txt = ", ".join([f"Grupo {g}" for g in grupos_ativos])
        ws_out.cell(1, 1, f"RELATÓRIO GERAL - ABA '{nome_aba}' vs CID ({grupos_txt})").font = Font(bold=True, size=12)
        ws_out.cell(2, 2, "Indicador").font = Font(bold=True)
        ws_out.cell(2, 4, "Table 1 (G1)").font = Font(bold=True)
        ws_out.cell(2, 5, "Table 2 (G2)").font = Font(bold=True)
        ws_out.cell(2, 6, "Table 3 (G3)").font = Font(bold=True)

        resumo = [
            (3, "Parâmetros Processados", [total if g in grupos_ativos else "-" for g in [1, 2, 3]]),
            (4, "Parâmetros Corretos", [corretos[g] if g in grupos_ativos else "-" for g in [1, 2, 3]]),
            (5, "Parâmetros Diferentes (Em Vermelho)", [diferentes[g] if g in grupos_ativos else "-" for g in [1, 2, 3]])
        ]
        for r_idx, lbl, vals in resumo:
            ws_out.cell(r_idx, 2, lbl).fill = fill_relatorio
            for idx_v, v in enumerate(vals, start=4):
                c = ws_out.cell(r_idx, idx_v, v)
                c.fill, c.font, c.alignment = fill_relatorio, Font(bold=True), Alignment(horizontal="center")

        for col_idx, w in {1: 30, 2: 58, 3: 26, 4: 26, 5: 26, 6: 26}.items():
            ws_out.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = w

        resultados_web.append({
            "nome": nome_aba, 
            "total": total, 
            "corretos": corretos, 
            "diferentes": diferentes, 
            "df": pd.DataFrame(linhas_tabela_web)
        })

    if processou_algo:
        wb_out.remove(aba_padrao)
        excel_io = io.BytesIO()
        wb_out.save(excel_io)
        excel_io.seek(0)
        return True, excel_io, resultados_web
    return False, None, []
import re
from mapeamentos import limpar_espacos, chave_normalizada, LN_MAP_PADRAO

def obter_abas_validas(wb):
    abas_validas = []
    for sheet in wb.worksheets:
        if sheet.sheet_state == 'visible':
            nome = sheet.title.upper()
            if re.search(r'GRUPO\s*[123456]|GROUP\s*[123456]', nome) or "AJUSTE" in nome or "SETTING" in nome:
                abas_validas.append(sheet.title)
    return abas_validas

def extrair_parametros_visiveis(ws, fabricante, de_para_cod, de_para_campos, dois_por_ln, conf_colunas=None):
    itens_oa = []
    
    if conf_colunas is None:
        conf_colunas = {"codigo": 2, "desc": 3, "ajuste_g1": 6, "ajuste_g2": 7, "ajuste_g3": 8}
        
    idx_cod = conf_colunas["codigo"]
    idx_desc = conf_colunas["desc"]
    idx_g1 = conf_colunas["ajuste_g1"]
    idx_g2 = conf_colunas["ajuste_g2"]
    idx_g3 = conf_colunas["ajuste_g3"]
    
    if fabricante == "SEL (Schweitzer)":
        for r in range(12, ws.max_row + 1):
            if ws.row_dimensions[r].hidden: continue
            
            b_val = limpar_espacos(ws.cell(r, idx_cod).value)
            c_val = limpar_espacos(ws.cell(r, idx_desc).value)
            
            v_g1 = ws.cell(r, idx_g1).value
            v_g2 = ws.cell(r, idx_g2).value
            v_g3 = ws.cell(r, idx_g3).value
            
            if not b_val: continue
            if b_val.upper() in ["PARAMETRO", "PARÂMETRO", "PARAMETER"]: continue
            if b_val.upper() == "ELABORAÇÃO E IMPLANTAÇÃO": break
            if b_val.startswith(">>") or "CONFIGURATION" in b_val.upper(): continue
            if b_val.upper().startswith("MOTIVO"): continue
            
            # Filtra linhas vazias (sem descrição e sem valor de ajuste)
            if b_val and not c_val and v_g1 is None and v_g2 is None and v_g3 is None: continue
            
            itens_oa.append({
                "col_a": b_val,     
                "col_b": c_val,     
                "val_g1": v_g1,
                "val_g2": v_g2,
                "val_g3": v_g3,
                "cod_cid": b_val    
            })
        return itens_oa

    eh_modelo_1 = False
    for r in range(12, min(45, ws.max_row + 1)):
        if ws.row_dimensions[r].hidden: continue
        b_v = limpar_espacos(ws.cell(r, idx_cod).value)
        d_v = limpar_espacos(ws.cell(r, idx_desc + 1).value)
        if "." in b_v and d_v and not b_v.upper().startswith("MOTIVO"):
            eh_modelo_1 = True
            break

    start_row = 14 if eh_modelo_1 else 12
    secao, subgrupo = "", ""

    for r in range(start_row, ws.max_row + 1):
        if ws.row_dimensions[r].hidden: continue
        
        b_val = limpar_espacos(ws.cell(r, idx_cod).value)
        c_val = limpar_espacos(ws.cell(r, idx_desc).value)
        
        if b_val.upper() in ["PARAMETRO", "PARÂMETRO"]: continue
        if b_val.upper() == "ELABORAÇÃO E IMPLANTAÇÃO": break
        
        if eh_modelo_1:
            d_val = limpar_espacos(ws.cell(r, idx_desc + 1).value) 
            v_g1 = ws.cell(r, idx_g1).value
            v_g2 = ws.cell(r, idx_g2).value
            v_g3 = ws.cell(r, idx_g3).value
            
            if b_val and not d_val and v_g1 is None and v_g2 is None and v_g3 is None and not "." in b_val:
                secao, subgrupo = b_val, ""
                continue
                
            if "." in b_val and d_val:
                if c_val: subgrupo = c_val
                cod_cid = de_para_cod.get(chave_normalizada(b_val))
                if not cod_cid:
                    cod_cid = de_para_campos.get((chave_normalizada(secao), chave_normalizada(subgrupo), chave_normalizada(d_val)), b_val)
                
                desc_completa = f"{subgrupo} - {d_val}" if subgrupo else d_val
                itens_oa.append({"col_a": b_val, "col_b": desc_completa, "val_g1": v_g1, "val_g2": v_g2, "val_g3": v_g3, "cod_cid": cod_cid})
        else:
            v_g1 = ws.cell(r, idx_g1).value
            v_g2 = ws.cell(r, idx_g2).value
            v_g3 = ws.cell(r, idx_g3).value
            
            if b_val and not c_val:
                secao, subgrupo = b_val, ""
                continue
                
            if c_val:
                if b_val: subgrupo = b_val
                chave = (chave_normalizada(secao), chave_normalizada(subgrupo), chave_normalizada(c_val))
                cod_cid = de_para_campos.get(chave)
                
                if not cod_cid:
                    full_ln = LN_MAP_PADRAO.get((secao, subgrupo))
                    if full_ln and chave_normalizada(c_val) in dois_por_ln.get(full_ln, {}):
                        doi_name = dois_por_ln[full_ln][chave_normalizada(c_val)][0]
                        cod_cid = f"{full_ln}.{doi_name}"
                        
                desc_completa = f"{subgrupo} - {c_val}" if subgrupo else c_val
                itens_oa.append({"col_a": "", "col_b": desc_completa, "val_g1": v_g1, "val_g2": v_g2, "val_g3": v_g3, "cod_cid": cod_cid if cod_cid else "SEM_DE_PARA"})

    return itens_oa

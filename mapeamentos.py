import os
import openpyxl

LN_MAP_PADRAO = {
    ("CTS VTS CONFIG.", "CTs Ratios"): "PROT.TCTR1",
    ("CTS VTS CONFIG.", "Side A VTs Characteristics"): "PROT.TVTR1",
    ("CTS VTS CONFIG.", "Side B VTs Characteristics"): "PROT.TVTR2",
    ("RATED CURRENTS", "Rated currents"): "PROT.TCIN1",
    ("UNIQUE GROUP PROTECTION", "Voltage Presence/Absence"): "PROT.RVPA1",
    ("UNIQUE GROUP PROTECTION", "Loop Aut/Self Healing"): "PROT.RLAU1",
    ("UNIQUE GROUP PROTECTION", "Hot Line Tag (General Config.)"): "PROT.RHLT1",
    ("PROTECTION/OVERCURRENT/TOC PHASES", "Unit 1 (51-1)"): "PROT.PTOC1",
    ("PROTECTION/OVERCURRENT/TOC PHASES", "Unit 2 (51-2)"): "PROT.PTOC2",
    ("PROTECTION/OVERCURRENT/IOC PHASES", "Unit 1 (50-1)"): "PROT.PIOC1",
    ("PROTECTION/OVERCURRENT/IOC PHASES", "Unit 2 (50-2)"): "PROT.PIOC2",
    ("PROTECTION/OVERCURRENT/IOC PHASES", "Unit 3 (50-3)"): "PROT.PIOC3",
    ("PROTECTION/OVERCURRENT/TOC Neutral", "Unit 1 (51N-1)"): "PROT.GPTOC1",
    ("PROTECTION/OVERCURRENT/TOC Neutral", "Unit 2 (51N-2)"): "PROT.GPTOC2",
    ("PROTECTION/OVERCURRENT/IOC Neutral", "Unit 1 (50N-1)"): "PROT.GPIOC1",
    ("PROTECTION/OVERCURRENT/IOC Neutral", "Unit 2 (50N-2)"): "PROT.GPIOC2",
    ("PROTECTION/OVERCURRENT/IOC Neutral", "Unit 3 (50N-3)"): "PROT.GPIOC3",
    ("PROTECTION/OVERCURRENT/TOC SENSITIVE NEUTRAL", "Unit 1 (51NS-1)"): "PROT.SGPTOC1",
    ("PROTECTION/OVERCURRENT/TOC SENSITIVE NEUTRAL", "Unit 1 (51NS1-1)"): "PROT.SGPTOC1",
    ("PROTECTION/OVERCURRENT/TOC SENSITIVE NEUTRAL", "Unit 2 (51NS-2)"): "PROT.SGPTOC2",
    ("PROTECTION/OVERCURRENT/IOC Sensitive Neutral", "Unit 1 (50NS-1)"): "PROT.SGPIOC1",
    ("PROTECTION/OVERCURRENT/IOC Sensitive Neutral", "Unit 2 (50NS-2)"): "PROT.SGPIOC2",
    ("PROTECTION/OVERCURRENT/IOC Sensitive Neutral", "Unit 3 (50NS-3)"): "PROT.SGPIOC3",
    ("PROTECTION/OVERCURRENT/BROKEN CONDUCTOR", "Unit 1 (50NS-1)"): "PROT.OPPTOC1",
    ("PROTECTION/VOLTAGE/PHASES OVERVOLTAGE", "Instantaneous Unit 1 (59-1)"): "PROT.PIOV1",
    ("PROTECTION/VOLTAGE/PHASES OVERVOLTAGE", "Instantaneous Unit 2 (59-2)"): "PROT.PIOV3",
    ("PROTECTION/VOLTAGE/IOV PHASES OVERVOLTAGE (59I)", "Instantaneous Unit 1 (59-1)"): "PROT.PIOV1",
    ("PROTECTION/VOLTAGE/IOV PHASES OVERVOLTAGE (59I)", "Instantaneous Unit 2 (59-2)"): "PROT.PIOV3",
    ("PROTECTION/VOLTAGE/V0 OVERVOLTAGE", "Instantaneous Unit 1 (59N)"): "PROT.GPIOV1",
    ("PROTECTION/VOLTAGE/IOV OVERVOLTAGE 3V0 (59NI)", "Instantaneous Unit 1 (59N)"): "PROT.GPIOV1",
    ("PROTECTION/VOLTAGE/PHASES UNDERVOLTAGE", "Instantaneous Unit 1 (27-1)"): "PROT.PIUV1",
    ("PROTECTION/VOLTAGE/PHASES UNDERVOLTAGE", "Instantaneous Unit 2 (27-2)"): "PROT.PIUV3",
    ("PROTECTION/VOLTAGE/IUV UNDERVOLTAGE (27I)", "Instantaneous Unit 1 (27-1)"): "PROT.PIUV1",
    ("PROTECTION/VOLTAGE/IUV UNDERVOLTAGE (27I)", "Instantaneous Unit 2 (27-2)"): "PROT.PIUV3",
    ("PROTECTION/FREQUENCY", "Under/Over freq. (81O/U)"): "PROT.PTGF1",
    ("PROTECTION/DIRECTIONAL/DIRECTIONAL PHASES", "Unit 1"): "PROT.RDIR1",
    ("PROTECTION/DIRECTIONAL/DIRECTIONAL NEUTRAL", "Ground Directional (67N)"): "PROT.G67NRDIR1",
    ("PROTECTION/SECOND HARMONIC RESTRAINT", "Phase - Unit 1 (Proteção Inrush)"): "PROT.PHAR1",
    ("PROTECTION/HOT LINE TAG/TOC PHASES", "Unit 1 (51-HLT)"): "PROT.HLTPTOC1",
    ("PROTECTION/HOT LINE TAG/TOC NEUTRAL", "Unit 1 (51N-HLT)"): "PROT.GHLTPTOC1",
    ("PROTECTION/HIGH CURRENT LOCKOUT", "Phases"): "PROT.PHCL1",
    ("PROTECTION/HIGH CURRENT LOCKOUT", "Neutral"): "PROT.GPHCL1",
    ("PROTECTION/AUTOMATISM/RECLOSER", "Autoreclose (79)"): "PROT.RREC1",
    ("LPHD INFO", "LPHD Info"): "GEN.LPHD1",
    ("LPHD INFO", ""): "GEN.LPHD1",
}

VARIACOES_DESCRICAO = {
    ("PROT.RLAU1", "AutoRestore ENBL"): "RestoreEna",
    ("PROT.RLAU1", "AutoRestore Enabled"): "RestoreEna",
    ("PROT.RLAU1", "Preferred side"): "PrefSide",
    ("PROT.RHLT1", "Enable"): "RHLTEna",
    ("PROT.PIOC3", "Delay time(ms)"): "OpDlTmms",
    ("PROT.GPIOC3", "Delay time(ms)"): "OpDlTmms",
    ("PROT.GPIOC3", "ENBL"): "PIOCEna",
    ("PROT.SGPIOC3", "Delay time(ms)"): "OpDlTmms",
    ("PROT.SGPTOC2", "Trip Permission by recloser"): "TripPerm",
    ("PROT.OPPTOC1", "Delay time(ms)"): "OpDlTmms",
    ("PROT.OPPTOC1", "Min. Phase I ((%)Irated)"): "Blklph",
    ("PROT.OPPTOC1", "Máx. Threshold I0/Irated (%)"): "BLkI0I1",
    ("PROT.PIOV1", "Operate delay (ms)"): "OpDlTmms",
    ("PROT.PIOV1", "Enabled Events Record"): "MaskEna",
    ("PROT.PIOV3", "Operate delay (ms)"): "OpDlTmms",
    ("PROT.PIOV3", "Enabled Events Record"): "MaskEna",
    ("PROT.GPIOV1", "Operate delay (ms)"): "OpDlTmms",
    ("PROT.GPIOV1", "Enabled Events Record"): "MaskEna",
    ("PROT.PIUV1", "Operate delay (ms)"): "OpDlTmms",
    ("PROT.PIUV1", "Enabled Events Record"): "MaskEna",
    ("PROT.PIUV3", "Operate delay (ms)"): "OpDlTmms",
    ("PROT.PIUV3", "Enabled Events Record"): "MaskEna",
    ("PROT.PTGF1", "Step 1. Enabled (81O - Estágio 1)"): "PTGFEna1",
    ("PROT.PTGF1", "Step 1. Star value (Hz) (81O - Estágio 1)"): "StrVal1",
    ("PROT.PTGF1", "Step 1. Delay time (ms) (81O - Estágio 1)"): "OpDlTmms1",
    ("PROT.PTGF1", "Step 2. Enabled (81O - Estágio 2)"): "PTGFEna2",
    ("PROT.PTGF1", "Step 2. Star value (Hz) (81O - Estágio 2)"): "StrVal2",
    ("PROT.PTGF1", "Step 2. Delay time (ms) (81O - Estágio 2)"): "OpDlTmms2",
    ("PROT.PTGF1", "Step 3. Enabled (81U - Estágio 1)"): "PTGFEna3",
    ("PROT.PTGF1", "Step 3. Star value (Hz) (81U - Estágio 1)"): "StrVal3",
    ("PROT.PTGF1", "Step 3. Delay time (ms) (81U - Estágio 1)"): "OpDlTmms3",
    ("PROT.PTGF1", "Step 4. Enabled (81U - Estágio 2)"): "PTGFEna4",
    ("PROT.PTGF1", "Step 4. Star value (Hz) (81U - Estágio 2)"): "StrVal4",
    ("PROT.PTGF1", "Step 4. Delay time (ms) (81U - Estágio 2)"): "OpDlTmms4",
    ("PROT.PTGF1", "Enabled Events Record"): "MaskEna",
    ("PROT.RDIR1", "Enabled Events Record"): "MaskEna",
    ("PROT.G67NRDIR1", "Directional method"): "DirTyp",
    ("PROT.G67NRDIR1", "Torque angle (degree)"): "ChrAng",
    ("PROT.G67NRDIR1", "Enabled Events Record"): "MaskEna",
    ("PROT.PHAR1", "I2/Ifun. Threshold (%)"): "StrVal",
    ("PROT.PHCL1", "79 Blocking Trip number"): "TrNum",
    ("PROT.GPHCL1", "79 Blocking Trip number"): "TrNum",
}

def limpar_espacos(texto):
    if texto is None: return ""
    return " ".join(str(texto).strip().split())

def chave_normalizada(texto):
    return "".join(limpar_espacos(texto).lower().split())

def carregar_biblioteca_prioritaria(file_bib_upload, fabricante, dois_por_ln):
    if file_bib_upload is not None:
        wb = openpyxl.load_workbook(file_bib_upload, data_only=True)
        return _extrair_de_para(wb, fabricante)
    
    caminho_local = "biblioteca_de_para.xlsx"
    if os.path.exists(caminho_local):
        wb = openpyxl.load_workbook(caminho_local, data_only=True)
        return _extrair_de_para(wb, fabricante)
    
    de_para_cod = {}
    de_para_campos = {}
    for (secao, subgrupo), full_ln in LN_MAP_PADRAO.items():
        curto_ln = full_ln.split(".")[-1] if full_ln.startswith("PROT.") else full_ln
        for _, (doi_name, doi_desc) in dois_por_ln.get(full_ln, {}).items():
            cod_excel = f"{curto_ln}.{doi_name}"
            cod_cid = f"{full_ln}.{doi_name}"
            de_para_cod[chave_normalizada(cod_excel)] = cod_cid
            de_para_campos[(chave_normalizada(secao), chave_normalizada(subgrupo), chave_normalizada(doi_desc))] = cod_cid
            
        for (ovr_ln, ovr_desc), doi_name in VARIACOES_DESCRICAO.items():
            if ovr_ln == full_ln:
                cod_excel = f"{curto_ln}.{doi_name}"
                cod_cid = f"{full_ln}.{doi_name}"
                de_para_cod[chave_normalizada(cod_excel)] = cod_cid
                de_para_campos[(chave_normalizada(secao), chave_normalizada(subgrupo), chave_normalizada(ovr_desc))] = cod_cid

    return de_para_cod, de_para_campos

def _extrair_de_para(wb, fabricante):
    ws = wb[fabricante] if fabricante in wb.sheetnames else wb.active
    de_para_cod, de_para_campos = {}, {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 5 or not row[4]: continue
        cod_xls, sec, sub, desc, cod_cid = [str(c).strip() for c in row[:5] if c is not None]
        if cod_xls: de_para_cod[chave_normalizada(cod_xls)] = cod_cid
        if desc: de_para_campos[(chave_normalizada(sec), chave_normalizada(sub), chave_normalizada(desc))] = cod_cid
    return de_para_cod, de_para_campos
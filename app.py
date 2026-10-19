import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import xml.etree.ElementTree as ET
import pandas as pd
import re
import os
import io

# ==============================================================================
# 1. MAPEAMENTOS PADRÃO (ARTECHE ADATECH RC / INGETEAM PT5)
# ==============================================================================
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

# ==============================================================================
# FUNÇÕES AUXILIARES
# ==============================================================================
def limpar_espacos(texto):
    if texto is None: return ""
    return " ".join(str(texto).strip().split())

def chave_normalizada(texto):
    return "".join(limpar_espacos(texto).lower().split())

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

class DriverArtecheAdaTech:
    @staticmethod
    def carregar_equipamento(bytes_cid):
        tree = ET.parse(io.BytesIO(bytes_cid))
        root = tree.getroot()
        ns = {'scl': 'http://www.iec.ch/61850/2003/SCL'}
        doi_por_codigo = {}
        dois_por_ln = {}

        def eh_doi_ajuste(doi):
            if doi.find('scl:SDI[@name="setMag"]', ns) is not None: return True
            for dai_name in ['setVal', 'setCharact', 'setMag']:
                if doi.find(f'scl:DAI[@name="{dai_name}"]', ns) is not None: return True
            for dai in doi.findall('.//scl:DAI', ns):
                if dai.attrib.get('sAddr', '').startswith('A,'): return True
            return False

        for ld in root.findall('.//scl:LDevice', ns):
            ld_inst = ld.attrib.get('inst', '')
            for ln in ld.findall('scl:LN0', ns) + ld.findall('scl:LN', ns):
                ln_name = ln.attrib.get('prefix', '') + ln.attrib.get('lnClass', '') + ln.attrib.get('inst', '')
                full_ln = f"{ld_inst}.{ln_name}"
                dois_por_ln.setdefault(full_ln, {})
                for doi in ln.findall('scl:DOI', ns):
                    if not eh_doi_ajuste(doi): continue
                    doi_name = doi.attrib.get('name', '')
                    doi_desc = doi.attrib.get('desc', '')
                    cod_completo = f"{ld_inst}.{ln_name}.{doi_name}"
                    cod_curto = f"{ln_name}.{doi_name}"
                    doi_por_codigo[cod_completo] = doi
                    if cod_curto not in doi_por_codigo or ld_inst == 'PROT':
                        doi_por_codigo[cod_curto] = doi
                    if doi_desc:
                        dois_por_ln[full_ln][chave_normalizada(doi_desc)] = (doi_name, doi_desc)
        return doi_por_codigo, dois_por_ln, ns

    @staticmethod
    def obter_valor(doi, ns, sgroup="1"):
        if doi is None: return "NÃO ENCONTRADO"
        sdi_setmag = doi.find('scl:SDI[@name="setMag"]', ns)
        if sdi_setmag is not None:
            for dai in sdi_setmag.findall('scl:DAI', ns):
                vals = dai.findall('scl:Val', ns)
                if vals:
                    for v in vals:
                        if v.attrib.get('sGroup') == str(sgroup): return v.text
                    return vals[0].text
        for dai_name in ['setVal', 'setCharact', 'setMag']:
            dai = doi.find(f'scl:DAI[@name="{dai_name}"]', ns)
            if dai is not None:
                vals = dai.findall('scl:Val', ns)
                if vals:
                    for v in vals:
                        if v.attrib.get('sGroup') == str(sgroup): return v.text
                    return vals[0].text
        for dai in doi.findall('.//scl:DAI', ns):
            if dai.attrib.get('sAddr', '').startswith('A,'):
                vals = dai.findall('scl:Val', ns)
                if vals:
                    for v in vals:
                        if v.attrib.get('sGroup') == str(sgroup): return v.text
                    return vals[0].text
        return "VAZIO"

def gerar_biblioteca_memoria(dois_por_ln, fabricante="ARTECHE_ADATECH"):
    wb_bib = openpyxl.Workbook()
    ws_bib = wb_bib.active
    ws_bib.title = "ARTECHE_ADATECH"
    ws_bib.append(["Codigo_Excel", "Secao", "Subgrupo", "Descricao_Excel", "Codigo_CID"])

    registros = set()
    for (secao, subgrupo), full_ln in LN_MAP_PADRAO.items():
        curto_ln = full_ln.split(".")[-1] if full_ln.startswith("PROT.") else full_ln
        for _, (doi_name, doi_desc) in dois_por_ln.get(full_ln, {}).items():
            cod_excel = f"{curto_ln}.{doi_name}"
            cod_cid = f"{full_ln}.{doi_name}"
            chave = (cod_excel, secao, subgrupo, limpar_espacos(doi_desc))
            if chave not in registros:
                ws_bib.append([cod_excel, secao, subgrupo, limpar_espacos(doi_desc), cod_cid])
                registros.add(chave)
        for (ovr_ln, ovr_desc), doi_name in VARIACOES_DESCRICAO.items():
            if ovr_ln == full_ln:
                cod_excel = f"{curto_ln}.{doi_name}"
                cod_cid = f"{full_ln}.{doi_name}"
                chave = (cod_excel, secao, subgrupo, limpar_espacos(ovr_desc))
                if chave not in registros:
                    ws_bib.append([cod_excel, secao, subgrupo, limpar_espacos(ovr_desc), cod_cid])
                    registros.add(chave)

    de_para_cod = {}
    de_para_campos = {}
    for row in ws_bib.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 5 or not row[4]: continue
        cod_xls, sec, sub, desc, cod_cid = [limpar_espacos(c) for c in row[:5]]
        if cod_xls: de_para_cod[chave_normalizada(cod_xls)] = cod_cid
        if desc: de_para_campos[(chave_normalizada(sec), chave_normalizada(sub), chave_normalizada(desc))] = cod_cid
        
    bib_io = io.BytesIO()
    wb_bib.save(bib_io)
    bib_io.seek(0)
    return de_para_cod, de_para_campos, bib_io

def ler_biblioteca_anexada(file_bib, fabricante):
    wb_bib = openpyxl.load_workbook(file_bib, data_only=True)
    ws_bib = wb_bib[fabricante] if fabricante in wb_bib.sheetnames else wb_bib.active
    de_para_cod = {}
    de_para_campos = {}
    for row in ws_bib.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 5 or not row[4]: continue
        cod_xls, sec, sub, desc, cod_cid = [limpar_espacos(c) for c in row[:5]]
        if cod_xls: de_para_cod[chave_normalizada(cod_xls)] = cod_cid
        if desc: de_para_campos[(chave_normalizada(sec), chave_normalizada(sub), chave_normalizada(desc))] = cod_cid
    return de_para_cod, de_para_campos

def extrair_parametros_da_aba(ws, de_para_cod, de_para_campos, dois_por_ln):
    eh_modelo_1 = False
    for r in range(12, min(45, ws.max_row + 1)):
        b_v = limpar_espacos(ws.cell(r, 2).value)
        d_v = limpar_espacos(ws.cell(r, 4).value)
        if "." in b_v and d_v and not b_v.upper().startswith("MOTIVO"):
            eh_modelo_1 = True
            break

    itens_oa = []
    secao, subgrupo = "", ""

    if eh_modelo_1:
        for r in range(14, ws.max_row + 1):
            b_val = limpar_espacos(ws.cell(r, 2).value)
            c_val = limpar_espacos(ws.cell(r, 3).value)
            d_val = limpar_espacos(ws.cell(r, 4).value)
            g_val = ws.cell(r, 7).value
            if b_val.upper() in ["PARAMETRO", "PARÂMETRO"]: continue
            if b_val.upper() == "ELABORAÇÃO E IMPLANTAÇÃO": break
            if b_val and not d_val and "." not in b_val:
                secao, subgrupo = b_val, ""
                continue
            if "." in b_val and d_val:
                if c_val: subgrupo = c_val
                cod_cid = de_para_cod.get(chave_normalizada(b_val))
                if not cod_cid: cod_cid = de_para_campos.get((chave_normalizada(secao), chave_normalizada(subgrupo), chave_normalizada(d_val)), b_val)
                itens_oa.append({"col_a": b_val, "col_b": f"{subgrupo} - {d_val}" if subgrupo else d_val, "col_c": g_val, "cod_cid": cod_cid})
    else:
        for r in range(12, ws.max_row + 1):
            b_val = limpar_espacos(ws.cell(r, 2).value)
            c_val = limpar_espacos(ws.cell(r, 3).value)
            l_val = ws.cell(r, 12).value
            r_val = ws.cell(r, 18).value
            if b_val.upper() in ["PARAMETRO", "PARÂMETRO"]: continue
            if b_val.upper() == "ELABORAÇÃO E IMPLANTAÇÃO": break
            if b_val and not c_val:
                secao, subgrupo = b_val, ""
                continue
            if c_val:
                if b_val: subgrupo = b_val
                val_oa = r_val if r_val is not None else (l_val if secao.upper() == "LPHD INFO" else None)
                chave = (chave_normalizada(secao), chave_normalizada(subgrupo), chave_normalizada(c_val))
                cod_cid = de_para_campos.get(chave)
                if not cod_cid:
                    full_ln = LN_MAP_PADRAO.get((secao, subgrupo))
                    if full_ln and chave_normalizada(c_val) in dois_por_ln.get(full_ln, {}):
                        cod_cid = f"{full_ln}.{dois_por_ln[full_ln][chave_normalizada(c_val)][0]}"
                itens_oa.append({"col_a": "", "col_b": f"{subgrupo} - {c_val}" if subgrupo else c_val, "col_c": val_oa, "cod_cid": cod_cid if cod_cid else "SEM_DE_PARA"})

    return itens_oa

# Formatação de cores para a Tabela Web
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

# ==============================================================================
# INTERFACE STREAMLIT
# ==============================================================================
st.set_page_config(page_title="Comparador de Ajustes - Proteção", layout="wide")
st.title("⚡ Comparador de Ajustes de Proteção (OA vs Equipamento)")

# Agora com 4 Colunas (incluindo o De-Para Opcional)
col1, col2, col3, col4 = st.columns(4)
with col1:
    fabricante = st.selectbox("1. Modelo do Equipamento:", ["ARTECHE_ADATECH", "SCHNEIDER_ADVC", "NOJA_RC10", "COOPER_FORM6"])
with col2:
    file_xlsx = st.file_uploader("2. Ordem de Ajuste (.xlsx):", type=["xlsx"])
with col3:
    file_cid = st.file_uploader("3. Arquivo do Equipamento (.cid):", type=["cid", "xml", "icd"])
with col4:
    file_bib = st.file_uploader("4. De-Para (.xlsx) [Opcional]:", type=["xlsx"], help="Se vazio, a biblioteca embutida será usada.")

if file_xlsx:
    wb_in = openpyxl.load_workbook(file_xlsx, data_only=True)
    nomes_abas = wb_in.sheetnames
    
    st.markdown("### 🔗 Associação de Abas da Planilha com os Grupos do CID")
    
    abas_grupo = [n for n in nomes_abas if re.search(r'GRUPO\s*[123456]|GROUP\s*[123456]', n.upper())]
    abas_ajuste = [n for n in nomes_abas if "AJUSTE" in n.upper() or "SETTING" in n.upper()]
    multi = len(abas_grupo) > 1

    dados_iniciais = []
    for nome in nomes_abas:
        n_up = nome.upper()
        proc = False
        g1 = g2 = g3 = False
        if multi and nome in abas_grupo:
            proc = True
            if "1" in n_up: g1 = True
            elif "2" in n_up: g2 = True
            elif "3" in n_up: g3 = True
            else: g1 = True
        elif not multi and (nome in abas_grupo or nome in abas_ajuste):
            proc = True
            g1 = g2 = g3 = True

        dados_iniciais.append({"Processar": proc, "Aba": nome, "Grupo 1": g1, "Grupo 2": g2, "Grupo 3": g3})

    df = pd.DataFrame(dados_iniciais)
    df_editado = st.data_editor(df, hide_index=True, column_config={
        "Processar": st.column_config.CheckboxColumn(required=True),
        "Aba": st.column_config.TextColumn(disabled=True)
    })

    if st.button("🔍 Executar Conferência de Ajustes", type="primary") and file_cid:
        if fabricante != "ARTECHE_ADATECH":
            st.error(f"O driver para {fabricante} será adicionado em breve.")
        else:
            with st.spinner("A processar a conferência..."):
                doi_por_codigo, dois_por_ln, ns = DriverArtecheAdaTech.carregar_equipamento(file_cid.read())
                
                # Lógica do De-Para Opcional
                bib_io = None
                if file_bib is not None:
                    de_para_cod, de_para_campos = ler_biblioteca_anexada(file_bib, fabricante)
                else:
                    de_para_cod, de_para_campos, bib_io = gerar_biblioteca_memoria(dois_por_ln, fabricante)

                wb_out = openpyxl.Workbook()
                aba_padrao = wb_out.active
                
                # Estilos do Excel
                fill_cab = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                font_cab = Font(color="FFFFFF", bold=True)
                fill_vermelho = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                font_vermelho = Font(color="9C0006", bold=True)
                fill_relatorio = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                borda = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                               top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

                processou_algo = False
                resultados_web = [] # Guardar resultados para mostrar na tela

                for index, row in df_editado.iterrows():
                    if not row["Processar"]: continue
                    nome_aba = row["Aba"]
                    grupos_ativos = []
                    if row["Grupo 1"]: grupos_ativos.append(1)
                    if row["Grupo 2"]: grupos_ativos.append(2)
                    if row["Grupo 3"]: grupos_ativos.append(3)
                    
                    if not grupos_ativos: continue

                    ws_in = wb_in[nome_aba]
                    itens_oa = extrair_parametros_da_aba(ws_in, de_para_cod, de_para_campos, dois_por_ln)
                    if not itens_oa: continue
                    processou_algo = True

                    ws_out = wb_out.create_sheet(title=f"CONF_{nome_aba}"[:31])
                    linha_inicio = 8
                    titulos = ["CÓDIGO", "DESCRIÇÃO DOS PARÂMETROS", f"AJUSTE OA ({nome_aba})", "CID - Table 1 (G1)", "CID - Table 2 (G2)", "CID - Table 3 (G3)"]
                    
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
                        
                        # Escreve no Excel
                        ws_out.cell(i, 1, item["col_a"]).border = borda
                        ws_out.cell(i, 2, item["col_b"]).border = borda
                        ws_out.cell(i, 3, "" if val_oa is None else val_oa).border = borda

                        # Prepara linha para a Tela (Web)
                        linha_web = {
                            "Código": str(item["col_a"]),
                            "Descrição": str(item["col_b"]),
                            "Ajuste OA": "" if val_oa is None else str(val_oa)
                        }

                        for tab_idx, col_excel in [(1, 4), (2, 5), (3, 6)]:
                            cel = ws_out.cell(i, col_excel)
                            cel.border = borda
                            if tab_idx in grupos_ativos:
                                val_cid = DriverArtecheAdaTech.obter_valor(doi, ns, str(tab_idx))
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

                    # Cabeçalho Relatório Excel
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

                    for col_idx, w in {1: 22, 2: 58, 3: 26, 4: 26, 5: 26, 6: 26}.items():
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
                    
                    st.success("✅ Conferência concluída com sucesso!")
                    
                    # Botões de Download
                    col_dl1, col_dl2 = st.columns(2)
                    col_dl1.download_button("📥 Baixar Planilha de Conferência (.xlsx)", excel_io, "Resultado_Conferencia.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    if bib_io is not None:
                        col_dl2.download_button("📚 Baixar Biblioteca De-Para Gerada (.xlsx)", bib_io, "biblioteca_de_para.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                    # ==== RENDERIZAR TABELAS E RESUMOS NA TELA ====
                    st.markdown("---")
                    st.subheader("📊 Relatório de Conferência")
                    
                    abas_nomes = [res["nome"] for res in resultados_web]
                    tabs = st.tabs(abas_nomes)
                    
                    for i, tab in enumerate(tabs):
                        res = resultados_web[i]
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
                else:
                    st.warning("⚠️ Nenhuma aba válida foi processada.")

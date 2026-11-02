import xml.etree.ElementTree as ET
import io
import zipfile
import olefile

def limpar_espacos(texto):
    if texto is None: return ""
    return " ".join(str(texto).strip().split())

def chave_normalizada(texto):
    return "".join(limpar_espacos(texto).lower().split())

class DriverArtecheAdaTech:
    @staticmethod
    def listar_reles(bytes_io, filename):
        return ["Arteche_Default"] # Arteche CID contém apenas 1 equipamento normalmente

    @staticmethod
    def carregar_equipamento(bytes_io, filename, rele_escolhido=None):
        tree = ET.parse(io.BytesIO(bytes_io))
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


class DriverSEL:
    @staticmethod
    def listar_reles(bytes_io, filename):
        reles = set()
        file_bytes = io.BytesIO(bytes_io)
        
        if filename.lower().endswith('.rdb'):
            if olefile.isOleFile(file_bytes):
                with olefile.OleFileIO(file_bytes) as ole:
                    for item in ole.listdir():
                        if len(item) >= 2 and item[0].lower() == 'relays':
                            reles.add(item[1])
        
        elif filename.lower().endswith('.zip'):
            with zipfile.ZipFile(file_bytes) as z:
                for info in z.infolist():
                    partes = info.filename.split('/')
                    if len(partes) >= 2:
                        if partes[0].lower() == 'relays':
                            reles.add(partes[1])
                        else:
                            reles.add(partes[0])
        
        return sorted(list(reles))

    @staticmethod
    def carregar_equipamento(bytes_io, filename, rele_escolhido):
        dados_sel = {'1': {}, '2': {}, '3': {}, '4': {}, '5': {}, '6': {}, 'global': {}}
        doi_por_codigo = {}
        file_bytes = io.BytesIO(bytes_io)

        def processar_texto(texto, nome_arquivo):
            dict_local = {}
            for linha in texto.splitlines():
                linha = linha.strip()
                if not linha or linha.startswith('[') or ',' not in linha: continue
                
                partes = linha.split(',', 1)
                chave = partes[0].strip()
                valor = partes[1].strip().strip('"').strip("'")
                
                # Remove comentários da SEL (tudo após o #)
                if ' # ' in valor:
                    valor = valor.split(' # ')[0].strip()
                    
                dict_local[chave] = valor
                doi_por_codigo[chave] = chave # O próprio código atua como o 'doi' para o SEL
            
            nome_lower = nome_arquivo.lower()
            if 'set_1.txt' in nome_lower: dados_sel['1'].update(dict_local)
            elif 'set_2.txt' in nome_lower: dados_sel['2'].update(dict_local)
            elif 'set_3.txt' in nome_lower: dados_sel['3'].update(dict_local)
            elif 'set_4.txt' in nome_lower: dados_sel['4'].update(dict_local)
            elif 'set_5.txt' in nome_lower: dados_sel['5'].update(dict_local)
            elif 'set_6.txt' in nome_lower: dados_sel['6'].update(dict_local)
            else: dados_sel['global'].update(dict_local)

        if filename.lower().endswith('.rdb'):
            with olefile.OleFileIO(file_bytes) as ole:
                for item in ole.listdir():
                    if len(item) >= 3 and item[0].lower() == 'relays' and item[1] == rele_escolhido:
                        nome_arq = item[-1]
                        if nome_arq.lower().endswith('.txt'):
                            stream = ole.openstream(item)
                            texto = stream.read().decode('latin-1', errors='ignore')
                            processar_texto(texto, nome_arq)

        elif filename.lower().endswith('.zip'):
            with zipfile.ZipFile(file_bytes) as z:
                for info in z.infolist():
                    partes = info.filename.split('/')
                    # Verifica se o arquivo pertence ao relé escolhido
                    if (len(partes) >= 2 and partes[1] == rele_escolhido) or (len(partes) >= 1 and partes[0] == rele_escolhido):
                        nome_arq = partes[-1]
                        if nome_arq.lower().endswith('.txt'):
                            with z.open(info) as f:
                                texto = f.read().decode('latin-1', errors='ignore')
                                processar_texto(texto, nome_arq)

        # O 'ns' não se aplica ao SEL, por isso retornamos os dicionários alocados em 'ns'
        return doi_por_codigo, {}, dados_sel

    @staticmethod
    def obter_valor(doi, dados_sel, sgroup="1"):
        if doi is None: return "NÃO ENCONTRADO"
        codigo = doi
        
        # 1º Tenta encontrar no grupo específico (ex: set_1.txt)
        if codigo in dados_sel.get(str(sgroup), {}):
            return dados_sel[str(sgroup)][codigo]
        
        # 2º Se não encontrar no grupo, busca nos arquivos globais (set_G, set_L1, set_P, etc.)
        if codigo in dados_sel.get('global', {}):
            return dados_sel['global'][codigo]
            
        return "NÃO ENCONTRADO"

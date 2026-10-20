import xml.etree.ElementTree as ET
import io

def limpar_espacos(texto):
    if texto is None: return ""  # Corrigido de 'text' para 'texto'
    return " ".join(str(texto).strip().split())

def chave_normalizada(texto):
    return "".join(limpar_espacos(texto).lower().split())

class DriverArtecheAdaTech:
    @staticmethod
    def carregar_equipamento(bytes_cid):
        tree = ET.parse(io.BytesIO(bytes_cid))
        root = tree.getroot()
        ns = {'scl': 'http://www.iec.ch/61850/2003/SCL'}
        doi_por_codigo = {}
        dois_por_ln = {}

        def eh_doi_ajuste(doi):
            if doi.find('scl:SDI[@name="setMag"]', ns) is not None:
                return True
            for dai_name in ['setVal', 'setCharact', 'setMag']:
                if doi.find(f'scl:DAI[@name="{dai_name}"]', ns) is not None:
                    return True
            for dai in doi.findall('.//scl:DAI', ns):
                if dai.attrib.get('sAddr', '').startswith('A,'):
                    return True
            return False

        for ld in root.findall('.//scl:LDevice', ns):
            ld_inst = ld.attrib.get('inst', '')
            for ln in ld.findall('scl:LN0', ns) + ld.findall('scl:LN', ns):
                ln_name = ln.attrib.get('prefix', '') + ln.attrib.get('lnClass', '') + ln.attrib.get('inst', '')
                full_ln = f"{ld_inst}.{ln_name}"
                dois_por_ln.setdefault(full_ln, {})
                for doi in ln.findall('scl:DOI', ns):
                    if not eh_doi_ajuste(doi):
                        continue
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
        if doi is None:
            return "NÃO ENCONTRADO"
        sdi_setmag = doi.find('scl:SDI[@name="setMag"]', ns)
        if sdi_setmag is not None:
            for dai in sdi_setmag.findall('scl:DAI', ns):
                vals = dai.findall('scl:Val', ns)
                if vals:
                    for v in vals:
                        if v.attrib.get('sGroup') == str(sgroup):
                            return v.text
                    return vals[0].text
        for dai_name in ['setVal', 'setCharact', 'setMag']:
            dai = doi.find(f'scl:DAI[@name="{dai_name}"]', ns)
            if dai is not None:
                vals = dai.findall('scl:Val', ns)
                if vals:
                    for v in vals:
                        if v.attrib.get('sGroup') == str(sgroup):
                            return v.text
                    return vals[0].text
        for dai in doi.findall('.//scl:DAI', ns):
            if dai.attrib.get('sAddr', '').startswith('A,'):
                vals = dai.findall('scl:Val', ns)
                if vals:
                    for v in vals:
                        if v.attrib.get('sGroup') == str(sgroup):
                            return v.text
                    return vals[0].text
        return "VAZIO"
# -*- coding: utf-8 -*-
"""
Gera o DANFSe (Documento Auxiliar da NFS-e) em PDF LOCALMENTE, a partir do
XML da nota já baixado — sem depender da API do governo.

Motivo: a API oficial de geração do DANFSe (GET /danfse/{chave} no ADN) foi
DESATIVADA pelo governo em 03/08/2026 (Nota Técnica 008/2026, "DANFSe 2.0")
por sobrecarregar a infraestrutura nacional. A partir de agora cada sistema
é responsável por montar o PDF a partir do XML já autorizado, seguindo o
mesmo leiaute padronizado. Não depende mais do Portal Nacional estar no ar.

Este leiaute segue a Nota Técnica nº 008 (SE/CGNFS-e), versão 1.02, de
14/07/2026 — posições, tamanhos de fonte, sombreamentos e textos de blocos
suprimidos conferidos campo a campo contra o Anexo I (modelo oficial) e
contra um DANFSe real baixado do Portal Nacional. Implementação própria em
reportlab; a logomarca (nfse_logo.png) é o arquivo oficial disponibilizado
pelo próprio governo em gov.br/nfse.

Pontos que seguem a norma "à risca" (por pedido explícito, mesmo quando
divergem do que pareceria mais "correto" contabilmente):
  - Valor Líquido da NFS-e = campo <vLiq> do XML, IMPRESSO DIRETO, sem
    recalcular. Se o emissor errou esse campo, o DANFSe reproduz o erro
    dele — é o que o Portal Nacional também faria.
  - "Contribuições Sociais - Retidas": tpRetPisCofins = 3 -> soma
    (CSLL+PIS+COFINS) aparece retida; tpRetPisCofins = 1 -> PIS/COFINS
    aparecem em "Débito Apuração Própria" (não retidos).

Uso:
    from gerar_danfse import gerar_pdf_danfse
    pdf_bytes = gerar_pdf_danfse(xml_bytes, marca_dagua="CANCELADA")  # ou ""
"""

import re
import sys
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET

import qrcode
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas

# Em .exe congelado (PyInstaller), __file__ não aponta pra um caminho real em
# disco — os dados empacotados (datas=[...] no .spec) ficam extraídos em
# sys._MEIPASS (onefile e onedir).
PASTA_SCRIPT = Path(getattr(sys, "_MEIPASS", "")) if getattr(sys, "frozen", False) \
    else Path(__file__).resolve().parent
ARQ_LOGO = PASTA_SCRIPT / "nfse_logo.png"

CINZA_5 = colors.Color(0.95, 0.95, 0.95)   # sombreamento 5% (item 2.2.3)
CINZA_MARCA = colors.Color(0.65, 0.65, 0.65)  # marca d'água K35 (item 2.5.1/2)
PRETO = colors.black
VERMELHO = colors.Color(1, 0, 0)

# Pequena tabela dos municípios mais frequentes na nossa base (a norma pede
# a tabela completa do IBGE — 5.570 municípios; por ora mantemos só os
# principais, com o código de 7 dígitos do IBGE como chave). Fora daqui,
# mostramos só a UF (derivada do prefixo do código) sem o nome da cidade.
NOME_MUNICIPIO_IBGE = {
    "3304557": "Rio de Janeiro", "3550308": "São Paulo", "3106200": "Belo Horizonte",
    "3303500": "Niterói", "3301702": "Duque de Caxias", "3170206": "Uberlândia",
    "5300108": "Brasília", "2927408": "Salvador", "4106902": "Curitiba",
    "4314902": "Porto Alegre", "2611606": "Recife", "2304400": "Fortaleza",
    "1302603": "Manaus", "1501402": "Belém", "5208707": "Goiânia",
    "3509502": "Campinas", "3518800": "Guarulhos", "3543402": "Santo André",
    "3548500": "São Bernardo do Campo", "3303302": "Nova Iguaçu",
    "3300100": "Angra dos Reis", "4205407": "Florianópolis", "3205309": "Vitória",
}

UF_POR_PREFIXO_IBGE = {
    "11": "RO", "12": "AC", "13": "AM", "14": "RR", "15": "PA", "16": "AP", "17": "TO",
    "21": "MA", "22": "PI", "23": "CE", "24": "RN", "25": "PB", "26": "PE", "27": "AL",
    "28": "SE", "29": "BA",
    "31": "MG", "32": "ES", "33": "RJ", "35": "SP",
    "41": "PR", "42": "SC", "43": "RS",
    "50": "MS", "51": "MT", "52": "GO", "53": "DF",
}


def _uf_do_codigo_mun(cod: str) -> str:
    return UF_POR_PREFIXO_IBGE.get((cod or "")[:2], "")


def _nome_municipio(cod: str) -> str:
    return NOME_MUNICIPIO_IBGE.get(cod or "", "")


def _localname(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _find(el, path):
    """Acha o 1º descendente cujo nome local bate com cada passo do path
    (ex.: 'infNFSe/emit/enderNac'), ignorando namespace."""
    if el is None:
        return None
    atual = [el]
    for passo in path.split("/"):
        prox = None
        for no in atual:
            for filho in no:
                if _localname(filho.tag) == passo:
                    prox = filho
                    break
            if prox is not None:
                break
        if prox is None:
            return None
        atual = [prox]
    return atual[0]


def _first(*elementos):
    for el in elementos:
        if el is not None:
            return el
    return None


def _find_deep(el, nome):
    """Busca em qualquer profundidade (usado pra campos que mudam de posição
    entre versões da NFS-e: retenções, IBS/CBS, informações complementares)."""
    if el is None:
        return None
    for no in el.iter():
        if _localname(no.tag) == nome:
            return no
    return None


def _txt(el, nome):
    f = _find(el, nome) if "/" in nome else None
    if f is None and el is not None:
        for filho in el:
            if _localname(filho.tag) == nome:
                f = filho
                break
    return (f.text or "").strip() if f is not None and f.text else ""


def _deep(el, nome):
    f = _find_deep(el, nome)
    return (f.text or "").strip() if f is not None and f.text else ""


def _num(v, padrao=0.0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return padrao


def _money(v):
    try:
        return f"R$ {float(v):,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")
    except (TypeError, ValueError):
        return ""


def _pct(v):
    try:
        return f"{float(v):.2f} %".replace(".", ",")
    except (TypeError, ValueError):
        return ""


def _doc(v):
    """Formata CNPJ (14 díg.) ou CPF (11 díg.) com pontuação. Deixa como veio
    se não bater nenhum dos dois tamanhos (ex.: NIF de estrangeiro)."""
    d = re.sub(r"\D", "", v or "")
    if len(d) == 14:
        return f"{d[0:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:14]}"
    if len(d) == 11:
        return f"{d[0:3]}.{d[3:6]}.{d[6:9]}-{d[9:11]}"
    return v


def _cod_nbs(v):
    d = re.sub(r"\D", "", v or "")
    return f"{d[0]}.{d[1:5]}.{d[5:7]}.{d[7:9]}" if len(d) == 9 else v


def _cod_ibge(v):
    d = re.sub(r"\D", "", v or "")
    return f"{d[0:2]}.{d[2:7]}" if len(d) == 7 else v


def _cep(v):
    d = re.sub(r"\D", "", v or "")
    return f"{d[0:2]}.{d[2:5]}-{d[5:8]}" if len(d) == 8 else v


def _data_br(iso: str) -> str:
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", iso or "")
    return f"{m.group(3)}/{m.group(2)}/{m.group(1)}" if m else (iso or "")


def _data_hora_br(iso: str) -> str:
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})T(\d{2}:\d{2}:\d{2})", iso or "")
    return f"{m.group(3)}/{m.group(2)}/{m.group(1)} {m.group(4)}" if m else _data_br(iso)


def _endereco(end_el):
    if end_el is None:
        return ""
    partes = [_txt(end_el, "xLgr"), _txt(end_el, "nro"), _txt(end_el, "xCpl"), _txt(end_el, "xBairro")]
    return ", ".join(p for p in partes if p)


def _municipio_uf(cod: str, uf_xml: str = "") -> str:
    uf = uf_xml or _uf_do_codigo_mun(cod)
    nome = _nome_municipio(cod)
    return " / ".join(p for p in [nome, uf] if p) or "-"


def _pessoa(el, end_el, end_nac_el):
    if el is None:
        return None
    cmun = _txt(end_nac_el, "cMun")
    return {
        "cnpjCpf": _txt(el, "CNPJ") or _txt(el, "CPF") or _txt(el, "NIF"),
        "im": _txt(el, "IM"),
        "nome": _txt(el, "xNome"),
        "endereco": _endereco(end_el),
        "municipioUF": _municipio_uf(cmun, _txt(end_nac_el, "UF")),
        "codIbge": _cod_ibge(cmun),
        "cep": _cep(_txt(end_nac_el, "CEP")),
        "fone": _txt(el, "fone"),
        "email": _txt(el, "email"),
    }


def parse_danfse(xml_bytes: bytes) -> dict | None:
    """Extrai do XML da NFS-e os campos necessários pro DANFSe. Devolve None
    se não for uma NFS-e válida (ex.: é um evento)."""
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return None

    infNFSe = root if _localname(root.tag) == "infNFSe" else _find_deep(root, "infNFSe")
    if infNFSe is None:
        return None

    id_attr = infNFSe.get("Id", "")
    chave = re.sub(r"\D", "", id_attr)
    if not chave:
        return None

    emit = _find(infNFSe, "emit")
    emitEnd = _first(_find(emit, "enderNac"), _find(emit, "enderExt"))
    valores = _find(infNFSe, "valores")
    ibsCbsNFSe = _find(infNFSe, "IBSCBS")
    DPS = _find(infNFSe, "DPS")
    infDPS = _find(DPS, "infDPS")
    prest = _find(infDPS, "prest")
    regTrib = _find(prest, "regTrib")
    toma = _find(infDPS, "toma")
    tomaEnd = _find(toma, "end")
    tomaEndNac = _first(_find(tomaEnd, "endNac"), _find(tomaEnd, "endExt"), tomaEnd)
    interm = _first(_find(infDPS, "interm"), _find(infDPS, "intermediario"))
    intermEnd = _find(interm, "end")
    intermEndNac = _first(_find(intermEnd, "endNac"), _find(intermEnd, "endExt"), intermEnd)
    dest = _find(ibsCbsNFSe, "dest") or _find(infDPS, "dest")
    destEnd = _find(dest, "end")
    destEndNac = _first(_find(destEnd, "endNac"), _find(destEnd, "endExt"), destEnd)
    serv = _find(infDPS, "serv")
    cServ = _find(serv, "cServ")
    locPrest = _find(serv, "locPrest")
    valDPS = _find(infDPS, "valores")
    trib = _find(valDPS, "trib")
    tribMun = _find(trib, "tribMun")
    tribFed = _find(trib, "tribFed")
    totTrib = _find(trib, "totTrib")
    ibsCbsDPS = _find(infDPS, "IBSCBS")

    opSimp = _txt(regTrib, "opSimpNac")
    opSimpLabel = {"1": "Não Optante", "2": "Optante - MEI", "3": "Optante - ME/EPP"}.get(opSimp, "-")
    cStat = _txt(infNFSe, "cStat")
    tpEmit = _txt(infDPS, "tpEmit")
    tpEmitLabel = {"1": "Prestador", "2": "Tomador", "3": "Intermediário"}.get(tpEmit, "-")
    finNFSe = _txt(infDPS, "finNFSe")
    finalidadeLabel = {"1": "NFS-e normal", "2": "NFS-e complementar",
                       "3": "NFS-e de ajuste", "4": "NFS-e de substituição"}.get(finNFSe, "NFS-e regular")

    municipioEmit = _txt(infNFSe, "xLocEmi")
    ufEmit = _txt(emitEnd, "UF") or _uf_do_codigo_mun(_txt(emitEnd, "cMun"))
    localPrestNome = _txt(infNFSe, "xLocPrestacao")
    ufLocPrest = _uf_do_codigo_mun(_txt(locPrest, "cLocPrestacao"))
    municipioIncidNome = _txt(infNFSe, "xLocIncid") or localPrestNome
    ufLocIncid = _uf_do_codigo_mun(_txt(infNFSe, "cLocIncid"))

    retISSQN = _deep(tribMun, "tpRetISSQN")
    retISSQNLabel = {"1": "Não Retido", "2": "Retido pelo Tomador", "3": "Retido pelo Intermediário"}.get(retISSQN, "Não Retido")

    # ── PIS/COFINS: código 3 = retido (some com CSLL); código 1 = débito
    # próprio do prestador (não retido). Confirmado pela Erika contra um
    # DANFSe real do Portal — a redação da NT-008 sobre esse campo está
    # incompleta/ambígua (a tabela de códigos completa não está no
    # documento), então seguimos o comportamento confirmado. ──
    tpRetPisCofins = _deep(tribFed, "tpRetPisCofins")
    pis = _num(_deep(valDPS, "vPis"))
    cofins = _num(_deep(valDPS, "vCofins"))
    csll = _num(_deep(valDPS, "vRetCSLL"))
    irrf = _num(_deep(valDPS, "vRetIRRF"))
    cp = _num(_deep(valDPS, "vRetCP"))

    if tpRetPisCofins == "3":
        contribSociaisRetidas = round(csll + pis + cofins, 2)
        pis_proprio, cofins_proprio = 0.0, 0.0
    elif tpRetPisCofins == "1":
        contribSociaisRetidas = round(csll, 2)
        pis_proprio, cofins_proprio = pis, cofins
    else:
        # indicador ausente/outro: usa o total autoritativo da nota (vTotalRet)
        # pra não perder retenção real quando o emissor não preenche o
        # indicador direito (caso observado em nota real do Google).
        vTotalRet_bruto = _num(_txt(valores, "vTotalRet"))
        vISSQNRet_bruto = _num(_txt(valores, "vISSQNRet")) or (_num(_txt(valores, "vISSQN")) if retISSQN in ("2", "3") else 0.0)
        sobra_federal = max(vTotalRet_bruto - vISSQNRet_bruto - irrf - cp, 0.0)
        if sobra_federal > 0:
            contribSociaisRetidas = round(sobra_federal, 2)
            pis_proprio, cofins_proprio = 0.0, 0.0
        else:
            contribSociaisRetidas = round(csll, 2)
            pis_proprio, cofins_proprio = pis, cofins

    descRetLabel = {
        "1": "1 - PIS/COFINS/CSLL Não Retido", "2": "2 - PIS/COFINS Retido",
        "3": "3 - PIS/COFINS/CSLL Retidos", "4": "4 - CSLL Retida",
    }.get(tpRetPisCofins, "-")

    vServ = _num(_txt(valores, "vServ")) or _num(_deep(valDPS, "vServ"))
    descIncond = _num(_txt(valores, "vDescIncond"))
    descCond = _num(_txt(valores, "vDescCond") or _deep(tribMun, "vDescCond"))
    # Valor Líquido = campo <vLiq> IMPRESSO DIRETO (norma NT-008, item
    # 2.1.11) — não recalculamos, mesmo que o emissor tenha preenchido
    # errado (o Portal Nacional também só imprime o campo).
    vLiq_xml = _txt(valores, "vLiq")

    ibs_vBC = _deep(ibsCbsNFSe, "vBC")
    vCBS = _num(_deep(ibsCbsNFSe, "vCBS"))
    vIBSTot = _num(_deep(ibsCbsNFSe, "vIBSTot"))
    vTotNF = _deep(ibsCbsNFSe, "vTotNF")

    return {
        "chave": chave,
        "municipioEmit": municipioEmit,
        "ufEmit": ufEmit,
        "ambGer": _txt(infNFSe, "ambGer"),
        "tpAmb": _txt(infDPS, "tpAmb"),
        "nNFSe": _txt(infNFSe, "nNFSe"),
        "competencia": _data_br(_txt(infDPS, "dCompet")[:10]),
        "dhEmiNFSe": _data_hora_br(_txt(infNFSe, "dhProc")),
        "dhEmiDPS": _data_hora_br(_txt(infDPS, "dhEmi")),
        "nDPS": _txt(infDPS, "nDPS"),
        "serie": _txt(infDPS, "serie"),
        "tpEmitLabel": tpEmitLabel,
        "situacao": "NFS-e Gerada" if cStat == "100" else (f"Situação {cStat}" if cStat else "-"),
        "finalidade": finalidadeLabel,
        "prest": {
            "cnpjCpf": _txt(emit, "CNPJ") or _txt(emit, "CPF"),
            "im": _txt(emit, "IM"),
            "nome": _txt(emit, "xNome"),
            "endereco": _endereco(emitEnd),
            "municipioUF": _municipio_uf(_txt(emitEnd, "cMun"), ufEmit) if _txt(emitEnd, "cMun") else (municipioEmit + (" / " + ufEmit if ufEmit else "")),
            "codIbge": _cod_ibge(_txt(emitEnd, "cMun")),
            "cep": _cep(_txt(emitEnd, "CEP")),
            "fone": _txt(emit, "fone"),
            "email": _txt(emit, "email"),
            "simples": opSimpLabel,
            "regApSN": _txt(regTrib, "regApTribSN") or "-",
        },
        "toma": _pessoa(toma, tomaEnd, tomaEndNac),
        "interm": _pessoa(interm, intermEnd, intermEndNac),
        "dest": _pessoa(dest, destEnd, destEndNac),
        "destEhTomador": bool(dest) and toma is not None and _txt(dest, "CNPJ") == _txt(toma, "CNPJ") and _txt(dest, "CNPJ") != "",
        "serv": {
            "cTribNac": _txt(cServ, "cTribNac"),
            "xTribNac": _txt(infNFSe, "xTribNac"),
            "cTribMun": _txt(cServ, "cTribMun"),
            "xTribMun": _txt(cServ, "xTribMun"),
            "xDescServ": _txt(cServ, "xDescServ"),
            "cNBS": _cod_nbs(_txt(cServ, "cNBS")),
            "localPrestacao": " / ".join(p for p in [localPrestNome, ufLocPrest, "-"] if p),
        },
        "issqn": {
            "tipoTrib": "Operação Tributável" if _deep(tribMun, "tribISSQN") == "1" else "-",
            "municipioIncid": " / ".join(p for p in [municipioIncidNome, ufLocIncid, "-"] if p),
            "bc": _txt(valores, "vBC"),
            "aliq": _txt(valores, "pAliqAplic"),
            "valor": _txt(valores, "vISSQN"),
            "retencaoLabel": retISSQNLabel,
            "retido": retISSQN in ("2", "3"),
            "temIncidencia": bool(_txt(valores, "vBC") or _deep(tribMun, "tribISSQN")),
        },
        "fed": {
            "irrf": irrf,
            "cp": cp,
            "contribSociaisRetidas": contribSociaisRetidas,
            "pisProprio": pis_proprio,
            "cofinsProprio": cofins_proprio,
            "descRetLabel": descRetLabel,
        },
        "ibscbs": {
            "cst": _deep(ibsCbsDPS, "CST"),
            "cClassTrib": _deep(ibsCbsDPS, "cClassTrib"),
            "cIndOp": _deep(ibsCbsDPS, "cIndOp"),
            "codIncid": _txt(ibsCbsNFSe, "cLocalidadeIncid"),
            "municipioIncid": _txt(ibsCbsNFSe, "xLocalidadeIncid"),
            "ufIncid": _uf_do_codigo_mun(_txt(ibsCbsNFSe, "cLocalidadeIncid")),
            "vBC": ibs_vBC,
            "pRedAliqUF": _deep(ibsCbsNFSe, "pRedAliqUF"),
            "pRedAliqMun": _deep(ibsCbsNFSe, "pRedAliqMun"),
            "pRedAliqCBS": _deep(ibsCbsNFSe, "pRedAliqCBS"),
            "pIBSUF": _deep(ibsCbsNFSe, "pIBSUF"),
            "pIBSMun": _deep(ibsCbsNFSe, "pIBSMun"),
            "pAliqEfetUF": _deep(ibsCbsNFSe, "pAliqEfetUF"),
            "vIBSUF": _deep(ibsCbsNFSe, "vIBSUF"),
            "pAliqEfetMun": _deep(ibsCbsNFSe, "pAliqEfetMun"),
            "vIBSMun": _deep(ibsCbsNFSe, "vIBSMun"),
            "vIBSTot": _deep(ibsCbsNFSe, "vIBSTot"),
            "pCBS": _deep(ibsCbsNFSe, "pCBS"),
            "pAliqEfetCBS": _deep(ibsCbsNFSe, "pAliqEfetCBS"),
            "vCBS": _deep(ibsCbsNFSe, "vCBS"),
            "totalIBSCBS": _money(round(vCBS + vIBSTot, 2)) if (ibsCbsNFSe is not None) else "",
            "temDados": ibsCbsNFSe is not None and bool(ibs_vBC),
        },
        "totais": {
            "vServ": vServ,
            "descIncond": descIncond,
            "descCond": descCond,
            "vTotalRet": _txt(valores, "vTotalRet"),
            "valorLiquido": vLiq_xml,
            "valorLiquidoIBSCBS": vTotNF or vLiq_xml,
        },
        "totaisAprox": {
            "fed": _deep(totTrib, "vTotTribFed"),
            "est": _deep(totTrib, "vTotTribEst"),
            "mun": _deep(totTrib, "vTotTribMun"),
        },
        "infoCompl": _deep(serv, "xInfComp"),
    }


def _wrap(texto: str, largura_max_chars: int) -> list[str]:
    texto = re.sub(r"\s+", " ", texto or "").strip()
    if not texto:
        return []
    linhas, atual = [], ""
    for palavra in texto.split(" "):
        teste = (atual + " " + palavra).strip()
        if len(teste) > largura_max_chars and atual:
            linhas.append(atual)
            atual = palavra
        else:
            atual = teste
    if atual:
        linhas.append(atual)
    return linhas


# ───────────────────────── Layout (coordenadas em cm, NT-008 Anexo I) ─────────────────────────

PAGINA_L, PAGINA_A = 21.0, 29.7   # A4 em cm
MARGEM = 0.17


def gerar_pdf_danfse(xml_bytes: bytes, marca_dagua: str = "") -> bytes | None:
    """Gera o PDF (DANFSe) a partir dos bytes do XML da NFS-e. `marca_dagua`
    = 'CANCELADA' / 'SUBSTITUÍDA' / '' (nota regular). Devolve None se o XML
    não for uma NFS-e (ex.: é um evento)."""
    dados = parse_danfse(xml_bytes)
    if dados is None:
        return None

    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    X0, X1 = MARGEM, PAGINA_L - MARGEM
    LARG = X1 - X0

    def Y(sup_cm):
        """Converte 'distância do topo' (como na norma) pra Y do reportlab
        (que mede a partir de baixo)."""
        return (PAGINA_A - sup_cm) * cm

    def cm_(v):
        return v * cm

    def _clip(texto, fonte, tam, largura_disp):
        if c.stringWidth(texto, fonte, tam) <= largura_disp:
            return texto
        while texto and c.stringWidth(texto + "...", fonte, tam) > largura_disp:
            texto = texto[:-1]
        return texto + "..." if texto else ""

    def separador(sup, grosso=False):
        """Linha divisória ENTRE blocos (não entre campos de um mesmo bloco) —
        é essa linha, junto com o sombreado do título, que separa visualmente
        uma seção da outra, igual ao padrão nacional."""
        c.setLineWidth(0.9 if grosso else 0.6)
        c.setStrokeColor(PRETO)
        c.line(cm_(X0), Y(sup), cm_(X1), Y(sup))

    def cel(esq, sup, larg, alt, label="", valor="", negrito=False, sombreado=False,
           tam_label=6, tam_valor=7, moldura=False, label_caps=False, centralizado=False,
           cor_valor=PRETO):
        x = cm_(X0 + esq)
        y_top = Y(sup)
        h = cm_(alt)
        w = cm_(larg)
        pad = 0.06 * cm
        if sombreado:
            c.setFillColor(CINZA_5)
            c.rect(x, y_top - h, w, h, fill=1, stroke=0)
        if moldura:
            c.setLineWidth(0.5)
            c.setStrokeColor(PRETO)
            c.rect(x, y_top - h, w, h, fill=0, stroke=1)
        if label:
            c.setFillColor(PRETO)
            c.setFont("Helvetica-Bold", tam_label)
            txt_label = label.upper() if label_caps else label
            c.drawString(x + pad, y_top - 0.24 * cm, _clip(txt_label, "Helvetica-Bold", tam_label, w - 2 * pad))
        if valor not in (None, ""):
            vfont = "Helvetica-Bold" if negrito else "Helvetica"
            c.setFillColor(cor_valor)
            c.setFont(vfont, tam_valor)
            yy = y_top - h + 0.16 * cm if label else y_top - h + (h - cm_(tam_valor) / 28.0) / 2 + 0.05 * cm
            texto = _clip(str(valor), vfont, tam_valor, w - 2 * pad)
            if centralizado:
                tw = c.stringWidth(texto, vfont, tam_valor)
                c.drawString(x + (w - tw) / 2, yy, texto)
            else:
                c.drawString(x + pad, yy, texto)
        c.setFillColor(PRETO)

    def linha_central(esq, sup, larg, alt, texto):
        separador(sup)
        c.setFont("Helvetica-Bold", 6.5)
        tw = c.stringWidth(texto, "Helvetica-Bold", 6.5)
        c.drawCentredString(cm_(X0 + esq + larg / 2), Y(sup) - cm_(alt) / 2 - 2.2, texto)

    # ══════════════════════════ CABEÇALHO (Sup 0,30 a 1,46) ══════════════════════════
    cel(0, 0.30, LARG, 1.16, moldura=True)
    try:
        logo = ET = None
        from reportlab.lib.utils import ImageReader
        if ARQ_LOGO.exists():
            img = ImageReader(str(ARQ_LOGO))
            iw, ih = img.getSize()
            lw = cm_(4.00)
            lh = lw * ih / iw
            c.drawImage(img, cm_(X0 + 0.49), Y(0.44) - lh, width=lw, height=lh, mask='auto')
    except Exception:
        pass

    c.setFont("Helvetica-Bold", 9)
    c.drawCentredString(cm_(X0 + 5.41 + 10.19 / 2), Y(0.30) - 0.42 * cm, "DANFSe v2.0")
    c.drawCentredString(cm_(X0 + 5.41 + 10.19 / 2), Y(0.30) - 0.80 * cm, "Documento Auxiliar da NFS-e")
    if dados["tpAmb"] == "2":
        c.setFillColor(VERMELHO)
        c.setFont("Helvetica-Bold", 9)
        c.drawCentredString(cm_(X0 + 5.41 + 10.19 / 2), Y(0.30) - 1.12 * cm, "NFS-e SEM VALIDADE JURÍDICA")
        c.setFillColor(PRETO)

    mun_txt = f"Município: {dados['municipioEmit'] or '-'}" + (f" / {dados['ufEmit']}" if dados["ufEmit"] else "")
    c.setFont("Helvetica", 8)
    c.drawString(cm_(X0 + 15.62), Y(0.30) - 0.28 * cm, mun_txt[:37])
    c.setFont("Helvetica", 6)
    c.drawString(cm_(X0 + 15.62), Y(0.97) - 0.18 * cm, f"Ambiente Gerador: {dados['ambGer'] or '-'}")
    c.drawString(cm_(X0 + 15.62), Y(1.22) - 0.18 * cm, f"Tipo de Ambiente: {dados['tpAmb'] or '-'}")

    # ══════════════════════════ DADOS DA NFS-e (Sup 1,48 a 4,32) ══════════════════════════
    separador(1.48, grosso=True)
    cel(0, 1.48, 15.30, 0.77, "CHAVE DE ACESSO DA NFS-e", dados["chave"], tam_label=7, label_caps=True, tam_valor=7.5, moldura=False)
    cel(0, 2.27, 5.09, 0.67, "NÚMERO DA NFS-e", dados["nNFSe"], tam_label=7, label_caps=True, moldura=False)
    cel(5.41, 2.27, 5.09, 0.67, "COMPETÊNCIA DA NFS-e", dados["competencia"], tam_label=7, label_caps=True, moldura=False)
    cel(10.51, 2.27, 5.09, 0.67, "DATA E HORA DA EMISSÃO DA NFS-e", dados["dhEmiNFSe"], tam_label=7, label_caps=True, moldura=False)
    cel(0, 2.96, 5.09, 0.67, "NÚMERO DA DPS", dados["nDPS"], tam_label=7, label_caps=True, moldura=False)
    cel(5.41, 2.96, 5.09, 0.67, "SÉRIE DA DPS", dados["serie"], tam_label=7, label_caps=True, moldura=False)
    cel(10.51, 2.96, 5.09, 0.67, "DATA E HORA DA EMISSÃO DA DPS", dados["dhEmiDPS"], tam_label=7, label_caps=True, moldura=False)
    cel(0, 3.65, 5.09, 0.67, "EMITENTE DA NFS-e", dados["tpEmitLabel"], tam_label=7, label_caps=True, moldura=False, sombreado=True)
    cel(5.41, 3.65, 5.09, 0.67, "SITUAÇÃO DA NFS-e", dados["situacao"], tam_label=7, label_caps=True, moldura=False)
    cel(10.51, 3.65, 5.09, 0.67, "FINALIDADE", dados["finalidade"], tam_label=7, label_caps=True, moldura=False)
    # linhas horizontais internas do bloco
    for sup_linha in (2.27, 2.96, 3.65, 4.32):
        c.setLineWidth(0.5)
        c.line(cm_(X0), Y(sup_linha), cm_(X0 + 15.30 if sup_linha == 2.27 else X0 + LARG), Y(sup_linha))

    # QR Code (X:17,48 Y:1,67, mínimo 1,52x1,52)
    url_consulta = f"https://www.nfse.gov.br/ConsultaPublica/?tpc=1&chave={dados['chave']}"
    qr = qrcode.QRCode(border=1, error_correction=qrcode.constants.ERROR_CORRECT_M)
    qr.add_data(url_consulta)
    qr.make(fit=True)
    matriz = qr.get_matrix()
    qr_tam = cm_(1.52)
    qr_x, qr_y_top = cm_(X0 + 17.48), Y(1.67)
    n = len(matriz)
    modulo = qr_tam / n
    c.setFillColor(PRETO)
    for lin_i, linha_mod in enumerate(matriz):
        for col_i, escuro in enumerate(linha_mod):
            if escuro:
                c.rect(qr_x + col_i * modulo, qr_y_top - qr_tam + (n - 1 - lin_i) * modulo,
                       modulo, modulo, fill=1, stroke=0)
    c.setFont("Helvetica", 6)
    legenda = _wrap("A autenticidade desta NFS-e pode ser verificada pela leitura deste código "
                    "QR ou pela consulta da chave de acesso no portal nacional da NFS-e", 44)
    ly = Y(3.36) - 0.22 * cm
    for linha_txt in legenda[:3]:
        c.drawString(cm_(X0 + 15.80), ly, linha_txt)
        ly -= 0.24 * cm

    # ══════════════════════════ Cursor dinâmico a partir daqui ══════════════════════════
    cursor = 4.32   # cm do topo

    def bloco_titulo_com_campos(titulo, campos, alt=0.63, sombrear_titulo=True):
        """Início de um BLOCO novo (Prestador, Tomador, Tributação...): linha
        separadora ACIMA (fronteira entre blocos) + título sombreado + campos
        na mesma linha. campos = [(esq, larg, label, valor, negrito), ...]."""
        nonlocal cursor
        separador(cursor, grosso=True)
        cel(0, cursor, 5.09, alt, titulo, "", tam_label=7, label_caps=True, sombreado=sombrear_titulo)
        for esq, larg, label, valor, negrito in campos:
            cel(esq, cursor, larg, alt, label, valor, negrito=negrito)
        cursor += alt

    def linha_campos(campos, alt=0.63):
        """Linha de CONTINUAÇÃO dentro do bloco atual — sem separador (as
        linhas de um mesmo bloco não têm grade entre si, só espaçamento)."""
        nonlocal cursor
        for esq, larg, label, valor, negrito in campos:
            cel(esq, cursor, larg, alt, label, valor, negrito=negrito)
        cursor += alt

    def bloco_ausente(texto, alt=0.32):
        nonlocal cursor
        linha_central(0, cursor, LARG, alt, texto)
        cursor += alt

    # ── Prestador / Fornecedor (sempre presente) ──
    p = dados["prest"]
    bloco_titulo_com_campos("Prestador / Fornecedor", [
        (5.41, 5.09, "CNPJ / CPF / NIF", _doc(p["cnpjCpf"]), False),
        (10.51, 5.09, "Indicador Municipal (Inscrição)", p["im"], False),
        (15.62, 5.09, "Telefone", p["fone"], False),
    ])
    linha_campos([
        (0, 10.19, "Nome / Nome Empresarial", p["nome"], False),
        (10.51, 5.09, "Município / Sigla UF", p["municipioUF"], False),
        (15.62, 5.09, "Código IBGE / CEP", f"{p['codIbge']} / {p['cep']}", False),
    ])
    linha_campos([
        (0, 10.19, "Endereço", p["endereco"], False),
        (10.51, 10.19, "E-mail", p["email"], False),
    ])
    linha_campos([
        (0, 10.19, "Simples Nacional na Data de Competência", p["simples"], False),
        (10.51, 10.19, "Regime de Apuração Tributária pelo SN", p["regApSN"], False),
    ])

    # ── Tomador / Adquirente (colapsa se ausente) ──
    if dados["toma"]:
        t = dados["toma"]
        bloco_titulo_com_campos("Tomador / Adquirente", [
            (5.41, 5.09, "CNPJ / CPF / NIF", _doc(t["cnpjCpf"]), False),
            (10.51, 5.09, "Indicador Municipal (Inscrição)", t["im"], False),
            (15.62, 5.09, "Telefone", t["fone"], False),
        ])
        linha_campos([
            (0, 10.19, "Nome / Nome Empresarial", t["nome"], False),
            (10.51, 5.09, "Município / Sigla UF", t["municipioUF"], False),
            (15.62, 5.09, "Código IBGE / CEP", f"{t['codIbge']} / {t['cep']}", False),
        ])
        linha_campos([
            (0, 10.19, "Endereço", t["endereco"], False),
            (10.51, 10.19, "E-mail", t["email"], False),
        ])
    else:
        bloco_ausente("TOMADOR/ADQUIRENTE DA OPERAÇÃO NÃO IDENTIFICADO NA NFS-e")

    # ── Destinatário da Operação (colapsa se ausente OU se igual ao tomador) ──
    if dados["destEhTomador"]:
        bloco_ausente("O DESTINATÁRIO É O PRÓPRIO TOMADOR/ADQUIRENTE DA OPERAÇÃO")
    elif dados["dest"]:
        d = dados["dest"]
        bloco_titulo_com_campos("Destinatário da Operação", [
            (5.41, 5.09, "CNPJ / CPF / NIF", _doc(d["cnpjCpf"]), False),
            (10.51, 5.09, "", "", False),
            (15.62, 5.09, "Telefone", d["fone"], False),
        ])
        linha_campos([
            (0, 10.19, "Nome / Nome Empresarial", d["nome"], False),
            (10.51, 5.09, "Município / Sigla UF", d["municipioUF"], False),
            (15.62, 5.09, "Código IBGE / CEP", f"{d['codIbge']} / {d['cep']}", False),
        ])
        linha_campos([
            (0, 10.19, "Endereço", d["endereco"], False),
            (10.51, 10.19, "E-mail", d["email"], False),
        ])
    else:
        bloco_ausente("DESTINATÁRIO DA OPERAÇÃO NÃO IDENTIFICADO NA NFS-e")

    # ── Intermediário da Operação (colapsa se ausente) ──
    if dados["interm"]:
        it = dados["interm"]
        bloco_titulo_com_campos("Intermediário da Operação", [
            (5.41, 5.09, "CNPJ / CPF / NIF", _doc(it["cnpjCpf"]), False),
            (10.51, 5.09, "Indicador Municipal (Inscrição)", it["im"], False),
            (15.62, 5.09, "Telefone", it["fone"], False),
        ])
        linha_campos([
            (0, 10.19, "Nome / Nome Empresarial", it["nome"], False),
            (10.51, 5.09, "Município / Sigla UF", it["municipioUF"], False),
            (15.62, 5.09, "Código IBGE / CEP", f"{it['codIbge']} / {it['cep']}", False),
        ])
        linha_campos([
            (0, 10.19, "Endereço", it["endereco"], False),
            (10.51, 10.19, "E-mail", it["email"], False),
        ])
    else:
        bloco_ausente("INTERMEDIÁRIO DA OPERAÇÃO NÃO IDENTIFICADO NA NFS-e")

    # ── Serviço Prestado ──
    s = dados["serv"]
    bloco_titulo_com_campos("Serviço Prestado", [
        (5.41, 5.09, "Código de Tributação Nacional / Municipal", f"{s['cTribNac']} / {s['cTribMun']}", False),
        (10.51, 5.09, "Código da NBS", s["cNBS"], False),
        (15.62, 5.09, "Local da Prestação / Sigla UF / País", s["localPrestacao"], False),
    ])
    desc_trib = s["xTribMun"] or s["xTribNac"]
    cel(0, cursor, LARG, 0.5, "", desc_trib, tam_valor=6.6)
    cursor += 0.5

    # "Descrição do Serviço" — altura elástica (consome o espaço sobrando)
    linhas_desc = _wrap(s["xDescServ"], 128)
    altura_desc = max(0.63, 0.30 + 0.30 * min(len(linhas_desc), 6))
    cel(0, cursor, LARG, altura_desc, "Descrição do Serviço", "")
    ty = Y(cursor) - 0.55 * cm
    c.setFont("Helvetica", 7)
    for linha_txt in linhas_desc[:6]:
        c.drawString(cm_(X0 + 0.06), ty, linha_txt)
        ty -= 0.28 * cm
    cursor += altura_desc

    # ── Tributação Municipal (ISSQN) — colapsa se não houver incidência ──
    issqn = dados["issqn"]
    if issqn["temIncidencia"]:
        bloco_titulo_com_campos("Tributação Municipal (ISSQN)", [
            (5.41, 5.09, "Tipo de Tributação do ISSQN", issqn["tipoTrib"], False),
            (10.51, 9.89, "Município / Sigla UF / País de Incidência do ISSQN", issqn["municipioIncid"], False),
        ])
        linha_campos([
            (0, 5.09, "BC ISSQN", _money(issqn["bc"]), False),
            (5.41, 5.09, "Alíquota Aplicada", _pct(issqn["aliq"]), False),
            (10.51, 5.09, "Retenção do ISSQN", issqn["retencaoLabel"], False),
            (15.62, 5.09, "ISSQN Apurado", _money(issqn["valor"]), True),
        ])
    else:
        bloco_ausente("TRIBUTAÇÃO MUNICIPAL (ISSQN) - OPERAÇÃO NÃO SUJEITA AO ISSQN")

    # ── Tributação Federal (Exceto CBS) ──
    fed = dados["fed"]
    bloco_titulo_com_campos("Tributação Federal (Exceto CBS)", [
        (5.41, 5.09, "IRRF", _money(fed["irrf"]) if fed["irrf"] else "-", False),
        (10.51, 5.09, "Contribuição Previdenciária - Retida", _money(fed["cp"]) if fed["cp"] else "-", False),
        (15.62, 5.09, "Contribuições Sociais - Retidas", _money(fed["contribSociaisRetidas"]) if fed["contribSociaisRetidas"] else "-", False),
    ])
    linha_campos([
        (0, 5.09, "PIS - Débito Apuração Própria", _money(fed["pisProprio"]) if fed["pisProprio"] else "-", False),
        (5.41, 5.09, "COFINS - Débito Apuração Própria", _money(fed["cofinsProprio"]) if fed["cofinsProprio"] else "-", False),
        (10.51, 10.19, "Descrição Contrib. Sociais - Retidas", fed["descRetLabel"], False),
    ])

    # ── Tributação IBS/CBS — colapsa se não houver dados ──
    ibs = dados["ibscbs"]
    if ibs["temDados"]:
        bloco_titulo_com_campos("Tributação IBS / CBS", [
            (5.41, 5.09, "CST / cClassTrib", f"{ibs['cst']} / {ibs['cClassTrib']}" if ibs["cst"] else "-", False),
            (10.51, 10.19, "Indicador de Operação / Código IBGE Incidência / Município Incidência / Sigla UF",
             " / ".join(p for p in [ibs["cIndOp"], ibs["codIncid"], ibs["municipioIncid"], ibs["ufIncid"]] if p) or "-", False),
        ])
        linha_campos([
            (0, 5.09, "Exclusões e Reduções da Base de Cálculo", "-", False),
            (5.41, 5.09, "Base de Cálculo Após Exclusões e Reduções", _money(ibs["vBC"]) or "-", False),
            (10.51, 5.09, "Red. Alíquota IBS / Red. Alíquota CBS",
             f"{_pct(ibs['pRedAliqUF']) or '-'} / {_pct(ibs['pRedAliqCBS']) or '-'}", False),
            (15.62, 5.09, "Alíquota - IBS UF / IBS Mun", f"{_pct(ibs['pIBSUF']) or '-'} / {_pct(ibs['pIBSMun']) or '-'}", False),
        ])
        linha_campos([
            (0, 5.09, "Alíq. Efetiva Municipal - IBS", _pct(ibs["pAliqEfetMun"]) or "-", False),
            (5.41, 5.09, "Valor Apurado Municipal - IBS", _money(ibs["vIBSMun"]) or "-", False),
            (10.51, 5.09, "Alíq. Efetiva Estadual - IBS", _pct(ibs["pAliqEfetUF"]) or "-", False),
            (15.62, 5.09, "Valor Apurado Estadual - IBS", _money(ibs["vIBSUF"]) or "-", False),
        ])
        linha_campos([
            (0, 5.09, "Valor Total Apurado - IBS", _money(ibs["vIBSTot"]) or "-", True),
            (5.41, 5.09, "Alíquota - CBS", _pct(ibs["pCBS"]) or "-", False),
            (10.51, 5.09, "Alíquota Efetiva - CBS", _pct(ibs["pAliqEfetCBS"]) or "-", False),
            (15.62, 5.09, "Valor Total Apurado - CBS", _money(ibs["vCBS"]) or "-", True),
        ])

    # ── Valor Total da NFS-e ──
    tot = dados["totais"]
    bloco_titulo_com_campos("Valor Total da NFS-e", [
        (5.41, 5.09, "Valor da Operação / Serviço", _money(tot["vServ"]), True),
        (10.51, 5.09, "Desconto Incondicionado", _money(tot["descIncond"]) if tot["descIncond"] else "-", False),
        (15.62, 5.09, "Desconto Condicionado", _money(tot["descCond"]) if tot["descCond"] else "-", False),
    ], alt=0.67)
    linha_campos([
        (0, 5.09, "Total das Retenções (ISSQN / Federais)", _money(tot["vTotalRet"]) or "-", False),
        (5.41, 5.09, "Valor Líquido da NFS-e", _money(tot["valorLiquido"]) or "-", True),
        (10.51, 5.09, "Total do IBS/CBS", ibs["totalIBSCBS"] or "-", False),
        (15.62, 5.09, "Valor Líquido da NFS-e + IBS/CBS", _money(tot["valorLiquidoIBSCBS"]) or "-", True),
    ], alt=0.67)
    # sombreado do último campo (norma 2.2.3) — redesenha por cima, sem moldura
    cel(15.62, cursor - 0.67, 5.09, 0.67, sombreado=True)
    cel(15.62, cursor - 0.67, 5.09, 0.67, "Valor Líquido da NFS-e + IBS/CBS",
       _money(tot["valorLiquidoIBSCBS"]) or "-", negrito=True)

    # ── Informações Complementares — altura elástica ──
    linhas_info = []
    if dados["infoCompl"]:
        linhas_info.extend(_wrap(dados["infoCompl"], 128))
    ta = dados["totaisAprox"]
    txt_aprox = (f"Totais Aproximados dos Tributos cfe. Lei nº 12.741/2012: "
                f"Federais: {_money(ta['fed']) if ta['fed'] else '-'}; "
                f"Estaduais: {_money(ta['est']) if ta['est'] else '-'}; "
                f"Municipais: {_money(ta['mun']) if ta['mun'] else '-'};")
    linhas_info.extend(_wrap(txt_aprox, 128))

    espaco_disponivel = (PAGINA_A - MARGEM) - cursor - 1.35   # reserva o canhoto
    altura_info = max(0.39 + 0.28 * len(linhas_info), min(espaco_disponivel, 0.39 + 0.28 * len(linhas_info)))
    altura_info = min(altura_info, max(espaco_disponivel, 0.7))
    alt_titulo_info = 0.42
    altura_info = max(altura_info, alt_titulo_info + 0.3)
    separador(cursor, grosso=True)
    cel(0, cursor, LARG, alt_titulo_info, "Informações Complementares", "", tam_label=7, label_caps=True, sombreado=True)
    ty = Y(cursor + alt_titulo_info) - 0.28 * cm
    c.setFont("Helvetica", 6.6)
    for linha_txt in linhas_info:
        if ty < Y(cursor + altura_info) + 0.1 * cm:
            break
        c.drawString(cm_(X0 + 0.06), ty, linha_txt)
        ty -= 0.26 * cm
    cursor += altura_info

    # ── Canhoto (protocolo p/ assinatura — mantém a caixa com divisórias) ──
    alt_canhoto = 0.75
    separador(cursor, grosso=True)
    c.setLineWidth(0.5)
    c.line(cm_(X0 + 5.09), Y(cursor), cm_(X0 + 5.09), Y(cursor + alt_canhoto))
    c.line(cm_(X0 + 10.51), Y(cursor), cm_(X0 + 10.51), Y(cursor + alt_canhoto))
    c.line(cm_(X0), Y(cursor + alt_canhoto), cm_(X1), Y(cursor + alt_canhoto))
    cel(0, cursor, 5.09, alt_canhoto, "Data de Cientificação:", "")
    cel(5.41, cursor, 5.09, alt_canhoto, "Identificação e Assinatura", "")
    cel(10.51, cursor, 9.79, alt_canhoto, "Nº NFS-e / Chave NFS-e", f"{dados['nNFSe']} / {dados['chave']}", tam_valor=6)

    # ── Marca d'água (nota cancelada/substituída) — cinza K35, diagonal, Arial 50pt+ ──
    if marca_dagua:
        c.saveState()
        c.setFillColor(CINZA_MARCA)
        c.setFont("Helvetica-Bold", 55)
        c.translate(cm_(PAGINA_L) / 2, cm_(PAGINA_A) / 2)
        c.rotate(45)
        c.drawCentredString(0, 0, marca_dagua)
        c.restoreState()

    # ── Borda externa (1pt) ──
    c.setLineWidth(1.0)
    c.setStrokeColor(PRETO)
    c.rect(cm_(MARGEM), cm_(MARGEM), cm_(PAGINA_L - 2 * MARGEM), cm_(PAGINA_A - 2 * MARGEM), fill=0, stroke=1)

    c.save()
    return buf.getvalue()


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Uso: python gerar_danfse.py <arquivo.xml> [saida.pdf]")
        raise SystemExit(1)
    xml_path = sys.argv[1]
    pdf_path = sys.argv[2] if len(sys.argv) > 2 else xml_path.rsplit(".", 1)[0] + ".pdf"
    with open(xml_path, "rb") as f:
        xml_bytes = f.read()
    pdf_bytes = gerar_pdf_danfse(xml_bytes)
    if pdf_bytes is None:
        print("XML não é uma NFS-e válida (pode ser um evento).")
        raise SystemExit(1)
    with open(pdf_path, "wb") as f:
        f.write(pdf_bytes)
    print(f"PDF gerado: {pdf_path}")

# -*- coding: utf-8 -*-
"""
Relatório de NFS-e em PDF, no formato do "Relatório NFSe" da extensão:
tabela detalhada das notas + resumo de totais, retenções e impostos.

Gera um PDF para Recebidas e outro para Emitidas (quando houver notas),
a partir dos XMLs já baixados, salvos na pasta da empresa:

    Relatorio_NFSe_Recebidas_EMPRESA_MM-AAAA.pdf
    Relatorio_NFSe_Emitidas_EMPRESA_MM-AAAA.pdf

Uso:
  python gerar_relatorio_pdf.py --empresa "DRA ANA" --competencia 2026-05
"""

import argparse
import re
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (BaseDocTemplate, Frame, PageTemplate,
                                Paragraph, Spacer, Table, TableStyle)

import baixar_nfse as nfse
import despacho
import leitor_nfse_excel as leitor

CINZA = colors.HexColor("#777777")
PRETO = colors.black
VERMELHO = colors.HexColor("#C00000")
LINHA = colors.HexColor("#BBBBBB")


# ------------------------------------------------------------------- parsing

def extrair_para_pdf(xml_path: Path) -> dict | None:
    """Extrai do XML os campos exibidos no relatório (usa os helpers do leitor)."""
    try:
        root = ET.parse(str(xml_path)).getroot()
    except ET.ParseError:
        return None
    inf = leitor.find_elem(root, "infNFSe")
    if inf is None:
        inf = root if root.tag.endswith("infNFSe") else None
        if inf is None:
            return None

    id_attr = inf.get("Id", "")
    chave = id_attr[3:] if id_attr.upper().startswith("NFS") else id_attr

    emit = leitor.find_elem(inf, "emit")
    dps = leitor.find_elem(inf, "DPS")
    idps = leitor.find_elem(dps, "infDPS") if dps is not None else None
    toma = leitor.find_elem(idps, "toma")
    vals_inf = leitor.find_elem(inf, "valores")
    vals_dps = leitor.find_elem(idps, "valores")
    serv = leitor.find_elem(idps, "serv")
    trib = leitor.find_elem(vals_dps, "trib")
    trib_mun = leitor.find_elem(trib, "tribMun")
    trib_fed = leitor.find_elem(trib, "tribFed")
    piscofins = leitor.find_elem(trib_fed, "piscofins")

    f, t = leitor.to_float, leitor.find_text
    v_serv = f(t(leitor.find_elem(vals_dps, "vServPrest"), "vServ"))
    if v_serv == 0.0:
        v_serv = f(t(vals_inf, "vBC"))

    return {
        "chave": chave,
        "nNFSe": t(inf, "nNFSe").lstrip("0") or "0",
        "dCompet": t(idps, "dCompet"),
        "dhEmi": t(idps, "dhEmi"),
        "prestador": t(emit, "xNome"),
        "tomador": t(toma, "xNome"),
        "xLocIncid": t(inf, "xLocIncid"),
        "cTribNac": t(leitor.find_elem(serv, "cServ"), "cTribNac"),
        "cNBS": t(leitor.find_elem(serv, "cServ"), "cNBS"),
        "tpRetISSQN": t(vals_inf, "tpRetISSQN") or t(trib_mun, "tpRetISSQN"),
        "pAliq": f(t(trib_mun, "pAliq")),
        "vISSQN": f(t(vals_inf, "vISSQN")),
        "tpRetPisCofins": t(piscofins, "tpRetPisCofins"),
        "CST": t(piscofins, "CST"),
        "vPis": f(t(piscofins, "vPis")),
        "vCofins": f(t(piscofins, "vCofins")),
        "vRetCP": f(t(trib_fed, "vRetCP")),
        "vRetIRRF": f(t(trib_fed, "vRetIRRF")),
        "vRetCSLL": f(t(trib_fed, "vRetCSLL")),
        "vDescIncond": f(t(vals_dps, "vDescIncond")),
        "vDescCond": f(t(vals_dps, "vDescCond")),
        "vServ": v_serv,
        "vLiq": f(t(vals_inf, "vLiq")),
    }


def chaves_canceladas_ou_substituidas(pasta_empresa: Path) -> set[str]:
    """
    Varre os eventos baixados (subpastas 'Eventos') e devolve as chaves das
    notas canceladas ou substituídas — que ficam FORA do relatório PDF.
    """
    excluidas = set()
    for arq in pasta_empresa.rglob("*.xml"):
        if "Eventos" not in arq.parts:
            continue
        try:
            texto = arq.read_text(encoding="utf-8", errors="replace")
        except OSError:
            continue
        m_cod = re.search(r"<e(\d{6})\b", texto) or re.search(r"<(?:tpEvento|cEvento)>(\d{6})", texto)
        cod = m_cod.group(1) if m_cod else ""
        if cod in ("101101", "101103", "105102", "105104", "105105"):
            m_ch = re.search(r"<chNFSe>(\d{50})</chNFSe>", texto)
            chave = m_ch.group(1) if m_ch else arq.stem.split("-")[0]
            excluidas.add(chave)
    return excluidas


# ---------------------------------------------------------------- formatação

def moeda(v: float) -> str:
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def num_ou_traco(v: float) -> str:
    return moeda(v) if v else "-"


def txt_ou_traco(s: str) -> str:
    return s if s else "-"


def data_br(iso: str) -> str:
    return leitor.fmt_date(iso) if iso else "-"


# ------------------------------------------------------------------ PDF

EST_TITULO = ParagraphStyle("titulo", fontName="Helvetica-Bold", fontSize=15, spaceAfter=2)
EST_SUB = ParagraphStyle("sub", fontName="Helvetica-Bold", fontSize=11, textColor=PRETO)
EST_INFO = ParagraphStyle("info", fontName="Helvetica", fontSize=8, textColor=CINZA)
EST_SECAO = ParagraphStyle("secao", fontName="Helvetica-Bold", fontSize=10, spaceBefore=10, spaceAfter=4)
EST_CEL = ParagraphStyle("cel", fontName="Helvetica", fontSize=7, leading=9)
EST_CEL_PEQ = ParagraphStyle("celpeq", fontName="Helvetica", fontSize=6.5, leading=8, textColor=CINZA)


def cab(label: str, tag_xml: str) -> Paragraph:
    return Paragraph(f"<b>{label}</b><br/><font size=5.5 color='#777777'>{tag_xml}</font>",
                     ParagraphStyle("h", fontName="Helvetica-Bold", fontSize=6.5, leading=8))


def linha_nota(reg: dict, tipo: str) -> list:
    comp = data_br(reg["dCompet"])
    emi = data_br(reg["dhEmi"])
    cor_comp = "#C00000" if comp[3:5] != emi[3:5] or comp[6:] != emi[6:] else "#000000"
    datas = Paragraph(f"<font color='{cor_comp}'>{comp}</font><br/>{emi}", EST_CEL)
    nome_parte = reg["prestador"] if tipo == "Recebidas" else reg["tomador"]
    local = Paragraph(
        f"{txt_ou_traco(reg['xLocIncid'])}<br/>"
        f"<font color='#555555'>{txt_ou_traco(reg['cTribNac'])}<br/>"
        f"{txt_ou_traco(reg['cNBS'])}</font>", EST_CEL)
    aliq = f"{reg['pAliq']:.2f}%".replace(".", ",") if reg["pAliq"] else "-"
    return [
        Paragraph("Recebida" if tipo == "Recebidas" else "Emitida", EST_CEL),
        Paragraph(reg["nNFSe"], EST_CEL),
        datas,
        Paragraph((nome_parte or "-")[:45], EST_CEL),
        local,
        moeda(reg["vServ"]),
        txt_ou_traco(reg["tpRetISSQN"]),
        aliq,
        num_ou_traco(reg["vISSQN"]),
        txt_ou_traco(reg["tpRetPisCofins"]),
        txt_ou_traco(reg["CST"]),
        num_ou_traco(reg["vPis"]),
        num_ou_traco(reg["vCofins"]),
        num_ou_traco(reg["vRetCP"]),
        num_ou_traco(reg["vRetIRRF"]),
        num_ou_traco(reg["vRetCSLL"]),
    ]


def card(titulo: str, tag_xml: str, valor: float, vermelho: bool = False) -> Table:
    cor = "#C00000" if vermelho else "#000000"
    p = Paragraph(
        f"<para align='center'><b><font size=7>{titulo}</font></b><br/>"
        f"<font size=6 color='#777777'>({tag_xml})</font><br/><br/>"
        f"<b><font size=9 color='{cor}'>R$ {moeda(valor)}</font></b></para>", EST_CEL)
    t = Table([[p]], colWidths=[44 * mm], rowHeights=[16 * mm])
    t.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.7, LINHA),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return t


def secao_cards(historia: list, titulo: str, cards: list) -> None:
    historia.append(Paragraph(titulo, EST_SECAO))
    historia.append(Table([cards], colWidths=[47 * mm] * len(cards),
                          style=TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")])))


def gerar_pdf(registros: list, tipo: str, empresa: dict, competencia: str,
              destino: Path) -> None:
    # paisagem: mais espaço horizontal para a coluna de VALOR e tributos
    doc = BaseDocTemplate(str(destino), pagesize=landscape(A4),
                          leftMargin=12 * mm, rightMargin=12 * mm,
                          topMargin=14 * mm, bottomMargin=12 * mm)

    def rodape(canvas, _doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(CINZA)
        canvas.drawString(12 * mm, 200 * mm, datetime.now().strftime("%d/%m/%Y, %H:%M"))
        canvas.drawCentredString(148.5 * mm, 200 * mm, "Relatório NFSe")
        canvas.drawRightString(285 * mm, 8 * mm, f"{canvas.getPageNumber()}")
        canvas.restoreState()

    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="f")
    doc.addPageTemplates([PageTemplate(id="pagina", frames=[frame], onPage=rodape)])

    rotulo = "Notas Recebidas" if tipo == "Recebidas" else "Notas Emitidas"
    historia = [
        Paragraph(f"Relatório de NFSe — {rotulo}", EST_TITULO),
        Paragraph(f"{empresa['nome']} · {nfse.somente_digitos(empresa['cnpj'])}", EST_SUB),
        Paragraph(f"Competência {competencia} | Total de {len(registros)} nota(s) "
                  f"fiscal(is) | Gerado em {datetime.now():%d/%m/%Y às %H:%M:%S}", EST_INFO),
        Spacer(1, 4 * mm),
    ]

    parte = "PRESTADOR" if tipo == "Recebidas" else "TOMADOR"
    cabecalho = [cab("TIPO", ""), cab("NOTA", "nNFSe"), cab("DATAS", "dCompet dhEmi"),
                 cab(parte, "xNome"), cab("LOCAL / SERVIÇO", "xLocIncid cTribNac cNBS"),
                 cab("VALOR", "vServ"),
                 cab("TPRET", "ISSQN"), cab("ALÍQ", "pAliq"), cab("ISSQN", "vISSQN"),
                 cab("TPRET", "PisCofins"), cab("CST", "PisCofins"), cab("PIS", "vPis"),
                 cab("COFINS", "vCofins"), cab("CP", "vRetCP"), cab("IRRF", "vRetIRRF"),
                 cab("CSLL", "vRetCSLL")]
    dados = [cabecalho] + [linha_nota(r, tipo) for r in registros]

    # linha de TOTAL no rodapé da tabela (soma do valor dos serviços)
    total_valor = sum(r["vServ"] for r in registros)
    estilo_total = ParagraphStyle("tot", fontName="Helvetica-Bold", fontSize=7.5)
    estilo_total_v = ParagraphStyle("totv", fontName="Helvetica-Bold",
                                    fontSize=7.5, alignment=2)
    linha_total = [""] * len(cabecalho)
    linha_total[3] = Paragraph(f"TOTAL ({len(registros)} nota(s))", estilo_total)
    linha_total[5] = Paragraph(moeda(total_valor), estilo_total_v)
    dados.append(linha_total)

    larguras = [16, 14, 20, 44, 25, 20, 12, 12, 16, 12, 10, 14, 16, 13, 14, 13]
    tabela = Table(dados, colWidths=[l * mm for l in larguras], repeatRows=1)
    tabela.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("FONTNAME", (5, 1), (-1, -1), "Helvetica"),
        ("LINEABOVE", (0, 0), (-1, 0), 1.2, PRETO),
        ("LINEBELOW", (0, 0), (-1, 0), 1.2, PRETO),
        ("LINEBELOW", (0, 1), (-1, -1), 0.4, LINHA),
        ("LINEABOVE", (0, -1), (-1, -1), 1.2, PRETO),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (5, 0), (-1, -1), "RIGHT"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
    ]))
    historia.append(tabela)
    historia.append(Spacer(1, 6 * mm))

    # ----- resumos (mesmas contas do leitor da Erika p/ retenções efetivas)
    tot_serv = sum(r["vServ"] for r in registros)
    iss_retido = sum(r["vISSQN"] for r in registros if r["tpRetISSQN"] == "2")
    iss_nao_retido = sum(r["vISSQN"] for r in registros if r["tpRetISSQN"] != "2")
    ret_cp = sum(r["vRetCP"] for r in registros)
    ret_irrf = sum(r["vRetIRRF"] for r in registros)
    ret_csll = sum(r["vRetCSLL"] for r in registros)
    descontos = sum(r["vDescIncond"] + r["vDescCond"] for r in registros)
    # PIS/COFINS só conta como retenção quando tpRetPisCofins indica retenção
    # (≠ "0"/vazio); senão é apuração própria do prestador (informativo, não
    # entra no total de retenções).
    pc_retido = lambda r: r["tpRetPisCofins"] not in ("", "0")
    ret_pis_cofins = sum(r["vPis"] + r["vCofins"] for r in registros if pc_retido(r))
    pis_cofins_proprio = sum(r["vPis"] + r["vCofins"] for r in registros if not pc_retido(r))
    tot_ret = iss_retido + ret_cp + ret_irrf + ret_csll + ret_pis_cofins
    tot_liq = sum(r["vLiq"] if r["vLiq"] else r["vServ"] for r in registros)

    secao_cards(historia, "RESUMO DE TOTAIS", [
        card("VALOR SERVIÇO", "vServ", tot_serv),
        card("VALOR TOTAL RETENÇÕES", "vTotalRet", tot_ret, vermelho=True),
        card("DESCONTOS", "vDescIncond + vDescCond", descontos, vermelho=True),
        card("VALOR LÍQUIDO", "vLiq", tot_liq)])
    secao_cards(historia, "RESUMO DE RETENÇÕES", [
        card("RETENÇÃO CP", "vRetCP", ret_cp, vermelho=True),
        card("RETENÇÃO IRRF", "vRetIRRF", ret_irrf, vermelho=True),
        card("RETENÇÃO CSLL", "vRetCSLL", ret_csll, vermelho=True),
        card("RETENÇÃO PIS/COFINS", "vPis + vCofins (retido)", ret_pis_cofins, vermelho=True),
        card("RETENÇÃO ISSQN", "vISSQN", iss_retido, vermelho=True)])
    secao_cards(historia, "RESUMO DE IMPOSTOS", [
        card("ISSQN NÃO RETIDO", "vISSQN", iss_nao_retido),
        card("PIS/COFINS OPERAÇÃO PRÓPRIA", "vPis + vCofins (não retido)", pis_cofins_proprio)])

    doc.build(historia)


# ----------------------------------------------------------------- principal

def main() -> int:
    parser = argparse.ArgumentParser(description="Gera relatório de NFS-e em PDF (formato da extensão).")
    parser.add_argument("--empresa", help="filtra empresas pelo nome (contém)")
    parser.add_argument("--competencia", required=True,
                        help="mês das notas (ex.: 2026-05 ou 05/2026)")
    args = parser.parse_args()

    config = nfse.carregar_json(nfse.ARQ_CONFIG, {})
    pasta_saida = Path(config.get("pasta_saida", nfse.PASTA_SCRIPT / "notas"))
    empresas = despacho.empresas_de(config)
    if args.empresa:
        empresas = [e for e in empresas if args.empresa.lower() in e["nome"].lower()]
    if not empresas:
        print("Nenhuma empresa encontrada na configuração com esse filtro.")
        return 1

    competencia = nfse.normalizar_competencia(args.competencia)
    sufixo = f"{competencia[5:7]}-{competencia[0:4]}"

    for empresa in empresas:
        nome = empresa["nome"]
        nfse.log(f"==== Relatório PDF: {nome} ({competencia}) ====")
        pasta_mes = pasta_saida / nome / competencia
        if not pasta_mes.exists():
            nfse.log(f"  Pasta não encontrada: {pasta_mes} (baixe as notas antes).")
            continue

        # só notas com Situação Normal entram no PDF
        excluidas = chaves_canceladas_ou_substituidas(pasta_saida / nome)

        for tipo in ("Recebidas", "Emitidas"):
            pasta_tipo = pasta_mes / tipo
            registros, puladas = [], 0
            if pasta_tipo.exists():
                for arq in sorted(pasta_tipo.glob("*.xml")):
                    reg = extrair_para_pdf(arq)
                    if not reg:
                        continue
                    if reg["chave"] in excluidas or arq.stem in excluidas:
                        puladas += 1
                        continue
                    registros.append(reg)
            if puladas:
                nfse.log(f"  {tipo}: {puladas} nota(s) cancelada(s)/substituída(s) fora do PDF.")
            if not registros:
                nfse.log(f"  {tipo}: nenhuma nota — PDF não gerado.")
                continue
            registros.sort(key=lambda r: (r["dhEmi"], r["nNFSe"].zfill(15)))
            destino = (pasta_saida / nome / f"{competencia}_Relatorios"
                       / f"Relatorio_NFSe_{tipo}_{nome}_{sufixo}.pdf")
            destino.parent.mkdir(parents=True, exist_ok=True)
            gerar_pdf(registros, tipo, empresa, competencia, destino)
            nfse.log(f"  {tipo}: {len(registros)} nota(s) -> {destino.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

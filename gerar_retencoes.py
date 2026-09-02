# -*- coding: utf-8 -*-
"""
Relatório de Retenções em Excel, a partir dos XMLs já baixados.

Usa o leitor de NFS-e da Erika (leitor_nfse_excel.py) para extrair as
retenções (ISS retido, INSS/CP, IRRF, CSLL, PIS/COFINS) de todos os XMLs
de uma empresa/mês e gera a planilha no mesmo local do Relatório Simples:

    pasta_saida/EMPRESA/Relatorio_Retencoes_EMPRESA_MM-AAAA.xlsx

Uso:
  python gerar_retencoes.py --empresa "DRA ANA" --competencia 2026-05
  python gerar_retencoes.py --empresa "TDS" --competencia 04/2026
  python gerar_retencoes.py --empresa "TDS"            # todos os meses baixados
"""

import argparse
import sys

# o leitor original imprime emojis; evita erro em consoles sem UTF-8
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import re
from datetime import datetime
from pathlib import Path

import baixar_nfse as nfse
import despacho
import leitor_nfse_excel as leitor

ARQ_RETENCOES_ESTADO = despacho.PASTA / "retencoes_estado.json"

# Eventos que mudam a situação da nota (mesmos códigos do Relatório Simples)
EVENTOS_CANCELAMENTO = {"101101", "101103"}
EVENTOS_SUBSTITUICAO = {"105102", "105104", "105105"}


def mapa_status_eventos(pasta_empresa: Path) -> dict:
    """
    Varre a pasta 'Eventos' de TODA a empresa (todos os meses) e monta um mapa
    chave da NFS-e -> status ('Cancelada' | 'Substituída'). Notas sem evento
    ficam como 'Normal'.
    """
    status: dict[str, str] = {}
    if not pasta_empresa.exists():
        return status
    for arq in pasta_empresa.rglob("*.xml"):
        if "Eventos" not in arq.parts:
            continue
        try:
            xml = arq.read_text(encoding="utf-8", errors="replace")
        except Exception:
            continue
        m_cod = re.search(r"<e(\d{6})\b", xml) or re.search(r"<(?:tpEvento|cEvento)>(\d{6})", xml)
        m_ch = re.search(r"<chNFSe>(\d{50})</chNFSe>", xml)
        if not m_cod or not m_ch:
            continue
        cod, ch = m_cod.group(1), m_ch.group(1)
        if cod in EVENTOS_CANCELAMENTO:
            status[ch] = "Cancelada"
        elif cod in EVENTOS_SUBSTITUICAO:
            status.setdefault(ch, "Substituída")
    return status


def coletar_registros(pasta_base: Path) -> list[dict]:
    """Lê recursivamente os XMLs de notas (ignora a pasta Eventos)."""
    registros = []
    erros = 0
    # ignora Eventos. Notas retroativas (emitidas no mês, mas com competência
    # de outro mês) ENTRAM no relatório do mês em que foram emitidas — é onde
    # o XML fica salvo (subpasta "XML retroativo" dentro do mês de emissão).
    arquivos = [p for p in sorted(pasta_base.rglob("*.xml"))
                if "Eventos" not in p.parts]
    for arq in arquivos:
        reg = leitor.parse_nfse(str(arq))
        if reg:
            # "papel" vem da pasta onde o XML está salvo (Emitidas/Recebidas),
            # a mesma classificação que o baixar_nfse.py já fez ao salvar —
            # não precisa reler o XML pra descobrir de novo.
            reg["papel"] = "Emitida" if "Emitidas" in arq.parts else "Recebida"
            registros.append(reg)
        else:
            erros += 1
    nfse.log(f"  {len(registros)} nota(s) lida(s), {erros} arquivo(s) ignorado(s).")
    return registros


def atualizar_estado_retencoes(nome: str, cnpj: str, competencia: str, total: float) -> None:
    """Guarda o total de retenção da empresa/mês em retencoes_estado.json —
    é o que alimenta o botão "Validação Retenções" e o aviso automático no
    Google Chat, sem precisar reabrir nenhum Excel depois. Sobrescreve a
    entrada anterior da mesma empresa (só importa o resultado mais recente)."""
    dados = nfse.carregar_json(ARQ_RETENCOES_ESTADO, {})
    dados[nome] = {
        "cnpj": cnpj,
        "competencia": competencia,
        "total": round(total, 2),
        "atualizado_em": datetime.now().isoformat(timespec="seconds"),
    }
    nfse.salvar_json(ARQ_RETENCOES_ESTADO, dados)


def main() -> int:
    parser = argparse.ArgumentParser(description="Gera o Relatório de Retenções (Excel) dos XMLs baixados.")
    parser.add_argument("--empresa", help="filtra empresas pelo nome (contém)")
    parser.add_argument("--competencia", help="só notas do mês (ex.: 2026-05 ou 05/2026)")
    args = parser.parse_args()

    config = nfse.carregar_json(nfse.ARQ_CONFIG, {})
    pasta_saida = Path(config.get("pasta_saida", nfse.PASTA_SCRIPT / "notas"))
    empresas = despacho.empresas_de(config)
    if args.empresa:
        empresas = [e for e in empresas if args.empresa.lower() in e["nome"].lower()]
    if not empresas:
        print("Nenhuma empresa encontrada na configuração com esse filtro.")
        return 1

    competencia = nfse.normalizar_competencia(args.competencia) if args.competencia else None

    for empresa in empresas:
        nome = empresa["nome"]
        nfse.log(f"==== Relatório de Retenções: {nome} ====")
        pasta_empresa = pasta_saida / nome
        pasta_base = pasta_empresa / competencia if competencia else pasta_empresa
        if not pasta_base.exists():
            nfse.log(f"  Pasta não encontrada: {pasta_base} (baixe as notas antes).")
            continue

        registros = coletar_registros(pasta_base)
        if not registros:
            nfse.log("  Nenhuma NFS-e válida encontrada.")
            continue

        # status (Cancelada/Substituída) a partir dos eventos da empresa
        status_map = mapa_status_eventos(pasta_empresa)
        for reg in registros:
            reg["situacao"] = status_map.get(reg["chNFSe"], "Normal")

        # mesma ordenação do leitor original (data de emissão + número)
        def sort_key(r):
            try:
                d = datetime.strptime(r["dhEmi"], "%d/%m/%Y")
            except ValueError:
                d = datetime.min
            try:
                n = int(r["nNFSe"])
            except ValueError:
                n = 0
            return (d, n)
        registros.sort(key=sort_key)

        if competencia:
            sufixo = f"{competencia[5:7]}-{competencia[0:4]}"   # MM-AAAA
            pasta_relatorios = pasta_empresa / f"{competencia}_Relatorios"
        else:
            sufixo = "TODOS-OS-MESES"
            pasta_relatorios = pasta_empresa / "Relatorios"
        pasta_relatorios.mkdir(parents=True, exist_ok=True)
        destino = pasta_relatorios / f"Relatorio_Retencoes_{nome}_{sufixo}.xlsx"
        leitor.gerar_excel(registros, str(destino))

        total_ret = sum(r["vTotalRet"] for r in registros)
        nfse.log(f"  Total de retenções no período: R$ {total_ret:,.2f}")
        nfse.log(f"  Planilha salva em: {destino}")
        if competencia:
            atualizar_estado_retencoes(nome, empresa.get("cnpj", ""), competencia, total_ret)
    return 0


if __name__ == "__main__":
    sys.exit(main())

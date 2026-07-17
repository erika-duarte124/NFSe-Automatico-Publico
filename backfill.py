# -*- coding: utf-8 -*-
"""
Busca do histórico inicial — a escolha feita na Tela 3 do assistente
("todo o histórico" ou "a partir de um mês específico"). Roda UMA VEZ,
disparada em segundo plano logo que o cadastro é concluído, sem travar a
tela do assistente.

- "mes_especifico": baixa + gera relatório de cada mês, do escolhido até o
  mês atual, para cada empresa.
- "completo": baixa o histórico inteiro (sem filtro de competência), depois
  descobre quais meses foram encontrados e gera relatório de cada um.

Tudo registrado em backfill.log. Notifica o Windows ao concluir.
"""

import json
import re
import subprocess
from datetime import date
from pathlib import Path

import despacho

PASTA = despacho.PASTA
LOG = PASTA / "backfill.log"
ARQ_CONFIG = PASTA / "config.json"

TIMEOUT_PASSO_SEGUNDOS = 30 * 60
PADRAO_COMPETENCIA = re.compile(r"^\d{4}-\d{2}$")


def registrar(texto: str) -> None:
    from datetime import datetime
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now():%d/%m/%Y %H:%M:%S}] {texto}\n")


def meses_entre(desde: str, ate: date) -> list[str]:
    ano, mes = (int(p) for p in desde.split("-"))
    meses = []
    while (ano, mes) <= (ate.year, ate.month):
        meses.append(f"{ano:04d}-{mes:02d}")
        mes += 1
        if mes > 12:
            mes = 1
            ano += 1
    return meses


def rodar_passo(nome_passo: str, args: list[str]) -> bool:
    comando = despacho.comando_base() + args
    registrar(f"  >>> {nome_passo}")
    with open(LOG, "a", encoding="utf-8") as saida:
        try:
            r = subprocess.run(comando, stdout=saida, stderr=subprocess.STDOUT,
                               cwd=str(PASTA), timeout=TIMEOUT_PASSO_SEGUNDOS)
            ok = r.returncode == 0
        except subprocess.TimeoutExpired:
            ok = False
            registrar(f"  <<< FALHOU (excedeu {TIMEOUT_PASSO_SEGUNDOS // 60} min de limite): {nome_passo}")
            return False
    if not ok:
        registrar(f"  <<< FALHOU: {nome_passo}")
    return ok


def gerar_relatorios(empresa: str, competencia: str) -> None:
    rodar_passo(f"{empresa} {competencia} - Relatorio Simples",
                ["gerar_relatorio", "--empresa", empresa, "--competencia", competencia])
    rodar_passo(f"{empresa} {competencia} - Relatorio Retencoes",
                ["gerar_retencoes", "--empresa", empresa, "--competencia", competencia])
    rodar_passo(f"{empresa} {competencia} - Relatorio PDF",
                ["gerar_relatorio_pdf", "--empresa", empresa, "--competencia", competencia])


def competencias_existentes(pasta_saida: str, empresa: str) -> list[str]:
    pasta_empresa = Path(pasta_saida) / empresa
    if not pasta_empresa.exists():
        return []
    return sorted(p.name for p in pasta_empresa.iterdir() if p.is_dir() and PADRAO_COMPETENCIA.match(p.name))


def main() -> None:
    config = json.loads(ARQ_CONFIG.read_text(encoding="utf-8"))
    empresas = [e["nome"] for grupo in config.get("grupos", []) for e in grupo.get("empresas", [])]
    pasta_saida = config.get("pasta_saida")
    periodo = config.get("periodo_inicial", {"tipo": "completo"})

    registrar(f"======== BUSCA DE HISTÓRICO INICIAL — {periodo} ========")

    if periodo.get("tipo") == "mes_especifico" and periodo.get("desde"):
        meses = meses_entre(periodo["desde"], date.today())
        for empresa in empresas:
            for competencia in meses:
                if not rodar_passo(f"{empresa} {competencia} - baixar XML+PDF",
                                   ["baixar_nfse", "--empresa", empresa, "--competencia", competencia]):
                    continue  # sem download, não adianta tentar relatorio desse mes
                rodar_passo(f"{empresa} {competencia} - completar PDFs",
                            ["baixar_nfse", "--empresa", empresa, "--completar-pdf", "--competencia", competencia])
                gerar_relatorios(empresa, competencia)
    else:
        for empresa in empresas:
            if not rodar_passo(f"{empresa} - baixar XML+PDF (histórico completo)",
                               ["baixar_nfse", "--empresa", empresa]):
                continue
            rodar_passo(f"{empresa} - completar PDFs",
                        ["baixar_nfse", "--empresa", empresa, "--completar-pdf"])
            for competencia in competencias_existentes(pasta_saida, empresa):
                gerar_relatorios(empresa, competencia)

    registrar("======== BUSCA DE HISTÓRICO INICIAL CONCLUÍDA ========")
    try:
        from win11toast import toast
        toast("NFS-e Automático", "Busca do histórico inicial concluída — confira os relatórios.", duration="short")
    except Exception as e:
        registrar(f"  (Não foi possível mostrar a notificação do Windows: {e})")


if __name__ == "__main__":
    main()

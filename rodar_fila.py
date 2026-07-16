# -*- coding: utf-8 -*-
"""
Executor de fila: roda os passos de download/relatório UM DE CADA VEZ,
em processo independente, registrando tudo em fila.log.

Para CADA empresa e competência, executa o pipeline completo:
  1) baixar XML+PDF   2) completar PDFs   3) Relatório Simples (Excel)
  4) Relatório Retenções (Excel)   5) Relatório PDF

Uso:
  python rodar_fila.py 2026-06          # todas as empresas, junho/2026
  python rodar_fila.py 06/2026          # mesmo mês (formato alternativo)
  python rodar_fila.py 2026-06 ANTARES  # só empresas cujo nome contenha "ANTARES"
  python rodar_fila.py                  # usa o mês ATUAL, todas as empresas
"""

import json
import re
import subprocess
import sys
from datetime import date, datetime
from pathlib import Path

import despacho

PASTA = despacho.PASTA
LOG = PASTA / "fila.log"
ARQ_CONFIG = PASTA / "config.json"


def carregar_empresas() -> list[str]:
    config = json.loads(ARQ_CONFIG.read_text(encoding="utf-8"))
    return [e["nome"] for e in config.get("empresas", [])]


EMPRESAS = carregar_empresas()


def normalizar_competencia(texto: str) -> str:
    """Aceita '2026-06', '06/2026' ou '2026/06' e devolve sempre 'AAAA-MM'."""
    texto = texto.strip()
    m = re.fullmatch(r"(\d{4})[-/](\d{1,2})", texto)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"
    m = re.fullmatch(r"(\d{1,2})[-/](\d{4})", texto)
    if m:
        return f"{m.group(2)}-{int(m.group(1)):02d}"
    raise SystemExit(f"Competência inválida: {texto!r} (use AAAA-MM, ex.: 2026-06)")


def montar_passos(empresas, competencia):
    base = despacho.comando_base()
    passos = []
    for emp in empresas:
        passos += [
            (f"{emp} {competencia} - baixar XML+PDF",      base + ["baixar_nfse", "--empresa", emp, "--competencia", competencia]),
            (f"{emp} {competencia} - completar PDFs",      base + ["baixar_nfse", "--empresa", emp, "--completar-pdf", "--competencia", competencia]),
            (f"{emp} {competencia} - Relatorio Simples",   base + ["gerar_relatorio", "--empresa", emp, "--competencia", competencia]),
            (f"{emp} {competencia} - Relatorio Retencoes", base + ["gerar_retencoes", "--empresa", emp, "--competencia", competencia]),
            (f"{emp} {competencia} - Relatorio PDF",       base + ["gerar_relatorio_pdf", "--empresa", emp, "--competencia", competencia]),
        ]
    return passos


def registrar(texto: str) -> None:
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now():%d/%m %H:%M:%S}] {texto}\n")


def main() -> None:
    args = sys.argv[1:]
    if args:
        competencia = normalizar_competencia(args[0])
        filtro = args[1] if len(args) > 1 else None
    else:
        hoje = date.today()
        competencia = f"{hoje.year:04d}-{hoje.month:02d}"
        filtro = None

    empresas = [e for e in EMPRESAS if filtro is None or filtro.lower() in e.lower()]
    if not empresas:
        raise SystemExit(f"Nenhuma empresa da fila contém {filtro!r}.")

    passos = montar_passos(empresas, competencia)
    registrar(f"================ FILA INICIADA — competência {competencia} "
              f"({len(empresas)} empresa(s)) ================")
    falhas = []
    for titulo, comando in passos:
        registrar(f">>> INICIANDO: {titulo}")
        with open(LOG, "a", encoding="utf-8") as saida:
            r = subprocess.run(comando, stdout=saida, stderr=subprocess.STDOUT, cwd=str(PASTA))
        if r.returncode == 0:
            registrar(f"<<< CONCLUIDO: {titulo}")
        else:
            falhas.append(titulo)
            registrar(f"<<< FALHOU (codigo {r.returncode}): {titulo} — seguindo para o próximo passo")
    if falhas:
        registrar(f"================ FILA TERMINADA COM {len(falhas)} FALHA(S): {falhas} ================")
    else:
        registrar("================ FILA TERMINADA COM SUCESSO ================")


if __name__ == "__main__":
    main()

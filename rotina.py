# -*- coding: utf-8 -*-
"""
Rotina agendada de retirada de NFS-e.

Chamada UMA VEZ POR DIA pelo Agendador de Tarefas do Windows, sempre no
mesmo horário (definido durante a configuração inicial). O próprio script
decide, consultando o "agendamento" do config.json, se hoje é dia de rodar
algo — e o quê.

Frequências suportadas (configuráveis em config.json -> "agendamento"),
podendo ativar até 2 ao mesmo tempo:

  - MENSAL: dia fixo do mês (ex.: todo dia 1). Fecha o mês ANTERIOR
    completo. Se o PC estiver desligado no dia certo, roda na próxima vez
    que ligar (recuperação automática de atraso).

  - SEMANAL: um dia da semana fixo (ex.: toda quarta-feira), toda semana.
    Baixa/gera relatórios do mês ATUAL (ainda em aberto).

  - QUINZENAL: um dia da semana fixo, mas só a cada 14 dias (controlado
    por "ultima_execucao_quinzenal" no rotina_estado.json).

Se hoje for dia de fechamento mensal, semanal/quinzenal não rodam nesse
dia (evita duplicar trabalho).

Tudo é registrado em rotina.log. Estado de controle em rotina_estado.json.
"""

import argparse
import json
import subprocess
from datetime import date, timedelta
from pathlib import Path

import despacho

PASTA = despacho.PASTA
LOG = PASTA / "rotina.log"
ARQ_ESTADO = PASTA / "rotina_estado.json"
ARQ_CONFIG = PASTA / "config.json"


def carregar_config() -> dict:
    return json.loads(ARQ_CONFIG.read_text(encoding="utf-8"))


def registrar(texto: str) -> None:
    from datetime import datetime
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now():%d/%m/%Y %H:%M:%S}] {texto}\n")


def mes_anterior(d: date) -> str:
    primeiro = d.replace(day=1)
    anterior = primeiro - timedelta(days=1)
    return f"{anterior.year:04d}-{anterior.month:02d}"


def pipeline(filtro_empresa: str, competencia: str) -> int:
    """Roda baixar -> completar PDFs -> 3 relatórios, p/ 1 empresa/mês."""
    base = despacho.comando_base()
    passos = [
        ("baixar XML+PDF",      base + ["baixar_nfse", "--empresa", filtro_empresa, "--competencia", competencia]),
        ("completar PDFs",      base + ["baixar_nfse", "--empresa", filtro_empresa, "--completar-pdf", "--competencia", competencia]),
        ("Relatorio Simples",   base + ["gerar_relatorio", "--empresa", filtro_empresa, "--competencia", competencia]),
        ("Relatorio Retencoes", base + ["gerar_retencoes", "--empresa", filtro_empresa, "--competencia", competencia]),
        ("Relatorio PDF",       base + ["gerar_relatorio_pdf", "--empresa", filtro_empresa, "--competencia", competencia]),
    ]
    falhas = 0
    for titulo, comando in passos:
        registrar(f"  >>> {filtro_empresa} {competencia} - {titulo}")
        with open(LOG, "a", encoding="utf-8") as saida:
            r = subprocess.run(comando, stdout=saida, stderr=subprocess.STDOUT, cwd=str(PASTA))
        if r.returncode != 0:
            falhas += 1
            registrar(f"  <<< FALHOU: {filtro_empresa} {competencia} - {titulo}")
    return falhas


def executar_para_todas(empresas: list[str], competencia: str, rotulo: str) -> None:
    registrar(f"======== {rotulo} — competência {competencia} ========")
    total_falhas = 0
    for empresa in empresas:
        total_falhas += pipeline(empresa, competencia)
    if total_falhas:
        registrar(f"======== {rotulo} TERMINADO COM {total_falhas} FALHA(S) ========")
    else:
        registrar(f"======== {rotulo} TERMINADO COM SUCESSO ========")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--modo", choices=["auto", "mensal", "semanal", "quinzenal"], default="auto",
                        help="qual verificação rodar (a tarefa do Agendador sempre chama sem --modo, ou seja, 'auto')")
    args = parser.parse_args()

    config = carregar_config()
    empresas = [e["nome"] for e in config.get("empresas", [])]
    agenda = config.get("agendamento", {})

    hoje = date.today()
    estado = json.loads(ARQ_ESTADO.read_text(encoding="utf-8")) if ARQ_ESTADO.exists() else {}

    # ---- MENSAL: dia fixo do mês, fecha o mês anterior, com recuperação de atraso
    cfg_mensal = agenda.get("mensal", {})
    mes_fechamento = mes_anterior(hoje)
    devida = date(hoje.year, hoje.month, cfg_mensal.get("dia_mes", 1))
    eh_dia_de_fechamento = (cfg_mensal.get("ativo")
                            and hoje >= devida
                            and estado.get("ultimo_fechamento_mensal") != mes_fechamento)

    if args.modo in ("auto", "mensal") and eh_dia_de_fechamento:
        if hoje == devida:
            registrar(f"Hoje é o dia do fechamento mensal de {mes_fechamento}.")
        else:
            registrar(f"Fechamento mensal de {mes_fechamento} estava atrasado "
                      f"(devido em {devida:%d/%m}); executando agora.")
        executar_para_todas(empresas, mes_fechamento, "FECHAMENTO MENSAL")
        estado["ultimo_fechamento_mensal"] = mes_fechamento
        ARQ_ESTADO.write_text(json.dumps(estado, indent=2), encoding="utf-8")
        return

    if args.modo == "auto" and eh_dia_de_fechamento:
        return  # já tratado acima; nunca deveria chegar aqui

    competencia_atual = f"{hoje.year:04d}-{hoje.month:02d}"

    # ---- SEMANAL: um dia da semana, toda semana
    cfg_semanal = agenda.get("semanal", {})
    if args.modo in ("auto", "semanal") and cfg_semanal.get("ativo"):
        if cfg_mensal.get("ativo") and eh_dia_de_fechamento:
            registrar("Hoje também é dia de fechamento mensal; pulando a retirada semanal.")
        elif hoje.weekday() == cfg_semanal.get("dia_semana"):
            registrar("Hoje é dia de retirada semanal do mês atual.")
            executar_para_todas(empresas, competencia_atual, "RETIRADA SEMANAL")
            return

    # ---- QUINZENAL: um dia da semana, só a cada 14 dias
    cfg_quinzenal = agenda.get("quinzenal", {})
    if args.modo in ("auto", "quinzenal") and cfg_quinzenal.get("ativo"):
        if cfg_mensal.get("ativo") and eh_dia_de_fechamento:
            registrar("Hoje também é dia de fechamento mensal; pulando a retirada quinzenal.")
            return
        ultima = estado.get("ultima_execucao_quinzenal")
        ultima_data = date.fromisoformat(ultima) if ultima else None
        due_por_data = ultima_data is None or (hoje - ultima_data).days >= 14
        if hoje.weekday() == cfg_quinzenal.get("dia_semana") and due_por_data:
            registrar("Hoje é dia de retirada quinzenal do mês atual.")
            executar_para_todas(empresas, competencia_atual, "RETIRADA QUINZENAL")
            estado["ultima_execucao_quinzenal"] = hoje.isoformat()
            ARQ_ESTADO.write_text(json.dumps(estado, indent=2), encoding="utf-8")
            return

    registrar("Nada agendado para hoje.")


if __name__ == "__main__":
    main()

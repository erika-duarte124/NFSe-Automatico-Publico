# -*- coding: utf-8 -*-
"""
Rotina agendada de retirada de NFS-e.

Chamada UMA VEZ POR DIA por grupo, pelo Agendador de Tarefas do Windows
(--grupo "Nome do Grupo" --modo mensal/semanal/quinzenal), sempre no mesmo
horário (definido durante a configuração inicial). O próprio script decide,
consultando o "agendamento" daquele grupo no config.json, se hoje é dia de
rodar algo — e o quê. Cada grupo tem sua própria agenda e seu próprio
estado de controle (rotina_estado.json -> "grupos" -> nome do grupo).

Frequências suportadas (configuráveis em config.json -> "agendamento"),
podendo ativar até 2 ao mesmo tempo:

  - MENSAL: dia fixo do mês (ex.: todo dia 1). Fecha o mês ANTERIOR
    completo. Se o PC estiver desligado no dia certo, roda na próxima vez
    que ligar (recuperação automática de atraso).

  - SEMANAL: um dia da semana fixo (ex.: toda quarta-feira), toda semana.
    Baixa/gera relatórios do mês ATUAL (ainda em aberto).

  - QUINZENAL: um dia da semana fixo, mas só a cada 14 dias (controlado
    por "ultima_execucao_quinzenal" no rotina_estado.json).

Mensal fecha o mês ANTERIOR e semanal/quinzenal tratam do mês ATUAL —
competências diferentes, então rodar os dois no mesmo dia nunca duplica
trabalho; cada um roda no seu próprio horário, independente do outro. Se os
dois calharem de coincidir (mesmo grupo ou grupos diferentes, e mesmo a
retirada manual "Rodar agora"), um espera o outro terminar antes de começar
(mutex ÚNICO pra máquina inteira) — nunca duas retiradas rodam ao mesmo
tempo neste PC, pra não sobrecarregar o Portal Nacional com pedidos
simultâneos (na prática já travou a resposta quando isso aconteceu, mesmo
entre empresas diferentes).

Cada passo (baixar, completar PDFs, cada relatório) tem um limite de tempo
de segurança (TIMEOUT_PASSO_SEGUNDOS) — se uma empresa travar por qualquer
motivo, só aquele passo é cancelado e registrado como falha; a fila segue
para a próxima empresa em vez de travar indefinidamente.

Ao final de cada execução, uma notificação do Windows resume o resultado
(quantas empresas OK, quantas com falha).

Tudo é registrado em rotina.log. Estado de controle em rotina_estado.json.
"""

import argparse
import json
import re
import subprocess
from datetime import date, datetime, timedelta
from pathlib import Path

import requests
import win32event

import despacho

LIMITE_HISTORICO = 500  # execuções mais antigas que isso são descartadas

PASTA = despacho.PASTA
LOG = PASTA / "rotina.log"
ARQ_ESTADO = PASTA / "rotina_estado.json"
ARQ_CONFIG = PASTA / "config.json"
ARQ_ULTIMA_EXECUCAO = PASTA / "ultima_execucao.json"
ARQ_RETENCOES_ESTADO = PASTA / "retencoes_estado.json"

TIMEOUT_PASSO_SEGUNDOS = 30 * 60  # limite de segurança por passo (folgado, p/ empresas com muitas notas)


def carregar_config() -> dict:
    return json.loads(ARQ_CONFIG.read_text(encoding="utf-8"))


def registrar(texto: str) -> None:
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now():%d/%m/%Y %H:%M:%S}] {texto}\n")


def _com_mutex_estado(func) -> None:
    """Executa func() protegido por um mutex nomeado do Windows — evita que
    duas execuções concorrentes (grupos ou frequências diferentes rodando ao
    mesmo tempo) leiam e gravem rotina_estado.json/ultima_execucao.json ao
    mesmo tempo, uma sobrescrevendo a atualização da outra. Se o processo
    morrer com o mutex em mãos, o Windows libera sozinho pro próximo — não
    precisa de limpeza manual."""
    mutex = win32event.CreateMutex(None, False, "Global\\NFSeAutomatico_EstadoLock")
    win32event.WaitForSingleObject(mutex, win32event.INFINITE)
    try:
        func()
    finally:
        win32event.ReleaseMutex(mutex)


def com_lock_execucao_global(func):
    """Executa func() protegido por um mutex ÚNICO pra máquina inteira —
    garante que só UMA retirada de notas roda por vez neste PC, seja
    agendada (mensal/semanal/quinzenal, de qualquer grupo) ou manual
    ("Rodar agora"). Se outra já estiver rodando, esta espera terminar em
    vez de ser pulada ou rodar em paralelo — evita pedidos simultâneos ao
    Portal Nacional, que na prática já travou a resposta quando isso
    aconteceu, mesmo entre empresas diferentes."""
    mutex = win32event.CreateMutex(None, False, "Global\\NFSeAutomatico_PipelineLock")
    if win32event.WaitForSingleObject(mutex, 0) not in (win32event.WAIT_OBJECT_0, win32event.WAIT_ABANDONED):
        registrar("Outra retirada já está rodando nesta máquina — aguardando terminar...")
        win32event.WaitForSingleObject(mutex, win32event.INFINITE)
    try:
        return func()
    finally:
        win32event.ReleaseMutex(mutex)


def salvar_estado_grupo(grupo: str, dados: dict) -> None:
    """Grava só os dados DESSE grupo em rotina_estado.json, relendo o
    arquivo na hora (protegido por mutex) — não usa uma cópia antiga lida
    no início da execução, que poderia já estar desatualizada."""
    def _fazer():
        atual = json.loads(ARQ_ESTADO.read_text(encoding="utf-8")) if ARQ_ESTADO.exists() else {}
        atual.setdefault("grupos", {})[grupo] = dados
        ARQ_ESTADO.write_text(json.dumps(atual, indent=2, ensure_ascii=False), encoding="utf-8")
    _com_mutex_estado(_fazer)


def mes_anterior(d: date) -> str:
    primeiro = d.replace(day=1)
    anterior = primeiro - timedelta(days=1)
    return f"{anterior.year:04d}-{anterior.month:02d}"


def pipeline(filtro_empresa: str, competencia: str, gerar_retencoes_e_pdf: bool = True) -> int:
    """Roda baixar -> completar PDFs -> relatórios, p/ 1 empresa/mês.

    Cada passo tem um limite de tempo (TIMEOUT_PASSO_SEGUNDOS): se travar,
    o passo é cancelado e registrado como falha. Se o PRIMEIRO passo
    (baixar XML+PDF) falhar, os demais passos dessa empresa são pulados —
    não faz sentido gerar relatório sem ter conseguido baixar nada, e isso
    limita o pior caso (uma empresa com problema) a 1 timeout, não 5.

    gerar_retencoes_e_pdf=False pula os relatórios de Retenções e PDF —
    usado pelas execuções SEMANAL/QUINZENAL do agendamento automático, que
    tratam do mês ainda em aberto (Retenções e PDF só fazem sentido pro
    fechamento MENSAL, mês completo). "Rodar agora" e a busca de histórico
    inicial continuam chamando com o padrão (True)."""
    base = despacho.comando_base()
    passos = [
        ("baixar XML+PDF",      base + ["baixar_nfse", "--empresa", filtro_empresa, "--competencia", competencia]),
        ("completar PDFs",      base + ["baixar_nfse", "--empresa", filtro_empresa, "--completar-pdf", "--competencia", competencia]),
        ("Relatorio Simples",   base + ["gerar_relatorio", "--empresa", filtro_empresa, "--competencia", competencia]),
    ]
    if gerar_retencoes_e_pdf:
        passos += [
            ("Relatorio Retencoes", base + ["gerar_retencoes", "--empresa", filtro_empresa, "--competencia", competencia]),
            ("Relatorio PDF",       base + ["gerar_relatorio_pdf", "--empresa", filtro_empresa, "--competencia", competencia]),
        ]
    falhas = 0
    for indice, (titulo, comando) in enumerate(passos):
        registrar(f"  >>> {filtro_empresa} {competencia} - {titulo}")
        with open(LOG, "a", encoding="utf-8") as saida:
            try:
                r = subprocess.run(comando, stdout=saida, stderr=subprocess.STDOUT,
                                   cwd=str(PASTA), timeout=TIMEOUT_PASSO_SEGUNDOS)
                codigo = r.returncode
            except subprocess.TimeoutExpired:
                codigo = None
        if codigo == 0:
            continue
        falhas += 1
        if codigo is None:
            registrar(f"  <<< FALHOU (excedeu {TIMEOUT_PASSO_SEGUNDOS // 60} min de limite de segurança): "
                      f"{filtro_empresa} {competencia} - {titulo}")
        else:
            registrar(f"  <<< FALHOU: {filtro_empresa} {competencia} - {titulo}")
        if indice == 0:
            falhas += len(passos) - 1  # conta os passos pulados como falha também
            registrar(f"  <<< Pulando os demais passos de {filtro_empresa} "
                      f"(download falhou — sem dados pra gerar relatório)")
            break
    return falhas


def notificar(titulo: str, mensagem: str) -> None:
    try:
        from win11toast import toast
        toast(titulo, mensagem, duration="short")
    except Exception as e:
        registrar(f"  (Não foi possível mostrar a notificação do Windows: {e})")


def enviar_alerta_chat(mensagem: str) -> None:
    """Manda `mensagem` pro webhook do Google Chat configurado no
    assistente (aba "Notificações..."), se houver um cadastrado. Silencioso
    se não tiver webhook configurado; registra no log se o envio falhar,
    mas nunca interrompe a rotina por causa disso."""
    webhook = carregar_config().get("webhook_chat", "").strip()
    if not webhook:
        return
    try:
        r = requests.post(webhook, json={"text": mensagem}, timeout=15)
        if r.status_code >= 300:
            registrar(f"  (Aviso no Google Chat falhou: HTTP {r.status_code})")
    except Exception as e:
        registrar(f"  (Não foi possível enviar o aviso no Google Chat: {e})")


def avisar(titulo: str, mensagem: str) -> None:
    """Notificação do Windows + aviso no Google Chat (se configurado) —
    ponto único usado por rotina.py e executar_agora.py ao final de cada
    execução."""
    notificar(titulo, mensagem)
    enviar_alerta_chat(f"*{titulo}*\n{mensagem}")


def formatar_cnpj(cnpj: str) -> str:
    n = "".join(c for c in (cnpj or "") if c.isdigit())
    if len(n) != 14:
        return cnpj or ""
    return f"{n[:2]}.{n[2:5]}.{n[5:8]}/{n[8:12]}-{n[12:]}"


def avisar_retencoes_destacadas(grupo: str, empresas: list[str], competencia: str) -> None:
    """Depois de um FECHAMENTO MENSAL, olha o retencoes_estado.json (gravado
    pelo gerar_retencoes.py de cada empresa) e, se alguma empresa do grupo
    teve retenção nesse mês, manda um aviso à parte no Google Chat — só
    quando há webhook configurado, mesma regra de enviar_alerta_chat."""
    estado = json.loads(ARQ_RETENCOES_ESTADO.read_text(encoding="utf-8")) if ARQ_RETENCOES_ESTADO.exists() else {}
    config = carregar_config()
    grupo_cfg = next((g for g in config.get("grupos", []) if g.get("nome") == grupo), {})
    cnpj_por_nome = {e["nome"]: e.get("cnpj", "") for e in grupo_cfg.get("empresas", [])}

    destacadas = [nome for nome in empresas
                  if estado.get(nome, {}).get("competencia") == competencia
                  and estado.get(nome, {}).get("total", 0) > 0]
    if not destacadas:
        return

    _NL = chr(10)
    linhas = _NL.join(f"{nome} - {formatar_cnpj(cnpj_por_nome.get(nome, ''))}" for nome in destacadas)
    enviar_alerta_chat(
        f"[{grupo}] Empresas com Retenções destacadas:{_NL}{linhas}{_NL}*Necessário conferência"
    )


def checar_notas_retroativas_emitidas(pasta_saida: Path, empresas: list[str],
                                       competencia: str, desde: datetime) -> list[dict]:
    """Procura, entre os XMLs baixados NESTA execução, notas EMITIDAS pela
    empresa (não Recebidas) com competência retroativa — emitidas no mês
    corrente mas referentes a um mês anterior já fechado. Essas notas ficam
    salvas em Emitidas/XML retroativo/ dentro do mês de emissão."""
    achados = []
    for nome in empresas:
        pasta = pasta_saida / nome / competencia / "Emitidas" / "XML retroativo"
        if not pasta.exists():
            continue
        for arq in pasta.glob("*.xml"):
            if arq.stat().st_mtime < desde.timestamp():
                continue
            try:
                texto = arq.read_text(encoding="utf-8", errors="replace")
            except OSError:
                continue
            m = re.search(r"<dCompet>(\d{4})-(\d{2})", texto)
            comp_real = f"{m.group(1)}-{m.group(2)}" if m else "desconhecida"
            achados.append({"empresa": nome, "chave": arq.stem, "competencia_real": comp_real})
    return achados


def avisar_notas_retroativas(grupo: str, empresas: list[str], competencia: str, desde: datetime) -> None:
    """Retirada SEMANAL/QUINZENAL: avisa à parte no Google Chat se apareceu
    nota EMITIDA com competência retroativa — situação que os relatórios
    normais não destacam sozinha."""
    config = carregar_config()
    pasta_saida = Path(config.get("pasta_saida", despacho.PASTA / "notas"))
    achados = checar_notas_retroativas_emitidas(pasta_saida, empresas, competencia, desde)
    if not achados:
        return
    _NL = chr(10)
    linhas = _NL.join(
        f"• {a['empresa']} — nota {a['chave']} (competência real: {a['competencia_real']})"
        for a in achados
    )
    enviar_alerta_chat(
        f"[{grupo}] Nota EMITIDA com data retroativa:{_NL}"
        f"Competência da retirada: {competencia}{_NL}"
        f"{len(achados)} nota(s) emitida(s) fora da competência corrente:{_NL}"
        f"{linhas}"
    )


def executar_para_todas(grupo: str, empresas: list[str], competencia: str, rotulo: str) -> None:
    """Roda o pipeline pra todas as empresas do grupo. Protegido pelo mutex
    global de execução (com_lock_execucao_global): se qualquer outra
    retirada já estiver rodando neste PC — mesmo grupo, outro grupo, ou
    manual — espera terminar antes de começar."""
    eh_mensal = rotulo == "FECHAMENTO MENSAL"

    def _rodar():
        registrar(f"======== [{grupo}] {rotulo} — competência {competencia} ========")
        inicio = datetime.now()
        resultados = {}
        for empresa in empresas:
            falhas = pipeline(empresa, competencia, gerar_retencoes_e_pdf=eh_mensal)
            resultados[empresa] = (falhas == 0)

        ok = [nome for nome, sucesso in resultados.items() if sucesso]
        com_falha = [nome for nome, sucesso in resultados.items() if not sucesso]

        registro = {
            "grupo": grupo,
            "rotulo": rotulo,
            "competencia": competencia,
            "data": date.today().isoformat(),
            "hora": datetime.now().strftime("%H:%M"),
            "empresas_ok": ok,
            "empresas_com_falha": com_falha,
        }

        def _salvar_no_historico():
            atual = json.loads(ARQ_ULTIMA_EXECUCAO.read_text(encoding="utf-8")) if ARQ_ULTIMA_EXECUCAO.exists() else {}
            execucoes = atual.setdefault("execucoes", [])
            execucoes.append(registro)
            if len(execucoes) > LIMITE_HISTORICO:
                del execucoes[:len(execucoes) - LIMITE_HISTORICO]
            ARQ_ULTIMA_EXECUCAO.write_text(json.dumps(atual, indent=2, ensure_ascii=False), encoding="utf-8")
        _com_mutex_estado(_salvar_no_historico)

        if com_falha:
            registrar(f"======== [{grupo}] {rotulo} TERMINADO — {len(ok)} OK, {len(com_falha)} COM FALHA ========")
            avisar("NFS-e Automático", f"[{grupo}] {rotulo.title()}: {len(ok)} empresa(s) OK, "
                                        f"{len(com_falha)} com falha ({', '.join(com_falha)}). "
                                        f"Abra o programa para ver detalhes.")
        else:
            registrar(f"======== [{grupo}] {rotulo} TERMINADO COM SUCESSO ========")
            avisar("NFS-e Automático", f"[{grupo}] {rotulo.title()}: {len(ok)} empresa(s) concluída(s) com sucesso.")

        if eh_mensal:
            avisar_retencoes_destacadas(grupo, empresas, competencia)
        else:
            avisar_notas_retroativas(grupo, empresas, competencia, inicio)

    com_lock_execucao_global(_rodar)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--grupo", required=True, help='nome do grupo (deve bater com "nome" em config.json -> grupos)')
    parser.add_argument("--modo", choices=["auto", "mensal", "semanal", "quinzenal"], default="auto",
                        help="qual verificação rodar (a tarefa do Agendador sempre chama com --modo fixo por grupo)")
    args = parser.parse_args()

    config = carregar_config()
    grupo_cfg = next((g for g in config.get("grupos", []) if g.get("nome") == args.grupo), None)
    if grupo_cfg is None:
        registrar(f"Grupo \"{args.grupo}\" não encontrado no config.json — nada a fazer.")
        return
    empresas = [e["nome"] for e in grupo_cfg.get("empresas", [])]
    agenda = grupo_cfg.get("agendamento", {})

    hoje = date.today()
    estado_completo = json.loads(ARQ_ESTADO.read_text(encoding="utf-8")) if ARQ_ESTADO.exists() else {}
    estado = estado_completo.get("grupos", {}).get(args.grupo, {})

    def salvar_estado():
        salvar_estado_grupo(args.grupo, estado)

    # ---- MENSAL: dia fixo do mês, fecha o mês anterior, com recuperação de atraso
    cfg_mensal = agenda.get("mensal", {})
    mes_fechamento = mes_anterior(hoje)
    devida = date(hoje.year, hoje.month, cfg_mensal.get("dia_mes", 1))
    eh_dia_de_fechamento = (cfg_mensal.get("ativo")
                            and hoje >= devida
                            and estado.get("ultimo_fechamento_mensal") != mes_fechamento)

    if args.modo in ("auto", "mensal") and eh_dia_de_fechamento:
        if hoje == devida:
            registrar(f"[{args.grupo}] Hoje é o dia do fechamento mensal de {mes_fechamento}.")
        else:
            registrar(f"[{args.grupo}] Fechamento mensal de {mes_fechamento} estava atrasado "
                      f"(devido em {devida:%d/%m}); executando agora.")
        executar_para_todas(args.grupo, empresas, mes_fechamento, "FECHAMENTO MENSAL")
        estado["ultimo_fechamento_mensal"] = mes_fechamento
        salvar_estado()
        return

    if args.modo == "auto" and eh_dia_de_fechamento:
        return  # já tratado acima; nunca deveria chegar aqui

    competencia_atual = f"{hoje.year:04d}-{hoje.month:02d}"

    # ---- SEMANAL: um dia da semana, toda semana — roda independente do
    # mensal (competências diferentes, nunca duplica trabalho)
    cfg_semanal = agenda.get("semanal", {})
    if (args.modo in ("auto", "semanal") and cfg_semanal.get("ativo")
            and hoje.weekday() == cfg_semanal.get("dia_semana")):
        registrar(f"[{args.grupo}] Hoje é dia de retirada semanal do mês atual.")
        executar_para_todas(args.grupo, empresas, competencia_atual, "RETIRADA SEMANAL")
        return

    # ---- QUINZENAL: um dia da semana, só a cada 14 dias — roda independente
    # do mensal (competências diferentes, nunca duplica trabalho)
    cfg_quinzenal = agenda.get("quinzenal", {})
    if args.modo in ("auto", "quinzenal") and cfg_quinzenal.get("ativo"):
        ultima = estado.get("ultima_execucao_quinzenal")
        ultima_data = date.fromisoformat(ultima) if ultima else None
        due_por_data = ultima_data is None or (hoje - ultima_data).days >= 14
        if hoje.weekday() == cfg_quinzenal.get("dia_semana") and due_por_data:
            registrar(f"[{args.grupo}] Hoje é dia de retirada quinzenal do mês atual.")
            executar_para_todas(args.grupo, empresas, competencia_atual, "RETIRADA QUINZENAL")
            estado["ultima_execucao_quinzenal"] = hoje.isoformat()
            salvar_estado()
            return

    registrar(f"[{args.grupo}] Nada agendado para hoje.")


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""
Retirada manual de UM mês específico, pra uma ou mais empresas (até 5) —
disparada pelo botão "Rodar agora..." da Tela 2 do assistente. Roda por
fora do ciclo do agendamento: não lê nem grava
rotina_estado.json/ultima_execucao.json, então não atrasa nem antecipa
nenhum fechamento automático. Reaproveita o mesmo pipeline() de rotina.py
(mesmo timeout de segurança e short-circuit), uma empresa de cada vez, com
UMA notificação final resumindo o total.

Uso:
  python assistente.py executar_agora --empresas "NOME1,NOME2" --competencia 2026-05
"""

import argparse
import sys

import rotina


def main() -> int:
    parser = argparse.ArgumentParser(description="Busca as notas de uma ou mais empresas num mês específico, fora do agendamento.")
    parser.add_argument("--empresas", required=True,
                         help='nomes separados por vírgula (devem bater com "nome" em config.json)')
    parser.add_argument("--competencia", required=True, help="mês alvo, ex.: 2026-05")
    args = parser.parse_args()

    nomes = [n.strip() for n in args.empresas.split(",") if n.strip()]

    rotina.registrar(f"======== RETIRADA MANUAL — {len(nomes)} empresa(s) — competência {args.competencia} ========")
    ok, com_falha = [], []
    for nome in nomes:
        falhas = rotina.pipeline(nome, args.competencia)
        (ok if falhas == 0 else com_falha).append(nome)

    if com_falha:
        rotina.registrar(f"======== RETIRADA MANUAL TERMINADA — {len(ok)} OK, {len(com_falha)} COM FALHA ========")
        rotina.notificar("NFS-e Automático",
                          f"Retirada manual ({args.competencia}): {len(ok)} empresa(s) OK, "
                          f"{len(com_falha)} com falha. Veja o rotina.log.")
    else:
        rotina.registrar(f"======== RETIRADA MANUAL TERMINADA COM SUCESSO ({len(ok)} empresa(s)) ========")
        rotina.notificar("NFS-e Automático",
                          f"Retirada manual ({args.competencia}): {len(ok)} empresa(s) concluída(s) com sucesso.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

# -*- coding: utf-8 -*-
"""
Retirada manual de UM mês específico, pra UMA empresa — disparada pelo botão
"Rodar agora..." da Tela 2 do assistente. Roda por fora do ciclo do
agendamento: não lê nem grava rotina_estado.json/ultima_execucao.json, então
não atrasa nem antecipa nenhum fechamento automático. Reaproveita o mesmo
pipeline() de rotina.py (mesmo timeout de segurança e short-circuit).

Uso:
  python assistente.py executar_agora --empresa "NOME" --competencia 2026-05
"""

import argparse
import sys

import rotina


def main() -> int:
    parser = argparse.ArgumentParser(description="Busca as notas de uma empresa/mês específico, fora do agendamento.")
    parser.add_argument("--empresa", required=True, help='nome exato da empresa (deve bater com "nome" em config.json)')
    parser.add_argument("--competencia", required=True, help="mês alvo, ex.: 2026-05")
    args = parser.parse_args()

    rotina.registrar(f"======== RETIRADA MANUAL — {args.empresa} — competência {args.competencia} ========")
    falhas = rotina.pipeline(args.empresa, args.competencia)

    if falhas:
        rotina.registrar("======== RETIRADA MANUAL TERMINADA — COM FALHA(S) ========")
        rotina.notificar("NFS-e Automático",
                          f"Retirada manual de \"{args.empresa}\" ({args.competencia}): "
                          "concluída com falha(s). Veja o rotina.log.")
    else:
        rotina.registrar("======== RETIRADA MANUAL TERMINADA COM SUCESSO ========")
        rotina.notificar("NFS-e Automático",
                          f"Retirada manual de \"{args.empresa}\" ({args.competencia}): concluída com sucesso.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

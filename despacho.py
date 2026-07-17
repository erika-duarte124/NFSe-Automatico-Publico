# -*- coding: utf-8 -*-
"""
Monta o comando para reinvocar este mesmo programa como um "sub-passo"
(baixar_nfse, gerar_relatorio, rotina, etc.) — funciona tanto rodando do
código-fonte (python assistente.py <subcomando> ...) quanto já empacotado
num .exe único (NFSe-Automatico.exe <subcomando> ...), sem precisar saber
qual dos dois é em nenhum outro lugar do código.

PASTA também é definida aqui (fonte única) porque, dentro de um .exe
gerado com PyInstaller --onefile, `Path(__file__)` aponta para uma pasta
TEMPORÁRIA de extração (ex.: AppData/Local/Temp/_MEIxxxxx) — não para
onde o .exe realmente está. Nesse caso, a pasta certa é a do próprio
executável (sys.executable).
"""

import sys
from pathlib import Path


def _pasta_base() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


PASTA = _pasta_base()


def comando_base(silencioso: bool = False) -> list[str]:
    """silencioso=True usa pythonw.exe (sem janela de console) quando roda
    do código-fonte — usado para a tarefa do Agendador do Windows."""
    if getattr(sys, "frozen", False):
        return [sys.executable]
    executavel = sys.executable
    if silencioso:
        executavel = executavel.replace("python.exe", "pythonw.exe")
    return [executavel, str(PASTA / "assistente.py")]


def empresas_de(config: dict) -> list[dict]:
    """Achata a lista de empresas de todos os grupos do config.json — os
    scripts que baixam/geram relatório não precisam saber de grupo, só de
    qual empresa processar."""
    return [e for grupo in config.get("grupos", []) for e in grupo.get("empresas", [])]

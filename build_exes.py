# -*- coding: utf-8 -*-
"""
Gera as duas edições do executável a partir do mesmo código-fonte:

  - NFSe-Automatico.exe      (público, com os limites de LIMITE_EMPRESAS
                               etc. de assistente.py)
  - NFSe-Automatico_All.exe  (interna, sem limite nenhum)

A diferença entre as duas é só a constante EDICAO_INTERNA no topo de
assistente.py. Este script troca essa constante pra True só durante o
build da versão interna, e desfaz a troca logo depois — o arquivo sempre
volta pro estado original (EDICAO_INTERNA = False) ao final, com sucesso
ou falha, então nunca fica commitado como True por engano.

Uso:
  python build_exes.py
"""

import subprocess
import sys
from pathlib import Path

PASTA = Path(__file__).resolve().parent
ARQ_ASSISTENTE = PASTA / "assistente.py"

MARCA_PUBLICA = "EDICAO_INTERNA = False"
MARCA_INTERNA = "EDICAO_INTERNA = True"


def rodar_pyinstaller(spec: str) -> None:
    print(f"\n{'=' * 60}\nCompilando {spec}...\n{'=' * 60}")
    r = subprocess.run([sys.executable, "-m", "PyInstaller", spec, "--noconfirm"], cwd=str(PASTA))
    if r.returncode != 0:
        raise RuntimeError(f"PyInstaller falhou compilando {spec} (código {r.returncode})")


def main() -> int:
    original = ARQ_ASSISTENTE.read_text(encoding="utf-8")
    if MARCA_PUBLICA not in original:
        print(f"ERRO: não encontrei {MARCA_PUBLICA!r} em assistente.py — "
              "abortando pra não arriscar gerar as duas edições iguais.")
        return 1

    try:
        # 1) edição pública — código-fonte como está (EDICAO_INTERNA = False)
        rodar_pyinstaller("NFSe-Automatico.spec")

        # 2) edição interna — troca a constante só pra esse build
        interno = original.replace(MARCA_PUBLICA, MARCA_INTERNA, 1)
        ARQ_ASSISTENTE.write_text(interno, encoding="utf-8")
        rodar_pyinstaller("NFSe-Automatico_All.spec")
    finally:
        # sempre desfaz a troca, mesmo se um build falhar no meio
        ARQ_ASSISTENTE.write_text(original, encoding="utf-8")
        restaurado = ARQ_ASSISTENTE.read_text(encoding="utf-8")
        if restaurado != original:
            print("ATENÇÃO: assistente.py pode não ter sido restaurado corretamente — confira antes de commitar!")
        else:
            print("\nassistente.py restaurado ao original (EDICAO_INTERNA = False).")

    print("\nProntos:")
    print(f"  {PASTA / 'dist' / 'NFSe-Automatico' / 'NFSe-Automatico.exe'}")
    print(f"  {PASTA / 'dist' / 'NFSe-Automatico_All' / 'NFSe-Automatico_All.exe'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

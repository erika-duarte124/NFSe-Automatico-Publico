# -*- mode: python ; coding: utf-8 -*-
#
# Edição interna (sem limites) — mesma fonte do NFSe-Automatico.spec, só o
# nome do executável e os metadados de versão mudam. Só gera a versão sem
# limites de fato quando assistente.py está com EDICAO_INTERNA = True no
# momento do build — o build_exes.py cuida disso sozinho (troca, compila,
# desfaz).


a = Analysis(
    ['assistente.py'],
    pathex=[],
    binaries=[],
    datas=[('nfse_logo.png', '.')],
    hiddenimports=['baixar_nfse', 'gerar_danfse', 'gerar_relatorio', 'gerar_relatorio_pdf', 'gerar_retencoes', 'rotina', 'rodar_fila', 'backfill', 'executar_agora'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

# onedir (não onefile) — ver comentário no NFSe-Automatico.spec.
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='NFSe-Automatico_All',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version='version_info_all.txt',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='NFSe-Automatico_All',
)

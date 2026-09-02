# -*- mode: python ; coding: utf-8 -*-


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

# onedir (não onefile): o .exe roda direto dos arquivos desta pasta, sem se
# autoextrair pra uma pasta temporária (_MEIxxxxx) a cada abertura — evita
# o padrão "self-extract + relançar" que heurísticas de antivírus associam
# a dropper de malware, e facilita rodar num compartilhamento de rede sem
# ganhar a marca de "veio da internet" a cada execução.
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='NFSe-Automatico',
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
    version='version_info.txt',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='NFSe-Automatico',
)

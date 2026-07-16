# -*- coding: utf-8 -*-
"""
Protege a senha do certificado digital usando o DPAPI do Windows
(win32crypt.CryptProtectData) — amarrado à conta Windows de quem
cadastrou. Só esse mesmo usuário, nesse mesmo PC, consegue descriptografar.

Um valor sem o prefixo "dpapi:" é tratado como texto puro (compatibilidade
com configs antigos ou com quem preferir não usar criptografia).
"""

import base64

PREFIXO = "dpapi:"
ENTIDADE = "NFSe-Automatico"


def proteger(texto: str) -> str:
    """Criptografa uma senha em texto puro. Retorna string pronta p/ salvar
    no config.json (prefixo + base64)."""
    if not texto:
        return texto
    import win32crypt
    protegido = win32crypt.CryptProtectData(texto.encode("utf-8"), ENTIDADE, None, None, None, 0)
    return PREFIXO + base64.b64encode(protegido).decode("ascii")


def revelar(valor: str) -> str:
    """Descriptografa um valor salvo por proteger(). Se não tiver o prefixo,
    devolve o valor como veio (texto puro)."""
    if not valor or not valor.startswith(PREFIXO):
        return valor
    import win32crypt
    bruto = base64.b64decode(valor[len(PREFIXO):])
    _, texto = win32crypt.CryptUnprotectData(bruto, None, None, None, 0)
    return texto.decode("utf-8")


def esta_protegida(valor: str) -> bool:
    return bool(valor) and valor.startswith(PREFIXO)

# -*- coding: utf-8 -*-
"""
Baixador automático de NFS-e do Portal Nacional (Sistema Nacional NFS-e).

Baixa XML e PDF (DANFSe) das notas em que cada empresa configurada figura
como emitente, tomador ou intermediário, usando as APIs oficiais do ADN
(Ambiente de Dados Nacional) com autenticação por certificado digital A1.

APIs utilizadas:
  - ADN Distribuição (XML):  GET https://adn.nfse.gov.br/contribuintes/DFe/{NSU}
  - ADN DANFSe (PDF):        GET https://adn.nfse.gov.br/danfse/{chaveAcesso}
  - Sefin Nacional (XML por chave, fallback):
                             GET https://sefin.nfse.gov.br/sefinnacional/nfse/{chaveAcesso}

Uso:
  python baixar_nfse.py                          # sincroniza todas as empresas do config.json
  python baixar_nfse.py --empresa "RMW"          # só empresas cujo nome contenha "RMW"
  python baixar_nfse.py --sem-pdf                # baixa só os XMLs (não gera DANFSe)
  python baixar_nfse.py --nsu-inicial 0          # recomeça a distribuição do NSU 0
  python baixar_nfse.py --chave <50 dígitos> --empresa "RMW"   # baixa 1 nota avulsa (PDF+XML)
"""

import argparse
import base64
import getpass
import gzip
import json
import re
import sys
import time
import zlib
from datetime import datetime
from pathlib import Path

import requests
from requests_pkcs12 import Pkcs12Adapter

import despacho
import seguranca

# ---------------------------------------------------------------- constantes

# Ambientes disponíveis. "producao_restrita" é o ambiente de testes
# (https://adn.producaorestrita.nfse.gov.br / sefin.producaorestrita.nfse.gov.br).
HOSTS = {
    "producao": {"adn": "adn.nfse.gov.br", "sefin": "sefin.nfse.gov.br"},
    "producao_restrita": {"adn": "adn.producaorestrita.nfse.gov.br",
                          "sefin": "sefin.producaorestrita.nfse.gov.br"},
}

URL_ADN_DISTRIBUICAO = "https://{adn}/contribuintes/DFe/{{nsu}}"
URL_ADN_DANFSE = "https://{adn}/danfse/{{chave}}"
URL_SEFIN_NFSE = "https://{sefin}/SefinNacional/nfse/{{chave}}"


def definir_ambiente(ambiente: str) -> None:
    """Resolve as URLs conforme o ambiente do config.json (padrão: producao)."""
    global URL_ADN_DISTRIBUICAO, URL_ADN_DANFSE, URL_SEFIN_NFSE
    hosts = HOSTS.get(ambiente)
    if hosts is None:
        raise ValueError(f"Ambiente inválido no config: {ambiente!r} "
                         f"(use 'producao' ou 'producao_restrita')")
    URL_ADN_DISTRIBUICAO = URL_ADN_DISTRIBUICAO.format(**hosts)
    URL_ADN_DANFSE = URL_ADN_DANFSE.format(**hosts)
    URL_SEFIN_NFSE = URL_SEFIN_NFSE.format(**hosts)

PASTA_SCRIPT = despacho.PASTA
ARQ_CONFIG = PASTA_SCRIPT / "config.json"
ARQ_ESTADO = PASTA_SCRIPT / "estado_nsu.json"
# Checkpoint de NSU por competência: {cnpj: {"2026-07": nsu}}. Permite que uma
# 2ª chamada com --competencia para o MESMO mês (ex.: fechamento mensal depois
# da retirada semanal) continue de onde a anterior parou, em vez de reescanear
# tudo do zero. Independente por mês, então pedir uma competência antiga nunca
# vista ainda escaneia do zero normalmente — nada se perde.
ARQ_ESTADO_COMP = PASTA_SCRIPT / "estado_nsu_competencia.json"

TIMEOUT = 60          # segundos por requisição
PAUSA_ENTRE_LOTES = 1.5   # pausa entre chamadas de distribuição (evita throttling)


# ---------------------------------------------------------------- utilidades

def log(msg: str) -> None:
    print(f"[{datetime.now():%H:%M:%S}] {msg}")


def carregar_json(caminho: Path, padrao):
    if caminho.exists():
        return json.loads(caminho.read_text(encoding="utf-8"))
    return padrao


def salvar_json(caminho: Path, dados) -> None:
    caminho.write_text(json.dumps(dados, indent=2, ensure_ascii=False), encoding="utf-8")


def get_ci(dicio: dict, chave: str, padrao=None):
    """Busca chave em dict ignorando maiúsculas/minúsculas (a API varia o padrão)."""
    if not isinstance(dicio, dict):
        return padrao
    alvo = chave.lower()
    for k, v in dicio.items():
        if k.lower() == alvo:
            return v
    return padrao


def descompactar_xml(conteudo_b64: str) -> bytes:
    """O XML vem em base64, normalmente comprimido com GZip. Tenta os formatos possíveis."""
    bruto = base64.b64decode(conteudo_b64)
    for descomp in (gzip.decompress, zlib.decompress, lambda b: zlib.decompress(b, -15)):
        try:
            return descomp(bruto)
        except Exception:
            pass
    return bruto  # já veio sem compressão


def somente_digitos(texto: str) -> str:
    return re.sub(r"\D", "", texto or "")


def normalizar_competencia(texto: str) -> str:
    """Aceita '2026-05', '05/2026' ou '2026/05' e devolve sempre 'AAAA-MM'."""
    texto = texto.strip()
    m = re.fullmatch(r"(\d{4})[-/](\d{1,2})", texto)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"
    m = re.fullmatch(r"(\d{1,2})[-/](\d{4})", texto)
    if m:
        return f"{m.group(2)}-{int(m.group(1)):02d}"
    raise ValueError(f"Competência inválida: {texto!r} (use AAAA-MM, ex.: 2026-05)")


def descrever_erro(resp: requests.Response) -> str:
    """Extrai código/descrição do JSON de erro padrão da API (ResponseErro)."""
    try:
        dados = resp.json()
        erros = get_ci(dados, "erros") or [get_ci(dados, "erro")]
        partes = [f"{get_ci(e, 'codigo')}: {get_ci(e, 'descricao')}"
                  for e in erros if e]
        if partes:
            return "; ".join(partes)
    except Exception:
        pass
    return resp.text[:300]


def info_da_chave(chave: str) -> dict:
    """
    Decompõe a chave de acesso da NFS-e (50 dígitos):
      [0:7] cód. município | [7] ambiente gerador | [8] tipo inscrição |
      [9:23] inscrição federal do emitente | [23:36] número da NFS-e (13) |
      [36:40] AAMM da emissão | [40:49] cód. numérico | [49] DV
    """
    return {
        "municipio": chave[0:7],
        "inscricao_emitente": chave[9:23],
        "numero": chave[23:36].lstrip("0") or "0",
        "competencia": f"20{chave[36:38]}-{chave[38:40]}",
    }


def extrair_emitente_do_xml(xml: bytes) -> str | None:
    """Extrai o CNPJ/CPF do emitente de dentro do XML da NFS-e (tag <emit>)."""
    try:
        texto = xml.decode("utf-8", errors="replace")
        m = re.search(r"<emit>.*?<(CNPJ|CPF)>(\d+)</\1>", texto, re.DOTALL)
        if m:
            return m.group(2)
    except Exception:
        pass
    return None


def competencia_do_xml(xml: bytes) -> str | None:
    """
    Lê a competência REAL (dCompet = AAAA-MM) de dentro do XML da NFS-e.
    É o mês de prestação do serviço — pode diferir do mês de emissão (que é o
    que está embutido na chave de acesso). Para contabilidade, vale o dCompet.
    """
    try:
        m = re.search(r"<dCompet>(\d{4})-(\d{2})", xml.decode("utf-8", errors="replace"))
        if m:
            return f"{m.group(1)}-{m.group(2)}"
    except Exception:
        pass
    return None


def emissao_do_xml(xml: bytes) -> str | None:
    """Mês de EMISSÃO (dhEmi = AAAA-MM) de dentro do XML da NFS-e."""
    try:
        m = re.search(r"<dhEmi>(\d{4})-(\d{2})", xml.decode("utf-8", errors="replace"))
        if m:
            return f"{m.group(1)}-{m.group(2)}"
    except Exception:
        pass
    return None


def obter_com_retentativas(sessao: requests.Session, url: str, params=None,
                           tentativas: int = 9):
    """
    GET resistente a instabilidade: o servidor do ADN oscila (timeouts, 429,
    502, 503), então espera com pausas crescentes antes de desistir.
    Retorna a resposta, ou None se todas as tentativas falharem.
    """
    espera = 5.0
    for tentativa in range(1, tentativas + 1):
        try:
            resp = sessao.get(url, params=params, timeout=TIMEOUT)
        except requests.RequestException as exc:
            log(f"    Conexão falhou ({exc}); nova tentativa em {espera:.0f}s "
                f"({tentativa}/{tentativas})")
        else:
            if resp.status_code not in (429, 500, 502, 503, 504):
                return resp
            log(f"    HTTP {resp.status_code}; nova tentativa em {espera:.0f}s "
                f"({tentativa}/{tentativas})")
        time.sleep(espera)
        espera = min(espera * 2, 180)
    return None


# ---------------------------------------------------------------- sessão TLS

def criar_sessao(caminho_pfx: str, senha: str) -> requests.Session:
    """Sessão HTTPS com o certificado digital A1 (.pfx/.p12) acoplado (mTLS)."""
    if not Path(caminho_pfx).exists():
        raise FileNotFoundError(f"Certificado não encontrado: {caminho_pfx}")
    sessao = requests.Session()
    adaptador = Pkcs12Adapter(pkcs12_filename=caminho_pfx, pkcs12_password=senha)
    sessao.mount("https://", adaptador)
    sessao.headers.update({"Accept": "application/json"})
    return sessao


# ---------------------------------------------------------------- downloads

def salvar_xml(pasta_empresa: Path, chave: str, xml: bytes, cnpj_empresa: str,
               tipo_doc: str, retroativa: bool = False,
               mes_alvo: str | None = None) -> Path:
    info = info_da_chave(chave)
    emitente = extrair_emitente_do_xml(xml) or info["inscricao_emitente"]

    if tipo_doc.upper().startswith("EVENTO"):
        subpasta = "Eventos"
        # eventos não têm dCompet; usam o mês da chave (referência)
        competencia_pasta = info["competencia"]
        destino = pasta_empresa / competencia_pasta / subpasta
    else:
        if somente_digitos(emitente).lstrip("0") == somente_digitos(cnpj_empresa).lstrip("0"):
            subpasta = "Emitidas"
        else:
            subpasta = "Recebidas"
        if retroativa and mes_alvo:
            # nota EMITIDA no mês alvo, mas com competência de outro mês:
            # guarda numa subpasta "XML retroativo" dentro do mês alvo
            destino = pasta_empresa / mes_alvo / subpasta / "XML retroativo"
        else:
            # nota vai para a pasta da COMPETÊNCIA real (dCompet), não da emissão
            competencia_pasta = competencia_do_xml(xml) or info["competencia"]
            destino = pasta_empresa / competencia_pasta / subpasta

    destino.mkdir(parents=True, exist_ok=True)

    if tipo_doc.upper().startswith("EVENTO"):
        # eventos podem repetir a mesma chave; inclui carimbo para não sobrescrever
        nome = f"{chave}-evento-{datetime.now():%Y%m%d%H%M%S%f}.xml"
    else:
        nome = f"{chave}.xml"

    arquivo = destino / nome
    arquivo.write_bytes(xml)
    return arquivo


def baixar_danfse(sessao: requests.Session, chave: str, destino: Path,
                  tentativas: int = 4) -> bool:
    """
    Baixa o PDF (DANFSe) pelo ADN. O servidor limita a frequência de chamadas
    (HTTP 429) e às vezes oscila (502/503), então tenta de novo com pausas
    crescentes antes de desistir.
    """
    url = URL_ADN_DANFSE.format(chave=chave)
    espera = 3.0
    for tentativa in range(1, tentativas + 1):
        try:
            resp = sessao.get(url, timeout=TIMEOUT, headers={"Accept": "application/pdf"})
        except requests.RequestException as exc:
            log(f"    DANFSe: erro de conexão ({exc}); tentativa {tentativa}/{tentativas}")
            time.sleep(espera)
            espera *= 2
            continue

        if resp.status_code == 200 and resp.content[:4] == b"%PDF":
            destino.parent.mkdir(parents=True, exist_ok=True)
            destino.write_bytes(resp.content)
            return True

        if resp.status_code == 200:
            # 200 mas não é PDF: pode vir JSON com o PDF em base64
            try:
                dados = resp.json()
                b64 = get_ci(dados, "pdf") or get_ci(dados, "danfse") or get_ci(dados, "arquivo")
                if b64:
                    destino.parent.mkdir(parents=True, exist_ok=True)
                    destino.write_bytes(base64.b64decode(b64))
                    return True
            except Exception:
                pass
            log(f"    DANFSe: resposta 200 em formato inesperado para {chave}")
            return False

        if resp.status_code in (429, 500, 502, 503, 504):
            retry_after = resp.headers.get("Retry-After", "")
            pausa = float(retry_after) if retry_after.isdigit() else espera
            log(f"    DANFSe: HTTP {resp.status_code}, aguardando {pausa:.0f}s "
                f"(tentativa {tentativa}/{tentativas})")
            time.sleep(pausa)
            espera *= 2
            continue

        log(f"    DANFSe: HTTP {resp.status_code} em {url} ({descrever_erro(resp)})")
        return False

    log(f"    DANFSe: desisti após {tentativas} tentativas: {chave}")
    return False


def baixar_xml_por_chave(sessao: requests.Session, chave: str) -> bytes | None:
    """Consulta avulsa do XML pela chave (API Sefin Nacional)."""
    url = URL_SEFIN_NFSE.format(chave=chave)
    try:
        resp = sessao.get(url, timeout=TIMEOUT)
    except requests.RequestException as exc:
        log(f"    XML por chave: erro de conexão: {exc}")
        return None
    if resp.status_code != 200:
        log(f"    XML por chave: HTTP {resp.status_code} em {url} "
            f"({descrever_erro(resp)})")
        return None
    # resposta usual: JSON com campo nfseXmlGZipB64
    try:
        dados = resp.json()
        b64 = (get_ci(dados, "nfseXmlGZipB64") or get_ci(dados, "arquivoXml")
               or get_ci(dados, "xml"))
        if b64:
            return descompactar_xml(b64)
    except ValueError:
        if resp.content.lstrip().startswith(b"<"):
            return resp.content
    return None


# ------------------------------------------------------- distribuição (NSU)

def processar_lote(resposta_json: dict) -> list[dict]:
    """Normaliza a lista de documentos de um lote da distribuição."""
    lote = (get_ci(resposta_json, "LoteDFe") or get_ci(resposta_json, "Lote")
            or get_ci(resposta_json, "documentos") or [])
    docs = []
    for item in lote:
        docs.append({
            "nsu": int(get_ci(item, "NSU") or 0),
            "chave": somente_digitos(str(get_ci(item, "ChaveAcesso") or "")),
            "tipo": str(get_ci(item, "TipoDocumento") or "NFSE"),
            "xml_b64": get_ci(item, "ArquivoXml") or get_ci(item, "XmlGZipB64")
                       or get_ci(item, "arquivo"),
        })
    return docs


def sincronizar_empresa(empresa: dict, pasta_saida: Path, estado: dict,
                        estado_comp: dict,
                        baixar_pdf: bool, nsu_inicial: int | None,
                        competencia: str | None = None) -> None:
    nome = empresa["nome"]
    cnpj = somente_digitos(empresa["cnpj"])
    log(f"==== {nome} (CNPJ {cnpj}) ====")
    if competencia:
        log(f"  Filtro de competência ativo: somente notas de {competencia}.")

    senha = seguranca.revelar(empresa.get("senha")) or getpass.getpass(f"Senha do certificado de {nome}: ")
    sessao = criar_sessao(empresa["certificado"], senha)

    pasta_empresa = pasta_saida / nome
    if nsu_inicial is not None:
        nsu = nsu_inicial
    elif competencia:
        # retoma de onde a última varredura DESSA competência parou (ex.: a
        # retirada semanal já avançou boa parte; o fechamento mensal só busca
        # o que apareceu de novo desde então). Independente por mês: pedir
        # uma competência nunca vista ainda começa do zero normalmente.
        nsu = int(estado_comp.get(cnpj, {}).get(competencia, 0))
        if nsu:
            log(f"  Retomando de onde a última varredura de {competencia} parou (NSU {nsu}).")
    else:
        nsu = int(estado.get(cnpj, 0))
    total_xml = total_pdf = ignorados = 0
    confirmacoes_fim = 0

    while True:
        url = URL_ADN_DISTRIBUICAO.format(nsu=nsu)
        params = {"cnpjConsulta": cnpj} if empresa.get("usar_cnpj_consulta") else None
        resp = obter_com_retentativas(sessao, url, params)
        if resp is None:
            log(f"  Distribuição indisponível mesmo após várias tentativas "
                f"(parou no NSU {nsu}). Rode de novo mais tarde para continuar.")
            break

        if resp.status_code == 404:
            # Em momentos de instabilidade o servidor responde 404 mesmo havendo
            # documentos; confirma o fim da fila 3 vezes antes de acreditar.
            confirmacoes_fim += 1
            if confirmacoes_fim < 3:
                log(f"  NSU {nsu}: sem documentos (confirmando fim da fila, "
                    f"{confirmacoes_fim}/3)...")
                time.sleep(10)
                continue
            log(f"  Sem documentos novos a partir do NSU {nsu}.")
            break
        confirmacoes_fim = 0
        if resp.status_code != 200:
            log(f"  HTTP {resp.status_code} na distribuição (NSU {nsu}): "
                f"{descrever_erro(resp)}")
            break

        try:
            dados = resp.json()
        except ValueError:
            log(f"  Resposta inesperada (não é JSON). Salvando para análise.")
            (PASTA_SCRIPT / f"resposta_bruta_{cnpj}_{nsu}.txt").write_bytes(resp.content)
            break

        docs = processar_lote(dados)
        if not docs:
            log(f"  Fim da fila de documentos (NSU {nsu}).")
            break

        for doc in docs:
            if not doc["xml_b64"]:
                continue
            try:
                xml = descompactar_xml(doc["xml_b64"])
            except Exception as exc:
                log(f"    NSU {doc['nsu']}: falha ao decodificar XML: {exc}")
                continue

            # filtro por competência REAL (dCompet) OU emissão (dhEmi). Eventos
            # passam sempre, para que a situação (cancelada/substituída) seja
            # sempre atualizada.
            eh_evento = doc["tipo"].upper().startswith("EVENTO")
            retroativa = False
            if competencia and not eh_evento:
                comp_nota = competencia_do_xml(xml) or info_da_chave(doc["chave"])["competencia"]
                if comp_nota != competencia:
                    # não é da competência do mês. Mas se foi EMITIDA neste mês,
                    # é uma nota retroativa -> guarda em "XML retroativo".
                    emi_nota = emissao_do_xml(xml) or info_da_chave(doc["chave"])["competencia"]
                    if emi_nota == competencia:
                        retroativa = True
                    else:
                        ignorados += 1
                        continue

            arquivo = salvar_xml(pasta_empresa, doc["chave"], xml, cnpj, doc["tipo"],
                                 retroativa=retroativa, mes_alvo=competencia)
            total_xml += 1
            marca = " [RETROATIVA]" if retroativa else ""
            log(f"    NSU {doc['nsu']} [{doc['tipo']}]{marca} -> {arquivo.relative_to(pasta_saida)}")

            if baixar_pdf and not doc["tipo"].upper().startswith("EVENTO"):
                destino_pdf = arquivo.with_suffix(".pdf")
                if not destino_pdf.exists():
                    # durante o download em massa, tenta o PDF UMA vez só;
                    # os que falharem ficam para a passada --completar-pdf
                    if baixar_danfse(sessao, doc["chave"], destino_pdf, tentativas=1):
                        total_pdf += 1
                    time.sleep(1.0)

        maior_nsu = max(d["nsu"] for d in docs)
        if maior_nsu <= nsu:
            break
        nsu = maior_nsu
        if competencia:
            estado_comp.setdefault(cnpj, {})[competencia] = nsu
            salvar_json(ARQ_ESTADO_COMP, estado_comp)
        else:
            estado[cnpj] = nsu
            salvar_json(ARQ_ESTADO, estado)
        time.sleep(PAUSA_ENTRE_LOTES)

    if competencia:
        estado_comp.setdefault(cnpj, {})[competencia] = nsu
        salvar_json(ARQ_ESTADO_COMP, estado_comp)
    else:
        estado[cnpj] = nsu
        salvar_json(ARQ_ESTADO, estado)
    resumo = f"  Concluído: {total_xml} XML(s) e {total_pdf} PDF(s) novos."
    if competencia:
        resumo += f" {ignorados} nota(s) de outros meses ignorada(s)."
    log(resumo + f" Último NSU: {nsu}")


def completar_pdfs(empresa: dict, pasta_saida: Path,
                   competencia: str | None = None) -> None:
    """Baixa o DANFSe dos XMLs já salvos que ainda estão sem o PDF correspondente."""
    nome = empresa["nome"]
    pasta_empresa = pasta_saida / nome
    if competencia:
        pasta_empresa = pasta_empresa / competencia
    if not pasta_empresa.exists():
        log(f"==== {nome}: nenhuma nota baixada ainda (rode a sincronização antes).")
        return

    pendentes = [xml for xml in sorted(pasta_empresa.rglob("*.xml"))
                 if "Eventos" not in xml.parts and not xml.with_suffix(".pdf").exists()]
    log(f"==== {nome}: {len(pendentes)} PDF(s) faltando ====")
    if not pendentes:
        return

    senha = seguranca.revelar(empresa.get("senha")) or getpass.getpass(f"Senha do certificado de {nome}: ")
    sessao = criar_sessao(empresa["certificado"], senha)

    baixados = 0
    for xml in pendentes:
        chave = xml.stem
        if baixar_danfse(sessao, chave, xml.with_suffix(".pdf")):
            baixados += 1
            log(f"    PDF ok -> {xml.with_suffix('.pdf').relative_to(pasta_saida)}")
        time.sleep(1.0)
    log(f"  Concluído: {baixados} de {len(pendentes)} PDF(s) baixados.")


def baixar_nota_avulsa(empresa: dict, pasta_saida: Path, chave: str,
                       baixar_pdf: bool) -> None:
    chave = somente_digitos(chave)
    if len(chave) != 50:
        log(f"Chave inválida (precisa ter 50 dígitos): {chave}")
        return
    nome = empresa["nome"]
    senha = seguranca.revelar(empresa.get("senha")) or getpass.getpass(f"Senha do certificado de {nome}: ")
    sessao = criar_sessao(empresa["certificado"], senha)
    cnpj = somente_digitos(empresa["cnpj"])
    pasta_empresa = pasta_saida / nome

    xml = baixar_xml_por_chave(sessao, chave)
    if xml:
        arquivo = salvar_xml(pasta_empresa, chave, xml, cnpj, "NFSE")
        log(f"  XML salvo em {arquivo}")
        destino_pdf = arquivo.with_suffix(".pdf")
    else:
        log("  XML não disponível por chave (a nota pode não ter sido emitida pelo "
            "Emissor Nacional). Tentando apenas o PDF.")
        info = info_da_chave(chave)
        destino_pdf = pasta_empresa / info["competencia"] / f"{chave}.pdf"

    if baixar_pdf:
        if baixar_danfse(sessao, chave, destino_pdf):
            log(f"  PDF salvo em {destino_pdf}")
        else:
            log("  Não foi possível baixar o DANFSe.")


# ----------------------------------------------------------------- principal

def main() -> int:
    parser = argparse.ArgumentParser(description="Baixa XML e PDF de NFS-e do Portal Nacional.")
    parser.add_argument("--empresa", help="filtra empresas pelo nome (contém, sem diferenciar maiúsculas)")
    parser.add_argument("--sem-pdf", action="store_true", help="não baixa o DANFSe (PDF)")
    parser.add_argument("--nsu-inicial", type=int, default=None,
                        help="ignora o estado salvo e começa deste NSU (ex.: 0 para tudo)")
    parser.add_argument("--chave", help="baixa uma única nota pela chave de acesso (50 dígitos)")
    parser.add_argument("--completar-pdf", action="store_true",
                        help="baixa apenas os PDFs que faltam para XMLs já salvos")
    parser.add_argument("--competencia",
                        help="baixa somente as notas do mês informado (ex.: 2026-05 ou 05/2026)")
    args = parser.parse_args()

    competencia = normalizar_competencia(args.competencia) if args.competencia else None

    if not ARQ_CONFIG.exists():
        print(f"Crie o arquivo de configuração: {ARQ_CONFIG}\n"
              f"Use o config.exemplo.json como modelo.")
        return 1

    config = carregar_json(ARQ_CONFIG, {})
    definir_ambiente(config.get("ambiente", "producao"))
    pasta_saida = Path(config.get("pasta_saida", PASTA_SCRIPT / "notas"))
    empresas = despacho.empresas_de(config)
    if args.empresa:
        empresas = [e for e in empresas if args.empresa.lower() in e["nome"].lower()]
    if not empresas:
        print("Nenhuma empresa encontrada na configuração com esse filtro.")
        return 1

    estado = carregar_json(ARQ_ESTADO, {})
    estado_comp = carregar_json(ARQ_ESTADO_COMP, {})

    for empresa in empresas:
        try:
            if args.chave:
                baixar_nota_avulsa(empresa, pasta_saida, args.chave, not args.sem_pdf)
            elif args.completar_pdf:
                completar_pdfs(empresa, pasta_saida, competencia)
            else:
                sincronizar_empresa(empresa, pasta_saida, estado, estado_comp,
                                    not args.sem_pdf, args.nsu_inicial,
                                    competencia)
        except Exception as exc:
            log(f"  ERRO em {empresa['nome']}: {exc}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

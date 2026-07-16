# Documentação técnica

Para quem quer entender por dentro, contribuir, ou rodar direto do código-fonte (em vez do `.exe`).

## Como funciona (as 3 APIs do Sistema Nacional NFS-e)

| API | URL | O que faz | O que retorna |
|---|---|---|---|
| **ADN Distribuição** | `https://adn.nfse.gov.br/contribuintes/DFe/{NSU}` | Entrega em lote (até 50 por vez) todos os documentos da empresa, em sequência numerada (NSU) | XML (gzip + base64) |
| **ADN DANFSe** | `https://adn.nfse.gov.br/danfse/{chaveAcesso}` | Gera o PDF da nota a partir da chave de acesso (50 dígitos) | PDF |
| **Sefin Nacional** | `https://sefin.nfse.gov.br/sefinnacional/nfse/{chaveAcesso}` | Consulta avulsa do XML pela chave (notas do Emissor Nacional) | XML (gzip + base64) |

O **NSU** funciona como um "número de página": o programa guarda o último NSU
baixado de cada CNPJ (`estado_nsu.json` / `estado_nsu_competencia.json`) e,
na próxima execução, continua de onde parou — só baixa o que é novo.

## Autenticação — ponto mais importante

Todas as APIs exigem **certificado digital ICP-Brasil** conectado na
requisição (TLS com autenticação mútua). Não existe login/senha nem token.

- O certificado precisa ter o **mesmo CNPJ raiz** da empresa consultada
  (matriz pode consultar filiais, conforme o manual do ADN).
- **Este programa trabalha com A1 (arquivo `.pfx`/`.p12`)**. A3 (cartão/token
  físico) é aceito pela API, mas exigiria integração com o dispositivo —
  fora do escopo deste projeto.
- A senha do certificado é criptografada com o **DPAPI do Windows**
  (`seguranca.py`) antes de ir para o `config.json` — só a mesma conta
  Windows que cadastrou consegue descriptografar.

## Rodando do código-fonte (sem o `.exe`)

```
python -m pip install -r requirements.txt
python assistente.py
```

Isso abre a tela de configuração normalmente. Os outros scripts também podem
ser chamados direto, para quem prefere linha de comando:

```
python assistente.py baixar_nfse --empresa "NOME" --competencia 2026-05
python assistente.py gerar_relatorio --empresa "NOME" --competencia 2026-05
python assistente.py gerar_retencoes --empresa "NOME" --competencia 2026-05
python assistente.py gerar_relatorio_pdf --empresa "NOME" --competencia 2026-05
python assistente.py rotina --modo mensal      # ou semanal / quinzenal / auto
python assistente.py rodar_fila 2026-05
```

`assistente.py` funciona como um **despachante** (`despacho.py`): o primeiro
argumento escolhe qual script "interno" roda, dentro do mesmo processo — sem
isso não seria possível empacotar tudo num `.exe` único (dentro de um `.exe`
não existem arquivos `.py` soltos para chamar como processo separado).

## Empacotando o `.exe`

```
pip install pyinstaller
pyinstaller --onefile --windowed --name NFSe-Automatico ^
  --hidden-import baixar_nfse --hidden-import gerar_relatorio ^
  --hidden-import gerar_relatorio_pdf --hidden-import gerar_retencoes ^
  --hidden-import rotina --hidden-import rodar_fila ^
  assistente.py
```

Os `--hidden-import` são necessários porque o despachante importa os módulos
dinamicamente (`__import__(nome)`), e o PyInstaller não detecta isso sozinho
na análise estática.

**Atenção**: dentro de um `.exe` gerado com `--onefile`, `Path(__file__)`
aponta para uma pasta temporária de extração (`AppData/Local/Temp/_MEIxxxxx`),
não para onde o `.exe` está de verdade. Por isso `despacho.py` calcula a
pasta base a partir de `sys.executable` quando `sys.frozen` é `True`. Se você
adicionar um novo arquivo ao projeto que precise saber "onde estou", importe
`PASTA` de `despacho.py` em vez de calcular `Path(__file__).resolve().parent`
de novo.

## Como as notas são organizadas em disco

```
pasta_saida/
  Certificados/
    <cnpj>.pfx                (cópia do certificado, feita pela tela de cadastro)
  NOME DA EMPRESA/
    2026-02/
      Emitidas/   <chave>.xml + .pdf
        XML retroativo/       (nota emitida em 02 mas com competência anterior)
      Recebidas/  ...
      Eventos/    (cancelamentos, substituições etc.)
    2026-02_Relatorios/
      Relatorio_Simples_...xlsx
      Relatorio_Retencoes_...xlsx
      Relatorio_...pdf
```

A separação Emitidas/Recebidas é automática: o programa compara o CNPJ do
emitente (de dentro do XML, ou das posições 10–23 da chave) com o CNPJ da
empresa cadastrada.

## `config.json` — estrutura completa

Ver `config.exemplo.json`. Campos principais:

- `pasta_saida`: onde tudo é salvo.
- `avisar_cert_vencido`: se `false`, a tela não avisa mais sobre certificados vencidos.
- `periodo_inicial`: `{"tipo": "completo"}` ou `{"tipo": "mes_especifico", "desde": "2026-01"}`.
- `agendamento`: um bloco por frequência (`mensal`, `semanal`, `quinzenal`), cada um com
  `ativo`, `hora`, e `dia_mes` (mensal) ou `dia_semana` (semanal/quinzenal — 0=segunda...6=domingo).
- `empresas`: lista de `{nome, cnpj, certificado, senha, cert_validade}`.

## Limitações que valem saber

- **Só vê o que está no ADN.** Notas de municípios **não conveniados** ao
  padrão nacional (sistema próprio, sem compartilhar) não aparecem.
- A consulta avulsa de XML por chave (Sefin) só retorna notas **emitidas pelo
  Emissor Nacional**; para as demais, o XML vem pela distribuição (NSU).
- O servidor do DANFSe limita a frequência de downloads (HTTP 429). O
  programa espera e tenta de novo sozinho; se ainda assim algum PDF ficar
  faltando, rode de novo — ele retoma automaticamente.
- O ADN oscila bastante (429/502/503) — isso é do servidor do governo, não
  do programa. Reexecutar resolve na maioria das vezes.

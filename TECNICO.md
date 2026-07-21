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
python assistente.py rotina --grupo "Grupo A" --modo mensal   # ou semanal / quinzenal / auto
python assistente.py rodar_fila 2026-05
python assistente.py backfill                  # busca do histórico inicial (ver abaixo)
python assistente.py executar_agora --empresa "NOME" --competencia 2026-05   # retirada manual (ver abaixo)
```

`assistente.py` funciona como um **despachante** (`despacho.py`): o primeiro
argumento escolhe qual script "interno" roda, dentro do mesmo processo — sem
isso não seria possível empacotar tudo num `.exe` único (dentro de um `.exe`
não existem arquivos `.py` soltos para chamar como processo separado).

## Empacotando o `.exe`

```
pip install pyinstaller
pyinstaller --onefile --windowed --noupx --name NFSe-Automatico ^
  --hidden-import baixar_nfse --hidden-import gerar_relatorio ^
  --hidden-import gerar_relatorio_pdf --hidden-import gerar_retencoes ^
  --hidden-import rotina --hidden-import rodar_fila --hidden-import backfill ^
  --hidden-import executar_agora ^
  assistente.py
```

Os `--hidden-import` são necessários porque o despachante importa os módulos
dinamicamente (`__import__(nome)`), e o PyInstaller não detecta isso sozinho
na análise estática.

`--noupx` desativa a compressão UPX do executável final — deixa o arquivo um
pouco maior, mas reduz bastante o risco de antivírus confundir o `.exe` com
malware (compressão UPX é uma técnica também usada por muitos packers de
vírus, então alguns antivírus flagram por semelhança, não por conteúdo real).

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

- `pasta_saida`: onde tudo é salvo (compartilhada por todos os grupos).
- `avisar_cert_vencido`: se `false`, a tela não avisa mais sobre certificados vencidos.
- `periodo_inicial`: `{"tipo": "completo"}` ou `{"tipo": "mes_especifico", "desde": "2026-01"}` — global,
  vale para o backfill inicial de todas as empresas de todos os grupos.
- `grupos`: lista de até 3 grupos (`LIMITE_GRUPOS` em `assistente.py`), cada um com:
  - `nome`: identifica o grupo — usado no nome da tarefa do Agendador e em `rotina.py --grupo`.
  - `agendamento`: um bloco por frequência (`mensal`, `semanal`, `quinzenal`), cada um com `ativo`,
    `hora`, e `dia_mes` (mensal) ou `dia_semana` (semanal/quinzenal — 0=segunda...6=domingo). Agenda
    independente por grupo — grupos diferentes podem rodar em dias/horários diferentes.
  - `empresas`: lista de até 10 (`LIMITE_POR_GRUPO`) `{nome, cnpj, certificado, senha, cert_validade}`.
    Total de empresas somando todos os grupos limitado a 20 (`LIMITE_EMPRESAS`).

`rotina_estado.json` e `ultima_execucao.json` também são organizados por grupo internamente
(`{"grupos": {"Nome do Grupo": {...}}}`), já que cada grupo tem seu próprio ciclo de fechamento.

## Segurança de execução (`rotina.py`)

- `rotina.py` roda **uma vez por grupo** (`--grupo "Nome do Grupo"`) — o
  Agendador do Windows tem uma tarefa por grupo, cada uma chamando com seu
  próprio `--grupo`. Estado (`rotina_estado.json`), resultado
  (`ultima_execucao.json`) e log ficam isolados por grupo, então um grupo
  travando ou atrasando não afeta o ciclo dos outros.
- **Timeout por passo** (`TIMEOUT_PASSO_SEGUNDOS`, 30 min): se um passo travar
  (rede lenta, senha em branco esperando input que nunca vem etc.), é
  cancelado e registrado como falha — não trava a fila inteira.
- **Short-circuit**: se "baixar XML+PDF" falhar, os 4 passos seguintes
  daquela empresa são pulados (não faz sentido gerar relatório sem dados).
  Isso limita o pior caso por empresa a 1 timeout, não 5.
- **Notificação do Windows** (`win11toast`) ao final de cada execução, com
  o resumo (quantas empresas OK, quantas com falha).
- **`ultima_execucao.json`**: resultado estruturado da última execução —
  `{rotulo, competencia, data, empresas_ok, empresas_com_falha}`. É o que
  alimenta a tela "Histórico de execuções..." do assistente (ver abaixo).

## Retirada manual de um mês específico (`executar_agora.py`)

Botão "Rodar agora (mês específico)..." na Tela 2 do assistente — abre um
diálogo (empresa + mês/ano) e dispara em segundo plano
(`subprocess.Popen`, não trava a tela). Reaproveita o mesmo `pipeline()` de
`rotina.py` (mesmo timeout de segurança e short-circuit), mas **não lê nem
grava `rotina_estado.json`/`ultima_execucao.json`** — roda inteiramente por
fora do ciclo do agendamento, então não atrasa nem antecipa nenhum
fechamento automático. Registra no mesmo `rotina.log`, com o cabeçalho
"RETIRADA MANUAL" pra diferenciar das execuções agendadas. Notifica o
Windows ao final, igual `rotina.py`.

Útil pra pegar uma nota emitida com atraso num mês já fechado — o
NSU-checkpoint por competência (`estado_nsu_competencia.json`) garante que
rodar de novo um mês já processado não duplica nada: só baixa o que for
realmente novo.

## Histórico de execuções (tela "Histórico de execuções...")

Botão na Tela 2, ao lado do "Rodar agora...". Lê `ultima_execucao.json` e
mostra numa tabela: grupo, tipo de execução, competência, data, empresa e
status (OK/FALHOU, com a falha destacada em vermelho claro). Tem botão
"Exportar para Excel..." (`openpyxl`) que gera a mesma tabela num `.xlsx`,
com a mesma marcação visual das falhas.

## Busca do histórico inicial (`backfill.py`)

Disparada uma vez, em processo separado (`subprocess.Popen`, não bloqueia a
tela), logo que o assistente conclui o cadastro — baseada em
`periodo_inicial`:
- `"mes_especifico"`: baixa + gera relatório de cada mês, de `desde` até o
  mês atual, pra cada empresa.
- `"completo"`: baixa o histórico inteiro (sem filtro de competência),
  depois varre as pastas de competência que foram criadas e gera relatório
  de cada uma.

Log em `backfill.log`. Usa o mesmo timeout de segurança do `rotina.py`.

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

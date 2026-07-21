# NFS-e Automático

Baixa automaticamente as **notas fiscais de serviço (NFS-e)** da sua empresa
direto do **Portal Nacional** (Sistema Nacional NFS-e), gera relatórios em
Excel e PDF, e roda sozinho, no dia e horário que você escolher e sem
precisar saber programar.

## O que o programa faz

- Baixa **XML e PDF (DANFSe)** de todas as notas emitidas e recebidas.
- Gera um **Relatório Simples** (Excel), um **Relatório de Retenções**
  (Excel — ISS, INSS, IRRF, CSLL, PIS/COFINS) e um **Relatório em PDF**.
- Roda automaticamente: mensal (fecha o mês anterior), semanal e/ou
  quinzenal — você escolhe.
- Opção de rodar data específica que você selecionar.
- Guarda a senha do certificado **criptografada**, presa à sua conta do
  Windows.
- Avisa quando um certificado está prestes a vencer.

## Como instalar (nenhuma programação necessária)

1. Baixe o `NFSe-Automatico.exe` na aba **Releases** deste repositório.
2. Salve numa pasta fixa do seu computador (ex.: `Documentos\NFSe-Automatico`)
   — não rode direto da pasta de Downloads, para não se perder se organizar
   os arquivos depois.
   — Lembre-se de não modificar a pasta ou excluí-la.
4. Dê dois cliques no arquivo. Pronto, o assistente abre.

O Windows pode mostrar um aviso do SmartScreen ("Windows protegeu seu PC"),
porque o programa é novo e ainda não tem muitos downloads — é normal para
programas independentes. Clique em **Mais informações → Executar assim
mesmo**.

## Antes de começar: você vai precisar de

- Um **certificado digital A1** (arquivo `.pfx` ou `.p12`) da empresa —
  emitido por uma Autoridade Certificadora (Serasa, Certisign, Soluti, etc.).
  O certificado **A3** (cartão/token físico) não funciona aqui.
- A **senha** desse certificado.

## Passo a passo do assistente

**1. Onde guardar as notas** — escolha uma pasta (o programa sugere uma
dentro de Downloads). Tudo fica organizado ali: empresa → mês → notas e
relatórios.

**2. Cadastro de empresas** — para cada empresa: nome, CNPJ, o arquivo do
certificado (botão "Escolher...") e a senha. Clique em **Validar
certificado** antes de adicionar — o programa confere localmente (sem
precisar de internet) se a senha está certa e mostra até quando o
certificado vale. Pode cadastrar até 20 empresas, e editar ou
remover qualquer uma antes de continuar.

**3. Período inicial** — escolha se quer buscar **todo o histórico** da
empresa ou **um mês específico**.

**4. Frequência** — escolha até 2: **Mensal** (fecha o mês anterior
completo), **Semanal** ou **Quinzenal** (mantêm o mês atual sempre
atualizado). Cada uma tem seu próprio dia e horário.

Ao clicar em **Concluir**, o programa já cadastra sozinho as tarefas no
Agendador do Windows — você não precisa abrir mais nada. A partir daí, ele
roda por conta própria, nos dias e horários escolhidos.

## Perguntas comuns

**"Certificado vencido" — o que faço?**
Renove o certificado com sua Autoridade Certificadora, ou remova a empresa
do cadastro (reabra o programa para editar). Se não quiser mais ver esse
aviso, marque "Não mostrar este aviso novamente" quando ele aparecer.

**Posso rodar o download manualmente, sem esperar o horário agendado?**
Sim — abra o assistente de novo a qualquer momento para revisar o cadastro;
para forçar uma busca imediata.

**Onde ficam as notas e relatórios?**
Na pasta que você escolheu no Passo 1, organizados por empresa e por mês.

**O programa mexe em alguma coisa da minha conta do Governo/certificado?**
Não. Ele só lê as notas via API oficial do Portal Nacional, usando seu
certificado para se autenticar — não emite, altera nem cancela nada.

## Para quem quer entender por dentro / contribuir

Veja [`TECNICO.md`](TECNICO.md) — arquitetura, as APIs usadas, como rodar
direto do código-fonte e como gerar o `.exe` você mesmo.

## Aviso legal

Este é um projeto independente, sem vínculo com a Receita Federal, a ENOTAS,
ou qualquer Autoridade Certificadora. Use por sua conta e risco — confira
sempre os relatórios antes de tomar decisões fiscais/contábeis com base
neles. Veja a licença [MIT](LICENSE): o software é fornecido "como está",
sem garantias.

## Contato

Feito por **Erika Duarte**.
[LinkedIn](https://www.linkedin.com/in/erika-duarte-tech/)

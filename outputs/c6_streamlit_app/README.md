# Auto C6BANK - Datafast for Argus

App desktop/Streamlit para receber a Extracao DTF CSV do C6 ou Banco Digio, aplicar filtros, cruzar telefones, cruzar BDD/DDB e gerar arquivos para Lemiti, Argus, SMS e auto-enviados.

## Banco automatico

O banco e identificado pelo nome do arquivo selecionado em `Extracao DTF CSV`.

- `c6` no nome aplica Banco C6 com coeficiente `0,0228`.
- `digio` ou `dígio` no nome aplica Banco Digio com coeficiente `0,02280`.
- A leitura ignora maiusculas/minusculas, acentos e separadores.
- Se o nome nao tiver `C6` nem `Digio`, o processamento e bloqueado.
- A area de filtros mostra o banco identificado e o coeficiente aplicado.
- No Digio, `Matricula` segue a mesma regra do campo `Beneficio` da base DTF.

## Fluxo

1. Selecione a Extracao DTF CSV.
2. Ajuste os filtros, incluindo `Periodo Telefone`.
3. Na aba `1. CPFs para Lemiti`, gere os CPFs sem telefone e a planilha de conferencia.
4. Na aba `2. Retorno Lemiti, XLSX e Argus`, selecione o retorno Lemiti e gere o XLSX final junto com o CSV Argus.
5. Na aba `3. SMS`, gere o arquivo CAMP de SMS e o arquivo auto-enviados.

O arquivo para Lemiti leva somente os CPFs que ficaram sem `TELEFONE BASE` apos a busca no HUB/base local. O retorno Lemiti concatena `DDD` + `TELEFONE` e preenche apenas esses vazios, sem alterar CPFs que ja tinham telefone da base.

## Interface guiada

- O desktop usa layout minimalista para caber melhor em 1366x768.
- Em 06/08/2026, o layout foi compactado novamente para notebooks menores, com cabecalho mais baixo, logo menor, filtros em menos linhas e area das abas ocupando melhor a altura disponivel.
- A aba `Finalizar Base e Argus` ancora as acoes no rodape da aba quando houver espaco vertical sobrando.
- O header mostra apenas o nome do app, uma frase curta e `Limpar Projeto`.
- Os cards permanentes foram removidos.
- Banco, UF, registros e coeficiente aparecem apenas depois da base carregada, em uma linha.
- A barra de progresso superior foi removida.
- As abas foram simplificadas para `Preparar para Lemiti`, `Finalizar Base e Argus` e `Gerar SMS`.
- Os filtros ficam em duas colunas.
- Os botoes principais ficam bloqueados ate uma base valida ser selecionada.
- A area de resultado so aparece depois que algum arquivo e gerado.
- As mensagens de erro evitam termos tecnicos e indicam o que corrigir.
- O envio para Argus exige confirmacao visual do banco, UF, campanha e arquivo.
- A campanha Argus e calculada pela pasta do mes atual e pela UF atual; a coluna `CAMPANHA` do CSV fica igual a campanha do nome do arquivo gerado.
- Depois de gerar o CSV, o desktop pergunta se deseja enviar a lista imediatamente para a Argus.
- Nenhuma dependencia visual nova foi adicionada; a versao desktop usa Tkinter/ttk.

## Icone

Atualizado em 05/08/2026.

- `app_icon.ico` combina Grupo Imperio, C6 Bank e Banco Digio em um simbolo integrado.
- O leao do Grupo Imperio e o elemento central e principal.
- C6 e Digio aparecem como marcas laterais, sem caixas brancas, para melhorar a leitura como icone.
- PNGs fonte: `app_icon_1024.png` e `app_icon_512.png`.
- O `.ico` inclui 16, 24, 32, 48, 64, 128 e 256 px.
- A interface desktop usa `brand_imperio.png` no cabecalho e `brand_c6.png` como selo quando uma base C6 e carregada.
- `brand_c6.png` foi recriado como PNG transparente e limpo, sem fundo quadriculado.
- A marca Digio aparece como wordmark discreto na interface; o arquivo fonte `brand_digio.webp` fica preservado no projeto.

## Colunas e calculos

O XLSX final mantem os cabecalhos tratados, cria `TELEFONE C6`, preenche `TELEFONE BASE` pelo HUB Consolidado e cruza `BENEFICIO` com `NB` na base estadual de DDB. Os arquivos derivados recebem `BDD`; o layout da Argus mantem o cabecalho `DDB`.

CPFs e telefones vindos do Excel como numero (`.0` no fim) sao normalizados antes do cruzamento. Telefones da base sao padronizados em 11 digitos.

O CSV de CPFs sem telefone usa `;` como separador para abrir corretamente no Excel em configuracao brasileira.

Ao selecionar uma nova base, os filtros voltam para os valores padrao para evitar reaproveitar taxa, troco ou prazo de uma campanha anterior sem perceber.

Quando houver varias bases DDB da mesma UF, o app consulta todas e junta os resultados. Arquivos vazios sao ignorados quando existem outras bases validas. Ele so bloqueia se o mesmo beneficio aparecer com DDBs diferentes.

- `COEF C6 = 0,0228`
- `COEF Digio = 0,02280`
- `BRUTO = PARCELA / COEF`
- `TROCO = BRUTO - SALDO DEV`

## Upload Argus

O campo `Tipo de operacao` controla o nome e o layout da Argus:

- `Refinanciamento`: usa `UF N REFIN INSS BANCO`.
- `Portabilidade`: seleciona automaticamente `SALA PORTABILIDADE` e usa `UF N PORTABILIDADE BANCO PARA QUALIBANKING`.

Exemplo: `SP 1 PORTABILIDADE DIGIO PARA QUALIBANKING.csv`.

Em `Finalizar Base e Argus`, a caixa `SALA` oferece `SALA MULTBANCOS`,
`PARCEIRO CLEITON` e `SALA PORTABILIDADE`. As duas primeiras usam o layout padrão. `SALA PORTABILIDADE` usa o
layout `44 - SALA PORTABILIDADE - OFC` e acrescenta `Cód. Espécie` ao CSV; se
a origem não trouxer espécie, o valor exportado é `999`.

Ordem do CSV de portabilidade:

```text
CAMPANHA;BENEFICIO;CPF;TELEFONE;DATA NASCIMENTO;CÓD. ESPÉCIE;TAXA;FEZ EM;FALTAM;PARCELA;DEVEDOR;TROCO;NOME;TELEFONE
```

- SALA MULTBANCOS: `POST https://argus.app.br/apiargus/2jovr4cpoyeb9bf/uploadmailing`.
- SALA PORTABILIDADE: `61vm1mnizndyeou`.
- PARCEIRO CLEITON: `ps4hqsyhyeobfiv`.
- Autenticacao: header `Token-Signature`.
- A Token-Signature fica em configuracao local segura ou variavel de ambiente; ela nao fica publicada no GitHub.
- Se a Argus responder `EndPoint invalido` para o hash salvo na maquina, o app busca o hash atual do layout `LAYOUT PADRAO C6` e tenta novamente.
- Se a conexao HTTPS do Python falhar em outra rede, o app tenta novamente usando o cliente HTTPS do Windows (`curl.exe`).

## Recompilar

Execute:

```powershell
outputs\c6_streamlit_app\recompilar_exe.bat
```

O executavel atualizado sera copiado para:

```text
outputs\c6_executavel\Auto C6BANK - Datafast for Argus.exe
```

# Auto C6BANK - Datafast for Argus

App desktop/Streamlit para receber a Extracao DTF CSV do C6 ou Banco Digio, aplicar filtros, cruzar telefones, cruzar BDD/DDB e gerar arquivos para Lemiti, Argus, SMS e auto-enviados.

## Banco automatico

O banco e identificado pelo nome do arquivo selecionado em `Extracao DTF CSV`.

- `c6` no nome aplica Banco C6 com coeficiente `0,0228`.
- `digio` ou `dígio` no nome aplica Banco Digio com coeficiente `0,022`.
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

## Interface guiada

- O desktop usa topo corporativo preto, detalhes dourados e cards permanentes de resumo.
- As abas foram simplificadas para `Preparar para Lemiti`, `Finalizar Base e Argus` e `Gerar SMS`.
- Os botoes principais ficam bloqueados ate uma base valida ser selecionada.
- A barra de progresso mostra a etapa atual: arquivo, filtros, processamento, Argus ou SMS.
- As mensagens de erro evitam termos tecnicos e indicam o que corrigir.
- O envio para Argus exige confirmacao visual do banco, UF, campanha e arquivo.
- Nenhuma dependencia visual nova foi adicionada; a versao desktop usa Tkinter/ttk.

## Colunas e calculos

O XLSX final mantem os cabecalhos tratados, cria `TELEFONE C6`, preenche `TELEFONE BASE` pelo HUB Consolidado e cruza `BENEFICIO` com `NB` na base estadual de DDB. Os arquivos derivados recebem `BDD`; o layout da Argus mantem o cabecalho `DDB`.

- `COEF C6 = 0,0228`
- `COEF Digio = 0,022`
- `BRUTO = PARCELA / COEF`
- `TROCO = BRUTO - SALDO DEV`

## Upload Argus

- Endpoint: `POST https://argus.app.br/apiargus/5ni7te02zzh1ka8/uploadmailing`.
- `hashSkill`: `5ni7te02zzh1ka8`.
- Autenticacao: header `Token-Signature`.
- A Token-Signature fica em configuracao local segura ou variavel de ambiente; ela nao fica publicada no GitHub.

## Recompilar

Execute:

```powershell
outputs\c6_streamlit_app\recompilar_exe.bat
```

O executavel atualizado sera copiado para:

```text
outputs\c6_executavel\Auto C6BANK - Datafast for Argus.exe
```

# Auto C6BANK - Datafast for Argus

Ferramenta para tratar bases C6, cruzar retorno Lemiti e gerar arquivos prontos para operacao no Argus.

## O que o sistema faz

- Filtra a extracao DTF do C6.
- Gera CSV de CPFs para consulta no Lemiti.
- Recebe o retorno Lemiti e gera, no mesmo fluxo, o XLSX tratado e o CSV da Discadora Argus.
- Preenche automaticamente a campanha por UF e proximo numero encontrado na pasta da Argus.
- Salva o CSV final na estrutura mensal:

```text
G:\Meu Drive\GRUPO IMPERIO\LEADS\ARGUS\DISC\MES ANO\UF
```

## Layout Discadora Argus

O CSV final para a discadora sai com as colunas:

```text
CAMPANHA;CPF;Taxa;FEZ EM;FALTAM; PARCELA ;TROCO;Telefone;Nome;Data Nascimento;TELEFONE C6
```

Regras principais:

- `Telefone` recebe o valor de `TELEFONE BASE`.
- `TELEFONE C6` e copiado exatamente da planilha tratada.
- `beneficio` e ignorado no CSV da Argus.
- Linhas ocultas por filtro no Excel nao entram no CSV.
- `Data Nascimento` sai em formato `dd/mm/aaaa`.
- `TROCO` sai em formato contabil brasileiro.

## Campanha automatica

O sistema identifica a UF no nome do retorno Lemiti. Exemplo:

```text
2022 refin c6 RS (1 ou 0 cartoes) sem RL pra disc
```

Depois busca na pasta da Argus a ultima campanha dessa UF e usa o proximo numero.

Exemplo:

```text
RS 6 REFIN INSS C6
```

O nome final fica:

```text
RS 6 REFIN INSS C6 2022 refin c6 RS (1 ou 0 cartoes) sem RL pra disc.csv
```

## Como usar

1. Abra o executavel `Auto C6BANK - Datafast for Argus.exe`.
2. Na aba `1. CPFs para Lemiti`, gere os CPFs sem telefone.
3. Consulte os CPFs no Lemiti.
4. Na aba `2. Retorno Lemiti, XLSX e Argus`, selecione o retorno Lemiti.
5. Confira a campanha sugerida e o caminho do CSV Argus.
6. Clique em `Gerar XLSX e CSV Argus`.

## Recompilar o executavel

Depois de alterar o codigo, rode:

```text
outputs\c6_streamlit_app\recompilar_exe.bat
```

O executavel atualizado sera copiado para:

```text
outputs\c6_executavel\Auto C6BANK - Datafast for Argus.exe
```

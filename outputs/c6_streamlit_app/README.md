# Auto C6BANK - Datafast for Argus

App Streamlit para receber a Extração DTF CSV, filtrar os registros e baixar um XLSX formatado.

Depois de gerar o CSV da Argus, a interface permite enviá-lo diretamente pelo endpoint de upload de mailing. A Token-Signature é digitada em campo de senha e não é salva pelo Streamlit.

## Upload Argus

- Operacao: Upload de Mailing por Arquivo CSV.
- Endpoint: `POST https://argus.app.br/apiargus/5ni7te02zzh1ka8/uploadmailing`.
- `hashSkill`: `5ni7te02zzh1ka8`.
- Skill/layout: `PADRAO` / `PADRAO DISC REFIN C6 2TEL`.
- Autenticacao: header `Token-Signature`.
- O Streamlit nao salva a Token-Signature; ela precisa ser digitada no envio.

## Rodar no VS Code

```powershell
cd C:\Users\imper\Documents\Codex\2026-07-03\es\outputs\c6_streamlit_app
python -m pip install -r requirements.txt
streamlit run app.py
```

## Filtros

- `FALTAM`: deixe vazio para todos ou digite um prazo, como `32`.
- `Regra FALTAM`: o padrão é `Maior que o prazo`, ou seja, `25` pega quem tem mais de 25 parcelas faltando. Também existe `Até o prazo` e `Exato`.
- `SALDO DEV`: sempre precisa ser maior que zero.
- `TAXA`: maior ou igual ao valor digitado.
- `TROCO`: maior que o valor digitado.
- `DATA DE NASC`: intervalo inclusivo.

## Fluxo Lemiti

1. Na aba `1. CPFs para Lemiti`, baixe o CSV de CPFs sem telefone.
2. Passe esse CSV no Lemiti.
3. Na aba `2. Retorno Lemiti e XLSX`, suba o retorno do Lemiti.
4. Baixe o XLSX final.

O retorno do Lemiti deve conter as colunas `CPF`, `DDD`, `Telefone`, `TipoTelefone`, `Ranking` e `Score`.

## Colunas geradas

O XLSX final mantém os cabeçalhos renomeados, ignora `cr_valor_contrato` e `cr_observacao`, cria `TELEFONE C6`, preenche `TELEFONE BASE` pela base de telefones e calcula:

- `COEF = 0,0228`
- `BRUTO = PARCELA / COEF`
- `TROCO = BRUTO - SALDO DEV`

Também dá para gerar direto pelo terminal:

As colunas `A`, `B`, `C`, `D`, `J`, `K` e `L` saem como número no Excel.

O campo `TELEFONE BASE` é preenchido pela primeira pasta encontrada:

`G:\Meu Drive\GRUPO IMPÉRIO\LEADS\TELEFONES\HUB CONSOLIDADO`

Se o caminho acima não existir, o sistema tenta automaticamente:

`H:\Meu Drive\GRUPO IMPÉRIO\LEADS\TELEFONES\HUB CONSOLIDADO`

Também há exportação de CSV com os CPFs filtrados que não têm telefone na base. O CSV sai com uma única coluna chamada `cpf`.

```powershell
python c6_processor.py --input "C:\Users\imper\Downloads\exportacao_dados_836_Exportacao_700_-_88k_c6_RS_1_ou_0_cartoes_sem.csv" --output "C:\Users\imper\Documents\Codex\2026-07-03\es\outputs\c6_streamlit_app\c6_filtrado.xlsx" --faltam 25 --faltam-modo maior --taxa-min 1.66 --troco-min 0 --data-inicial 01/01/1954 --data-final 31/12/1995
```

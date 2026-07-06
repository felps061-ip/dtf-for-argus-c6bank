# C6 para XLSX

App Streamlit para receber o CSV padrão extraído do C6, filtrar os registros e baixar um XLSX formatado.

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

## Colunas geradas

O XLSX final mantém os cabeçalhos renomeados, ignora `cr_valor_contrato` e `cr_observacao`, cria `TELEFONE C6`, deixa `TELEFONE BASE` vazio e calcula:

- `COEF = 0,0228`
- `BRUTO = PARCELA / COEF`
- `TROCO = BRUTO - SALDO DEV`

Também dá para gerar direto pelo terminal:

As colunas `A`, `B`, `C`, `D`, `J`, `K` e `L` saem como número no Excel.

```powershell
python c6_processor.py --input "C:\Users\imper\Downloads\exportacao_dados_836_Exportacao_700_-_88k_c6_RS_1_ou_0_cartoes_sem.csv" --output "C:\Users\imper\Documents\Codex\2026-07-03\es\outputs\c6_streamlit_app\c6_filtrado.xlsx" --faltam 25 --faltam-modo maior --taxa-min 1.66 --troco-min 0 --data-inicial 01/01/1954 --data-final 31/12/1995
```

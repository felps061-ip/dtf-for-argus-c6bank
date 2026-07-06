# C6 para XLSX - Executavel

Use `C6_para_XLSX.exe` com duplo clique.

1. Clique em `Escolher` no campo `CSV C6`.
2. Escolha onde salvar o arquivo em `Salvar XLSX em`.
3. Preencha os filtros:
   - `FALTAM`: pode ficar vazio ou receber um prazo, como `32`.
   - `Regra FALTAM`: o padrao e `Maior que o prazo`, ou seja, `25` pega quem tem mais de 25 parcelas faltando. Tambem existe `Ate o prazo` e `Exato`.
   - `TAXA minima`: filtra taxa maior ou igual.
   - `TROCO maior que`: filtra troco maior que o valor digitado.
   - `Data inicial` e `Data final`: use formato `dd/mm/aaaa`.
4. Clique em `Gerar XLSX`.

O arquivo final sai com os cabecalhos tratados, `TELEFONE C6`, `TELEFONE BASE` vazio, `COEF`, `BRUTO` e `TROCO`, com formatos de numero, data e moeda.

As colunas `A`, `B`, `C`, `D`, `J`, `K` e `L` saem como numero no Excel.

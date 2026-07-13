# Auto C6BANK - Datafast for Argus - Executavel

Use `Auto C6BANK - Datafast for Argus.exe` com duplo clique.

Depois de gerar o CSV Argus, clique em `Enviar este CSV para a Argus`. Na primeira utilização, use `Configurar acesso Argus`; a Token-Signature fica protegida pelo Windows para o usuário da máquina e não é gravada no GitHub.

Upload Argus configurado:

- Endpoint: `POST https://argus.app.br/apiargus/5ni7te02zzh1ka8/uploadmailing`
- `hashSkill`: `5ni7te02zzh1ka8`
- Skill/layout: `PADRAO` / `PADRAO DISC REFIN C6 2TEL`
- Todas as UFs usam o mesmo `hashSkill`

1. Clique em `Escolher` no campo `Extração DTF CSV`.
2. Escolha onde salvar o arquivo em `Salvar em`.
3. Preencha os filtros:
   - `FALTAM`: pode ficar vazio ou receber um prazo, como `32`.
   - `Regra FALTAM`: o padrao e `Maior que o prazo`, ou seja, `25` pega quem tem mais de 25 parcelas faltando. Tambem existe `Ate o prazo` e `Exato`.
   - `TAXA minima`: filtra taxa maior ou igual.
   - `TROCO maior que`: filtra troco maior que o valor digitado.
   - `Data inicial` e `Data final`: use formato `dd/mm/aaaa`.
4. Na aba `1. CPFs para Lemiti`, clique em `Extrair CPFs para Lemiti (.csv)`.
5. Passe esse CSV no Lemiti.
6. Na aba `2. Retorno Lemiti e XLSX`, selecione o retorno do Lemiti.
7. Clique em `Gerar XLSX`.

O arquivo final sai com os cabecalhos tratados, `TELEFONE C6`, `TELEFONE BASE` preenchido pela base, `COEF`, `BRUTO` e `TROCO`, com formatos de numero, data e moeda.

As colunas `A`, `B`, `C`, `D`, `J`, `K` e `L` saem como numero no Excel.

O campo `TELEFONE BASE` e preenchido pela primeira pasta encontrada:

`G:\Meu Drive\GRUPO IMPÉRIO\LEADS\TELEFONES\HUB CONSOLIDADO`

Se o caminho acima nao existir, o sistema tenta automaticamente:

`H:\Meu Drive\GRUPO IMPÉRIO\LEADS\TELEFONES\HUB CONSOLIDADO`

Use `Extrair CPFs para Lemiti (.csv)` para gerar um CSV somente com a coluna `cpf`, contendo os CPFs filtrados que nao foram encontrados na base de telefones.

O retorno do Lemiti deve conter as colunas:

`CPF`, `DDD`, `Telefone`, `TipoTelefone`, `Ranking`, `Score`

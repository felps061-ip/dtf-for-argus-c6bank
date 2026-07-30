# Auto C6BANK - Datafast for Argus

Ferramenta para tratar bases C6, cruzar retorno Lemiti e gerar arquivos prontos para operacao no Argus.

## O que o sistema faz

- Filtra a extracao DTF do C6.
- Gera CSV de CPFs para consulta no Lemiti, incluindo `beneficio` e `BDD`.
- Recebe o retorno Lemiti e gera, no mesmo fluxo, o XLSX tratado e o CSV da Discadora Argus.
- Identifica a UF da extração, cruza `BENEFÍCIO` com `NB` na base estadual e preenche `BDD`.
- Envia o CSV recém-gerado diretamente para a API Argus, mediante confirmação do usuário.
- Gera todos os XLSX/CSV derivados com a coluna `BDD`, inclusive conferência, SMS e auto-enviados.
- Tem botao `Reiniciar / F5` para limpar os campos e voltar a tela ao inicio sem fechar o app.
- O executavel usa a logo `app_icon.ico` como icone do app.
- Preenche automaticamente a campanha por UF e proximo numero encontrado na pasta da Argus.
- Salva o CSV final na estrutura mensal:

```text
G:\Meu Drive\GRUPO IMPERIO\LEADS\ARGUS\DISC\MES ANO\UF
```

Se o caminho em `G:` nao existir, o sistema tenta a mesma estrutura em `H:`.
Exemplo para julho de 2026:

```text
G:\Meu Drive\GRUPO IMPERIO\LEADS\ARGUS\DISC\JULHO 2026
H:\Meu Drive\GRUPO IMPERIO\LEADS\ARGUS\DISC\JULHO 2026
```

## Layout Discadora Argus

O CSV final para a discadora sai com as colunas:

```text
CAMPANHA;CPF / CNPJ;Taxa;FEZ EM;FALTAM;PARCELA;TROCO;Telefone;Nome;Data Nascimento;Telefone;DDB;Beneficio
```

Regras principais:

- `Telefone` recebe o valor de `TELEFONE BASE`.
- O primeiro `Telefone` recebe `TELEFONE BASE`; o segundo recebe `TELEFONE C6`.
- `Beneficio` é cruzado com `NB` no arquivo estadual localizado em `G:\Meu Drive\GRUPO IMPÉRIO\LEADS\INSS DT DDB`.
- `DDB` recebe a data encontrada; benefícios não localizados permanecem vazios e são contabilizados na tela.
- Bases com DDBs conflitantes para o mesmo benefício bloqueiam a geração para evitar associação incorreta.
- Linhas ocultas por filtro no Excel nao entram no CSV.
- `Data Nascimento` sai em formato `dd/mm/aaaa`.
- `TROCO` sai em formato contabil brasileiro.
- O filtro `Periodo Telefone` abre sugerido como `01/01/2026` e sempre e aplicado na busca do HUB.

## Layout SMS

A aba `3. SMS` gera um XLSX com as colunas:

```text
Telefone;Nome;N;BDD
```

Regras principais:

- `Telefone` vem de `TELEFONE BASE` da base final tratada.
- Quando houver mais de um telefone na mesma linha, o sistema cria uma linha para cada telefone.
- Telefones repetidos sao removidos do arquivo CAMP de SMS, mantendo a primeira ocorrencia.
- `Telefone` sai em formato especial de 11 digitos: `00000000000`.
- `Nome` recebe o nome completo vinculado ao telefone.
- `N` recebe apenas o primeiro nome da pessoa.
- `BDD` vem da mesma base estadual usada no arquivo da Argus.
- O XLSX do SMS e salvo automaticamente em:

```text
G:\Meu Drive\GRUPO IMPERIO\LEADS\DISPARO PRO\Enviados\MES ANO
```

- Se a pasta do mes ainda nao existir, ela e criada automaticamente.
- O nome segue o padrao `CAMP N quantidade refin c6 UF N (1 ou 0 cartoes) sem RL pra disc.xlsx`, usando o proximo numero `CAMP` encontrado na pasta do mes.
- Ao gerar SMS, o sistema tenta criar uma planilha Google real de auto-enviados na pasta:

```text
G:\Meu Drive\GRUPO IMPERIO\LEADS\DISPARO PRO\auto-enviados
```

- A pasta Google usada como destino ja fica configurada no app pelo ID `1ZBgiFOqO18Eez_LacdDAD8skvjD3tE_4`.
- Se precisar trocar esse destino no futuro, crie um arquivo `google_auto_enviados_folder_id.txt` ao lado do executavel ou em `%USERPROFILE%\.auto_c6bank\` com o novo ID da pasta.
- Para ativar a API, coloque o OAuth Client do Google Cloud como `google_client_secret.json` ao lado do executavel ou em `%USERPROFILE%\.auto_c6bank\google_client_secret.json`.
- Na primeira geracao, o navegador abre para login do Google; depois o app salva o token em `%USERPROFILE%\.auto_c6bank\token.json`.
- Se o Google ainda nao estiver configurado ou falhar, o sistema salva um XLSX de reserva na pasta `auto-enviados` e mostra uma mensagem curta explicando o motivo.
- Esse segundo arquivo nao recebe o prefixo `CAMP`; o nome segue o padrao `quantidade refin c6 UF N (1 ou 0 cartoes) sem RL pra disc`.
- Esse segundo arquivo usa as colunas `cpf`, `beneficio`, `fez em`, `parcela`, `taxa`, `nome`, `troco`, `TELEFONE BASE` e `BDD`.

## Acesso Argus

- A Token-Signature fica em configuracao local segura ou variavel de ambiente; ela nao fica publicada no GitHub.
- O `hashSkill` usado no envio permanece fixo em `5ni7te02zzh1ka8`.
- O acesso Argus pode ser salvo localmente na maquina de uso, sem expor a chave no codigo fonte.

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
7. Garanta que o acesso a Argus esteja configurado localmente na maquina.
8. Clique em `Enviar este CSV para a Argus` e confira o `idArquivo` retornado.
9. Para SMS, use a aba `3. SMS` e clique em `Gerar XLSX SMS`.

## Upload pela API Argus

Detalhes confirmados pela documentacao/API Argus:

- Operacao usada: Upload de Mailing por Arquivo CSV.
- Endpoint real: `POST https://argus.app.br/apiargus/5ni7te02zzh1ka8/uploadmailing`.
- `hashSkill` confirmado: `5ni7te02zzh1ka8`.
- O `hashSkill`/caminho permanece em `5ni7te02zzh1ka8`; o layout é administrado diretamente na Argus.
- Todas as UFs usam o mesmo `hashSkill`.
- Autenticacao: header `Token-Signature`.
- O CSV e enviado como `multipart/form-data`.
- A Token-Signature usada pelo envio fica fora do codigo publicado e deve ser configurada localmente.
- O botão de envio só é habilitado depois que o CSV foi gerado e salvo com sucesso.
- O programa registra localmente o hash do arquivo enviado e pede confirmação antes de reenviar o mesmo conteúdo.

## Recompilar o executavel

Depois de alterar o codigo, rode:

```text
outputs\c6_streamlit_app\recompilar_exe.bat
```

O executavel atualizado sera copiado para:

```text
outputs\c6_executavel\Auto C6BANK - Datafast for Argus.exe
```


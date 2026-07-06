# Instalador - Auto C6BANK - Datafast for Argus

Use `Instalador Auto C6BANK - Datafast for Argus.exe` para instalar o programa na maquina.

O instalador:

- copia `Auto C6BANK - Datafast for Argus.exe` para `%LOCALAPPDATA%\Programs\Auto C6BANK - Datafast for Argus`;
- cria atalho na Area de Trabalho;
- cria atalhos no Menu Iniciar;
- cria um desinstalador simples.

Nao precisa executar como administrador, pois a instalacao e feita somente para o usuario atual do Windows.

Para teste silencioso:

```powershell
.\Instalador Auto C6BANK - Datafast for Argus.exe --silent --install-dir C:\Temp\AutoC6Teste --no-shortcuts --no-registry
```

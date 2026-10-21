# QA-automatizado

Bot de testes em Python (Playwright + openpyxl) que conversa com o chat de atendimento de um site, faz uma lista de perguntas e registra as respostas e prints em uma planilha.

## Como o bot funciona

1. Lê `venv/perguntas.xlsx`. Cada linha é uma conversa; as colunas `Pergunta 1`, `Pergunta 2`, ... trazem as perguntas em ordem. Células vazias são ignoradas.
2. Para cada linha, abre uma sessão nova do navegador em `URL_DO_SITE`, abre o chat e envia as `MENSAGENS_INICIAIS` (ex.: nome e confirmação que o bot pede), sem registrá-las.
3. Envia as perguntas na ordem, na mesma conversa. Depois de cada uma, espera a resposta terminar (nova bolha do bot, sem "digitando..." e sem mudar por `ESTAVEL_S` segundos) e tira um print da janela do chat.
4. Grava tudo na pasta `output/` (a planilha em `output/resultado.xlsx`), a cada pergunta, então nada se perde se o script cair.

Se uma pergunta falha (timeout, seletor não encontrado), o erro é registrado, as perguntas restantes daquela linha entram como `NÃO EXECUTADA` e o bot segue para a próxima linha.

### Saídas (pasta `output/`)

Tudo que o bot gera fica na pasta **`output/`**, que o bot cria sozinho, está no `.gitignore` e **não vai para o git**:

| Arquivo em `output/` | Conteúdo |
|---|---|
| `resultado.xlsx` | linha, nº da pergunta, pergunta, resposta do bot, caminho do print, status (`OK`/`ERRO`) e tempo de resposta em segundos |
| `prints/linha_<N>/pergunta_<M>.png` | print do chat após cada pergunta (`N` = linha da planilha, contando o cabeçalho como 1) |
| `videos/`, `traces/` | só se `GRAVAR_VIDEO`/`GRAVAR_TRACE` estiverem ligados |

A pasta **`output-example/`** mostra como é uma saída (planilha com dados ilustrativos, incluindo um `ERRO` e uma `NÃO EXECUTADA`, e dois prints de exemplo). Ela **vai para o git**.

## Pastas `venv/` e `venv-example/` (dados locais)

Os dados que mudam de pessoa para pessoa ficam na pasta **`venv/`**, que está no `.gitignore` e **não vai para o git**:

| Arquivo em `venv/` | O que é |
|---|---|
| `perguntas.xlsx` | planilha de entrada com as conversas |
| `.env` | variáveis de ambiente: `NOME` e `MENSAGENS_INICIAIS` |
| `*.html` | Playwright Inspector salvo, se quiser guardá-lo |

A pasta **`venv-example/`** é o modelo e **vai para o git**. Contém um `perguntas.xlsx` de exemplo, o `.env.example` e um `LEIA-ME.txt`. Para começar:

```bash
cp -r venv-example venv
mv venv/.env.example venv/.env
# edite venv/.env e venv/perguntas.xlsx
```

Se a pasta `venv/` não existir, o bot para com uma mensagem explicando isso.

### Variáveis do `venv/.env`

| Variável | O que é |
|---|---|
| `NOME` | nome que o bot de atendimento pede no início da conversa |
| `MENSAGENS_INICIAIS` | mensagens enviadas no início de toda conversa, separadas por `\|`; `{nome}` vira o valor de `NOME`. Deixe vazio se o bot não pedir nada |

Exemplo: `MENSAGENS_INICIAIS={nome}|Sim` envia o nome e depois "Sim".

> `venv/` (dados) **não** é o ambiente virtual do Python. O ambiente virtual é `.venv/` (com ponto), também ignorado pelo git.

## Começando do zero em outro computador

### 1. Pré-requisitos

- **Python 3.10 ou mais novo** (`python3 --version`). No Windows, instale em python.org marcando "Add Python to PATH".
- **Git** (`git --version`).
- Conexão com a internet (o Playwright baixa o Chromium, cerca de 150 MB).
- Em Linux, o `venv` do Python pode vir separado: `sudo apt install python3-venv` (Debian/Ubuntu).

### 2. Baixar o projeto

```bash
git clone https://github.com/jvfsouza/QA-automatizado.git
cd QA-automatizado
```

### 3. Criar o ambiente virtual e instalar as dependências

Linux / macOS:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
playwright install chromium
```

Windows (PowerShell):

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
playwright install chromium
```

Se o Chromium abrir com erro de bibliotecas faltando no Linux, rode `playwright install-deps chromium` (pede `sudo`).

### 4. Criar a pasta de dados `venv/`

A pasta `venv/` não vem no git, então crie-a a partir do modelo:

```bash
cp -r venv-example venv
mv venv/.env.example venv/.env
```

No Windows: `Copy-Item -Recurse venv-example venv` e `Rename-Item venv\.env.example .env`.

Depois edite:

- `venv/.env`: `NOME` e `MENSAGENS_INICIAIS` (veja a seção abaixo).
- `venv/perguntas.xlsx`: suas conversas, uma por linha, nas colunas `Pergunta 1`, `Pergunta 2`, ...

### 5. Configurar o site

No topo do `bot_teste.py`, ajuste `URL_DO_SITE` e os seletores `SEL_*` (veja "Usar em outro site"). Os valores que vêm no projeto são do chat do Brazino777.

### 6. Rodar um teste curto

Deixe só 1 conversa na planilha e `HEADLESS = False` no `bot_teste.py`, para ver o navegador abrindo:

```bash
python bot_teste.py
```

O resultado aparece em `output/resultado.xlsx` e os prints em `output/prints/`. Se funcionou, volte `HEADLESS = True` (se quiser) e rode com a planilha completa.

### 7. VS Code (opcional)

Abra a pasta no VS Code e selecione o interpretador `.venv` (`Ctrl+Shift+P` → **Python: Select Interpreter**). O projeto traz `.vscode/settings.json` apontando para ele.

Em **dias seguintes** não precisa reinstalar nada: entre na pasta, ative o ambiente (`source .venv/bin/activate`) e rode `python bot_teste.py`. Para atualizar o código, `git pull`; se o `requirements.txt` mudou, rode `pip install -r requirements.txt` de novo.

### Problemas comuns

| Erro | Causa e solução |
|---|---|
| `ModuleNotFoundError: No module named 'openpyxl'` | O Python em uso não é o do `.venv`. Ative o ambiente ou use `.venv/bin/python bot_teste.py` |
| `Executable doesn't exist ... playwright install` | Falta baixar o navegador: `playwright install chromium` |
| `Pasta venv/ não encontrada` | Falta o passo 4 |
| `Timeout ... exceeded` nos seletores | Os `SEL_*` não batem com o site; refaça com `playwright codegen` |
| `externally-managed-environment` ao instalar | Instalou fora do `.venv`; ative o ambiente virtual antes do `pip install` |

## Comandos

```bash
# rodar o bot
.venv/bin/python bot_teste.py

# descobrir seletores do chat de um site (grava os cliques e mostra o código)
.venv/bin/playwright codegen https://site.com/

# rodar passo a passo com o Playwright Inspector
PWDEBUG=1 .venv/bin/python bot_teste.py

# abrir o trace de uma conversa (com GRAVAR_TRACE = True)
.venv/bin/playwright show-trace output/traces/linha_2.zip
```

Se rodar pelo VS Code, selecione o interpretador `.venv` (`Ctrl+Shift+P` → **Python: Select Interpreter**). O projeto já traz `.vscode/settings.json` apontando para ele.

## Configuração

Tudo fica no bloco de configuração no topo do `bot_teste.py`.

### Execução

| Variável | O que faz |
|---|---|
| `URL_DO_SITE` | página onde está o chat |
| `ENTRADA` / `SAIDA` | planilha de perguntas (em `venv/`) / planilha de resultado (em `output/`) |
| `PASTA_SAIDA` / `PASTA_PRINTS` | pasta das saídas (`output/`) / onde salvar os prints (`output/prints/`) |
| `TIMEOUT_MS` | limite total por pergunta (do envio até a resposta terminar) e por navegação |
| `ESTAVEL_S` | a resposta só é dada como pronta se o texto não mudar por este tempo |
| `HEADLESS` | `True` roda sem abrir janela; `False` mostra o navegador |
| `SLOW_MO_MS` | pausa entre ações (ex.: `500`), útil com `HEADLESS = False` |
| `GRAVAR_VIDEO` / `GRAVAR_TRACE` | gravam vídeo e trace de cada conversa, para investigar falhas |

### Seletores do chat

São seletores CSS/Playwright. Descubra-os com `playwright codegen` e o botão **Pick locator** do Inspector. Prefira `id`, `aria-label`, `title` e `data-*`; evite classes geradas (ex.: `sc-8173l2-0`), que mudam a cada atualização.

| Variável | O que é |
|---|---|
| `SEL_IFRAME` | iframe do widget |
| `SEL_ABRIR_CHAT` | botão que abre o chat, fora do iframe (`""` se já abre sozinho) |
| `SEL_INPUT` | campo de texto |
| `SEL_ENVIAR` | botão enviar |
| `SEL_MSG_BOT` | cada bolha de mensagem do bot (não pode pegar as do usuário) |
| `SEL_DIGITANDO` | indicador "digitando..." (`""` se não existir) |
| `SEL_JANELA` | elemento fotografado nos prints |

### Comportamento do bot de atendimento

| Variável | O que é |
|---|---|
| `MENSAGENS_INICIAIS` | mensagens enviadas no início de toda conversa, sem registrar (ex.: nome e "Sim"). Vêm do `venv/.env` (`NOME` e `MENSAGENS_INICIAIS`); vazio se o bot não pedir nada |
| `TEXTOS_INTERMEDIARIOS` | textos curtos como "um momento" que o bot manda antes da resposta real; enquanto a última bolha for assim, o bot continua esperando |

## Usar em outro site

1. Troque `URL_DO_SITE`.
2. Rode o `codegen`, abra o chat, envie uma mensagem e copie os seletores para as variáveis `SEL_*`.
3. Veja o que o bot do site faz antes de aceitar perguntas e ajuste `MENSAGENS_INICIAIS` (no `venv/.env`) e `TEXTOS_INTERMEDIARIOS`.
4. Teste com uma planilha de 1 conversa e `HEADLESS = False` antes de rodar tudo.

O script assume que o chat está **dentro de um iframe**. Chat sem iframe ou que responde só com botões precisa de adaptação.

## Limitações conhecidas

- Banner de cookies ou captcha que bloqueie o clique no chat não são tratados.
- Pedido de transferência para humano pode estourar o `TIMEOUT_MS` e ser registrado como `ERRO`.
- A resposta gravada inclui as bolhas intermediárias ("Um momento, estou verificando.") junto da resposta final.
- Não deixe `output/resultado.xlsx` aberto no Excel/LibreOffice durante a execução: o salvamento falha (o bot só avisa no terminal e continua).
- `perguntas.xlsx` aberto no LibreOffice podem converter `=...` em fórmula ao salvar.

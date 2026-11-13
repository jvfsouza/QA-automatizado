# AGENTS.md

Guia para assistentes de IA (e pessoas) que vão ler, rodar ou alterar este projeto. Para o passo a passo completo de instalação e configuração, veja o `README.md`. O projeto está em português; mantenha comentários, mensagens e documentação em português.

## O que é

Um bot de QA em Python (Playwright + openpyxl) que testa o **chat de atendimento de um site**. Ele lê perguntas de uma planilha, conversa com o chat do site e grava as respostas, tempos e prints. Serve para verificar como o bot de atendimento responde a perguntas conhecidas.

O site configurado hoje é `https://brazino777.com/` (widget Zendesk em iframe, bot "Mía", uma IA). O código é genérico: outro site exige trocar `URL_DO_SITE` e os seletores `SEL_*`.

## Estrutura

```
bot_teste.py        # todo o código, em um único arquivo (decisão deliberada: manter simples)
requirements.txt    # playwright, openpyxl, python-dotenv
README.md           # documentação para pessoas
AGENTS.md           # este arquivo
FORMATO_PLANILHA.md # como criar o perguntas.xlsx
PROMPT_GERAR_PLANILHA.md # prompt para uma IA gerar o perguntas.xlsx
venv/               # DADOS LOCAIS, fora do git: perguntas.xlsx, .env, HTML do Playwright Inspector
venv-example/       # modelo de venv/ (vai para o git): perguntas.xlsx, .env.example
output/             # SAÍDAS do bot, fora do git: resultado.xlsx, prints/, videos/, traces/
output-example/     # modelo de output/ (vai para o git)
.venv/              # ambiente virtual do Python (fora do git)
```

Atenção ao nome: **`venv/` guarda dados, não é o ambiente virtual do Python**. O ambiente virtual é `.venv/` (com ponto).

## Como rodar

```bash
source .venv/bin/activate          # Linux/macOS; no Windows: .venv\Scripts\Activate.ps1
python bot_teste.py                # ou, sem ativar: .venv/bin/python bot_teste.py
python bot_teste.py --paralelo 3   # 3 conversas (linhas) ao mesmo tempo
```

Primeira vez em um computador novo:

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
playwright install chromium
cp -r venv-example venv && mv venv/.env.example venv/.env
```

Rode **sempre a partir da raiz do projeto**: os caminhos `venv/` e `output/` são relativos ao diretório atual. Se a pasta `venv/` não existir, o script encerra com uma mensagem explicando.

Ferramentas úteis:

```bash
playwright codegen https://site.com/        # descobrir seletores
PWDEBUG=1 python bot_teste.py               # passo a passo no Playwright Inspector
playwright show-trace output/traces/linha_2.zip
```

## Entradas

Formato completo da planilha em `FORMATO_PLANILHA.md`; prompt para gerá-la com uma IA em `PROMPT_GERAR_PLANILHA.md`.

- `venv/perguntas.xlsx`: cada linha é uma conversa; colunas `Pergunta 1`, `Pergunta 2`, ... em ordem cronológica. Células vazias são ignoradas. O número da pergunta vem do cabeçalho, não da posição.
- `venv/.env`:
  - `NOME`: nome que o bot do site pede no início.
  - `LOGIN_EMAIL` / `LOGIN_SENHA`: usados por `entrar()` quando o site redireciona para `/auth/` (login a cada conversa; seletores `SEL_LOGIN_*`). Nunca commitar.
  - `MENSAGENS_INICIAIS`: mensagens enviadas no começo de toda conversa, separadas por `|`; `{nome}` vira o valor de `NOME`. Exemplo: `{nome}|Sim`. Vazio = não envia nada.

## Saídas (`output/`)

- `resultado.xlsx`: linha, nº da pergunta, pergunta, resposta do bot, caminho do print, status (`OK`/`ERRO`/`TRANSBORDO`) e tempo de resposta em segundos. É salvo **a cada pergunta**.
- `prints/linha_<N>/pergunta_<M>.png`: `N` é a linha da planilha (cabeçalho = linha 1, então a primeira conversa é `linha_2`).
- `videos/`, `traces/`: só com `GRAVAR_VIDEO` / `GRAVAR_TRACE` ligados.

## Como o código funciona

Fluxo em `main()` → `conversar()` → `trocar_mensagem()`:

1. Para cada linha da planilha, cria um **contexto novo** do navegador (sessão limpa), abre `URL_DO_SITE` e clica em `SEL_ABRIR_CHAT`.
2. Espera o campo de texto aparecer e a saudação inicial estabilizar.
3. Envia as `MENSAGENS_INICIAIS` (sem registrar), depois cada pergunta da linha, na mesma conversa.
4. Para cada pergunta, `trocar_mensagem()` preenche o campo, clica em enviar, espera a **primeira nova bolha do bot** com `expect(...)` e depois chama `esperar_estavel()`.
5. `esperar_estavel()` faz polling (200 ms) e só considera a resposta pronta quando o texto das bolhas novas **não muda por `ESTAVEL_S` segundos**, o indicador "digitando" (`SEL_DIGITANDO`) não está visível e a última bolha não é uma mensagem intermediária (`TEXTOS_INTERMEDIARIOS`, ex.: "Um momento, estou verificando.").
6. Tira o print do `SEL_JANELA` e registra a linha no Excel.
7. Com `--paralelo N` (ou `PARALELO`), `main()` cria um `multiprocessing.Pool` (contexto `spawn`) e cada linha roda em `trabalhador()` → `executar_linha()`, um processo com navegador próprio. Os processos filhos só enviam resultados por uma fila; **o processo principal é o único que grava o Excel** (`salvar()`, sempre ordenado por linha e nº da pergunta). Com `N = 1` as linhas rodam em sequência no mesmo processo.
8. Depois de cada pergunta, se o fim da resposta casa com `RESPOSTAS_AUTOMATICAS` (`resposta_automatica()`), envia a resposta automática e repete, até `MAX_AUTOMATICAS`. Cada envio é registrado como `[auto] ...` com nº `3.1`, `3.2`... (print `pergunta_3.1.png`).
9. Se a resposta contém algum trecho de `TEXTOS_TRANSBORDO` (`transbordou()`), imprime um aviso, registra `TRANSBORDO`, marca as demais perguntas da linha como `NÃO EXECUTADA (...)` e, no fim, lista as linhas com transbordo.
10. Se uma pergunta falha (timeout, seletor, etc.), registra `ERRO` com o motivo, marca as demais perguntas da linha como `NÃO EXECUTADA (...)` e segue para a próxima linha.

Funções principais: `ler_conversas`, `esperar_estavel`, `trocar_mensagem`, `limpar` (remove o rótulo "Mía diz:"), `intermediaria`, `tirar_print`, `celula_segura` (remove caracteres ilegais, limita o tamanho, evita fórmulas).

## Configuração

Tudo no bloco de configuração no topo de `bot_teste.py`:

- **Execução:** `URL_DO_SITE`, `TIMEOUT_MS` (limite total por pergunta), `ESTAVEL_S`, `HEADLESS`, `TELA_LARGURA`/`TELA_ALTURA` (grade de janelas com `--paralelo` e `HEADLESS = False`; vagas distribuídas por uma fila em `trabalhador()`), `SLOW_MO_MS`, `GRAVAR_VIDEO`, `GRAVAR_TRACE`.
- **Seletores do chat:** `SEL_IFRAME`, `SEL_ABRIR_CHAT`, `SEL_INPUT`, `SEL_ENVIAR`, `SEL_MSG_BOT`, `SEL_DIGITANDO`, `SEL_JANELA`.
- **Comportamento do bot de atendimento:** `MENSAGENS_INICIAIS` (vem do `.env`), `TEXTOS_INTERMEDIARIOS`, `TEXTOS_TRANSBORDO`, `RESPOSTAS_AUTOMATICAS`, `MAX_AUTOMATICAS`.

Para outro site: trocar `URL_DO_SITE`, refazer os `SEL_*` com `playwright codegen` e ajustar `MENSAGENS_INICIAIS` e `TEXTOS_INTERMEDIARIOS` conforme o que o bot do site faz antes de responder. O código assume que o chat está **dentro de um iframe** (`page.frame_locator`).

## Convenções e decisões (respeite ao alterar)

- **Um único arquivo** (`bot_teste.py`) e código simples. Não quebre em pacotes nem adicione frameworks sem pedido.
- **Sem `sleep` fixo.** Esperas usam `expect(...)` com timeout; a única espera por polling é `esperar_estavel()`, necessária para detectar que o texto parou de mudar.
- **API síncrona do Playwright.** Não é segura entre threads; o paralelismo é por **processos** (`--paralelo`), com um único processo gravando o `.xlsx`. Não use threads nem faça filhos gravarem a planilha. Como o contexto é `spawn`, cada filho reimporta `bot_teste.py` com a configuração do arquivo: alterar variáveis do módulo em tempo de execução não chega aos filhos.
- **Seletores estáveis.** Prefira `id`, `aria-label`, `title`, `data-*`, `aria-live`. Evite classes geradas (`sc-8173l2-0`, `StyledButton-...`), que mudam a cada atualização do widget.
- **Dados sensíveis/locais** ficam em `venv/` e `output/`, ambos no `.gitignore`. Não commite planilhas reais, prints ou `.env`. Exemplos vão em `venv-example/` e `output-example/`.
- Os `*-example/` usam `.env.example` (e não `.env`) porque o `.gitignore` ignora qualquer arquivo chamado `.env`.
- Ao mudar variáveis de configuração, caminhos ou comportamento, **atualize o `README.md`** e este arquivo.

## Armadilhas conhecidas

- O bot do Brazino777 pede nome e confirmação antes de aceitar perguntas; sem `MENSAGENS_INICIAIS` a primeira pergunta é respondida com "qual o seu nome?".
- As respostas levam cerca de 20 s e vêm em várias bolhas (a primeira costuma ser "Um momento, estou verificando."). O texto gravado junta todas as bolhas novas.
- O indicador de digitação é uma bolha de texto ("Mía está a escrever"), não um elemento separado; `SEL_MSG_BOT` a exclui com `:not(:has-text('a escrever'))`.
- Bolhas do usuário e do bot são separadas por `aria-live` (`off` para o usuário, `polite` para o bot). Usar `has-text('diz:')` pega as duas.
- Células que começam com `=` viram fórmula no Excel; o script força texto literal. Abrir o `perguntas.xlsx` no LibreOffice e salvar pode reconverter.
- Não deixe `output/resultado.xlsx` aberto no Excel/LibreOffice durante a execução: o salvamento falha (o bot só avisa no terminal e continua).
- Em paralelo, o site pode limitar ou bloquear sessões simultâneas; os tempos de resposta medidos ficam menos confiáveis. Use `--paralelo 1` para medir tempo.
- O bot da Magalu anuncia "vou transferir" e depois coleta dados (nº do pedido, "Sim"/"Não", data) antes do transbordo real ("TU FOSTE TRANSBORDADO" no ambiente de testes). Por isso `TEXTOS_TRANSBORDO` só tem a transferência de fato, e `RESPOSTAS_AUTOMATICAS` responde a coleta. Palavras soltas como "atendente" dão falso transbordo.
- Não tratados: banner de cookies/captcha que bloqueie o clique, chat sem iframe, chat só com botões, transbordo em que o bot fica mudo ou usa uma frase fora de `TEXTOS_TRANSBORDO` (estoura o timeout).

## Antes de dar o trabalho por feito

- `python -m py_compile bot_teste.py` deve passar.
- Teste ponta a ponta com uma planilha de 1 conversa (2 perguntas bastam), de preferência com `HEADLESS = False`, e confira `output/resultado.xlsx` e um print. Os testes batem no site real, então use poucas perguntas e não rode em volume sem necessidade.
- Não há suíte de testes automatizados.

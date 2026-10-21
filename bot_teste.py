import os
import re
import shutil
import time
from pathlib import Path

from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from playwright.sync_api import expect, sync_playwright

# Dados locais (planilhas, .env) ficam em venv/, fora do git. Modelo em venv-example/.
PASTA_DADOS = Path("venv")
if not PASTA_DADOS.is_dir():
    raise SystemExit("Pasta venv/ não encontrada. Copie venv-example/ para venv/ e renomeie .env.example para .env")
load_dotenv(PASTA_DADOS / ".env")

# ---------- Configuração ----------
URL_DO_SITE = "https://brazino777.com/"
ENTRADA = str(PASTA_DADOS / "perguntas.xlsx")
PASTA_SAIDA = Path("output")  # tudo que o bot gera (planilha, prints, vídeos, traces); fora do git. Modelo em output-example/
SAIDA = str(PASTA_SAIDA / "resultado.xlsx")
PASTA_PRINTS = PASTA_SAIDA / "prints"
TIMEOUT_MS = 60_000   # limite total por pergunta (do envio até a resposta terminar) e por navegação
ESTAVEL_S = 3         # a resposta só é dada como pronta se não mudar durante este tempo

# determina se a aba do chrome fica aberta ou fechada (true - nao abre, false-abre)
HEADLESS = False
SLOW_MO_MS = 500       # pausa entre ações (ex.: 500) para acompanhar com HEADLESS = False
GRAVAR_VIDEO = False  # salva output/videos/linha_<N>/*.webm
GRAVAR_TRACE = False  # salva output/traces/linha_<N>.zip (abrir com: playwright show-trace)

# Seletores: ajuste com `playwright codegen URL_DO_SITE`
SEL_IFRAME = "iframe[title='janela de mensagens']"  # iframe do widget (Zendesk)
SEL_ABRIR_CHAT = ".chat-button"             # botão que abre o chat (fora do iframe); "" se já abre sozinho
SEL_INPUT = "#composer-input"               # campo de texto (dentro do iframe)
SEL_ENVIAR = "button[aria-label='Enviar mensagem']"  # botão enviar (dentro do iframe)
SEL_MSG_BOT = "div[shape][aria-live=polite]:not(:has-text('a escrever'))"  # bolhas do bot (as do usuário têm aria-live=off)
SEL_DIGITANDO = "div[shape][aria-live=polite]:has-text('a escrever')"      # "Mía está a escrever"; "" se não existir
SEL_JANELA = "iframe[title='janela de mensagens']"  # elemento da página que será fotografado

# O bot pede nome e confirmação antes de aceitar perguntas; enviadas a cada conversa, sem registrar.
MENSAGENS_INICIAIS = [
    m.strip().format(nome=os.getenv("NOME", ""))
    for m in os.getenv("MENSAGENS_INICIAIS", "").split("|")
    if m.strip()
]  # vêm do venv/.env (NOME e MENSAGENS_INICIAIS)
# Bolha final igual a um destes textos curtos = resposta ainda não chegou.
TEXTOS_INTERMEDIARIOS = ["um momento", "estou verificando"]
# ----------------------------------

POLL_S = 0.2
MAX_CELULA = 32_000  # limite do Excel é 32.767 caracteres


def ler_conversas(caminho):
    """Gera (linha, [(nº da pergunta, texto), ...]); o nº vem do cabeçalho 'Pergunta N'."""
    ws = load_workbook(caminho, read_only=True).active
    linhas = ws.iter_rows(values_only=True)
    cab = [str(c or "").strip() for c in next(linhas, [])]
    colunas = []
    for i, h in enumerate(cab):
        if h.lower().startswith("pergunta"):
            achado = re.search(r"\d+", h)
            colunas.append((i, int(achado.group()) if achado else len(colunas) + 1))
    if not colunas:
        raise SystemExit(f"Nenhuma coluna 'Pergunta N' encontrada em {caminho}")
    for n, linha in enumerate(linhas, start=2):
        perguntas = [
            (num, str(linha[i]).strip())
            for i, num in colunas
            if i < len(linha) and linha[i] is not None and str(linha[i]).strip()
        ]
        if perguntas:
            yield n, perguntas


def esperar_estavel(page, obter, timeout_s):
    """Espera `obter()` -> (estado, ocupado) ficar igual e sem 'ocupado' por ESTAVEL_S.
    Devolve (estado, instante da última mudança)."""
    fim = time.perf_counter() + timeout_s
    ultimo, mudou_em = object(), time.perf_counter()
    while time.perf_counter() < fim:
        estado, ocupado = obter()
        agora = time.perf_counter()
        if ocupado or estado != ultimo:
            ultimo, mudou_em = estado, agora
        elif agora - mudou_em >= ESTAVEL_S:
            return estado, mudou_em
        page.wait_for_timeout(POLL_S * 1000)
    raise TimeoutError(f"resposta não estabilizou em {timeout_s:.0f}s")


def textos_bot(msgs, a_partir_de=0):
    return [msgs.nth(i).inner_text() for i in range(a_partir_de, msgs.count())]


def limpar(texto):
    """Tira o rótulo 'Mía diz:' do início da bolha."""
    return re.sub(r"^[^\n]*\bdiz:\s*", "", texto).strip()


def intermediaria(texto):
    t = limpar(texto).lower()
    return len(t) < 80 and any(x in t for x in TEXTOS_INTERMEDIARIOS)


def trocar_mensagem(page, chat, msgs, texto):
    """Envia `texto` e devolve (resposta do bot, segundos até a última mudança)."""
    antes = msgs.count()
    chat.locator(SEL_INPUT).fill(texto)
    inicio = time.perf_counter()
    chat.locator(SEL_ENVIAR).click()
    expect(msgs.nth(antes)).to_be_visible(timeout=TIMEOUT_MS)

    def estado():
        novas = textos_bot(msgs, antes)
        return tuple(novas), digitando(chat) or intermediaria(novas[-1])

    restante = TIMEOUT_MS / 1000 - (time.perf_counter() - inicio)
    novas, ultima_mudanca = esperar_estavel(page, estado, max(restante, 1))
    return "\n".join(limpar(t) for t in novas), round(ultima_mudanca - inicio, 2)


def tirar_print(page, msgs, caminho):
    try:
        if msgs.count():
            msgs.last.scroll_into_view_if_needed(timeout=2000)
        page.locator(SEL_JANELA).screenshot(path=str(caminho), timeout=5000)
    except Exception as e:
        page.screenshot(path=str(caminho))
        print(f"  aviso: print da janela falhou ({str(e).splitlines()[0]}); salvei a página inteira")


def conversar(page, n_linha, perguntas, registrar):
    pasta = PASTA_PRINTS / f"linha_{n_linha}"
    shutil.rmtree(pasta, ignore_errors=True)
    pasta.mkdir(parents=True, exist_ok=True)

    chat = page.frame_locator(SEL_IFRAME)
    msgs = chat.locator(SEL_MSG_BOT)
    falha = None
    try:
        page.goto(URL_DO_SITE, wait_until="domcontentloaded")
        if SEL_ABRIR_CHAT:
            page.click(SEL_ABRIR_CHAT)
        expect(chat.locator(SEL_INPUT)).to_be_visible(timeout=TIMEOUT_MS)
        # deixa a saudação inicial (se houver) terminar antes de contar as mensagens
        esperar_estavel(page, lambda: (tuple(textos_bot(msgs)), digitando(chat)), TIMEOUT_MS / 1000)
        for texto in MENSAGENS_INICIAIS:
            trocar_mensagem(page, chat, msgs, texto)
    except Exception as e:
        falha = f"falha ao abrir o chat: {str(e).splitlines()[0]}"

    for num, pergunta in perguntas:
        caminho = pasta / f"pergunta_{num}.png"
        if falha:
            registrar(n_linha, num, pergunta, f"NÃO EXECUTADA ({falha})", "", "ERRO", "")
            continue
        resposta, status, tempo = "", "OK", ""
        try:
            resposta, tempo = trocar_mensagem(page, chat, msgs, pergunta)
        except Exception as e:
            status, resposta = "ERRO", str(e).splitlines()[0] if str(e) else type(e).__name__
            falha = f"pergunta {num} falhou"
        try:
            tirar_print(page, msgs, caminho)
            caminho_txt = str(caminho)
        except Exception as e:
            caminho_txt = f"(print falhou: {str(e).splitlines()[0]})"
        registrar(n_linha, num, pergunta, resposta, caminho_txt, status, tempo)


def digitando(chat):
    return bool(SEL_DIGITANDO) and chat.locator(SEL_DIGITANDO).first.is_visible()


def celula_segura(v):
    """Remove caracteres ilegais, limita o tamanho e impede que '=...' vire fórmula."""
    if isinstance(v, str):
        v = ILLEGAL_CHARACTERS_RE.sub("", v)[:MAX_CELULA]
    return v


def main():
    PASTA_SAIDA.mkdir(exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.append(["Linha", "Nº pergunta", "Pergunta", "Resposta do bot", "Print", "Status", "Tempo de resposta (s)"])

    def registrar(*valores):
        ws.append([celula_segura(v) for v in valores])
        for c in ws[ws.max_row]:
            if isinstance(c.value, str):
                c.data_type = "s"  # texto literal, nunca fórmula
        try:
            wb.save(SAIDA)  # salva a cada pergunta: nada se perde se o script cair
        except OSError as e:
            print(f"  aviso: não consegui salvar {SAIDA} ({e}); feche o arquivo se estiver aberto")
        print(f"linha {valores[0]} pergunta {valores[1]}: {valores[5]}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS, slow_mo=SLOW_MO_MS)
        for n_linha, perguntas in ler_conversas(ENTRADA):
            # sessão nova por linha
            context = browser.new_context(record_video_dir=str(PASTA_SAIDA / "videos" / f"linha_{n_linha}") if GRAVAR_VIDEO else None)
            context.set_default_timeout(TIMEOUT_MS)
            if GRAVAR_TRACE:
                context.tracing.start(screenshots=True, snapshots=True)
            try:
                conversar(context.new_page(), n_linha, perguntas, registrar)
            finally:
                if GRAVAR_TRACE:
                    context.tracing.stop(path=str(PASTA_SAIDA / "traces" / f"linha_{n_linha}.zip"))
                context.close()
        browser.close()
    print(f"Concluído: {ws.max_row - 1} registros em {SAIDA}")


if __name__ == "__main__":
    main()

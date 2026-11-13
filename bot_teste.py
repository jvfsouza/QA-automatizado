import argparse
import math
import multiprocessing
import os
import re
import shutil
import time
from pathlib import Path
from queue import Empty

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
URL_DO_SITE = "https://magalusellers.zendesk.com/hc/pt-br"
ENTRADA = str(PASTA_DADOS / "perguntas_transbordo_seller.xlsx")
PASTA_SAIDA = Path("output")  # tudo que o bot gera (planilha, prints, vídeos, traces); fora do git. Modelo em output-example/
SAIDA = str(PASTA_SAIDA / "resultado.xlsx")
PASTA_PRINTS = PASTA_SAIDA / "prints"
TIMEOUT_MS = 60_000   # limite total por pergunta (do envio até a resposta terminar) e por navegação
ESTAVEL_S = 3         # a resposta só é dada como pronta se não mudar durante este tempo
PARALELO = 1          # conversas (linhas) ao mesmo tempo; também via `python bot_teste.py --paralelo N`

# determina se a aba do chrome fica aberta ou fechada (true - nao abre, false-abre)
HEADLESS = False
TELA_LARGURA = 1920   # resolução da tela, usada para dividir as janelas em grade (HEADLESS = False com --paralelo)
TELA_ALTURA = 1080
SLOW_MO_MS = 500       # pausa entre ações (ex.: 500) para acompanhar com HEADLESS = False
GRAVAR_VIDEO = False  # salva output/videos/linha_<N>/*.webm
GRAVAR_TRACE = False  # salva output/traces/linha_<N>.zip (abrir com: playwright show-trace)

# Seletores: ajuste com `playwright codegen URL_DO_SITE`
SEL_IFRAME = "iframe[title='janela de mensagens' i]"  # iframe do widget (Zendesk)
SEL_IFRAME_ABRIR = "iframe#launcher"        # iframe do botão que abre o chat; "" se o botão está na própria página
SEL_ABRIR_CHAT = "button[aria-label*='Abrir' i]"  # botão que abre o chat (dentro de SEL_IFRAME_ABRIR); "" se já abre sozinho
SEL_INPUT = "#composer-input"               # campo de texto (dentro do iframe)
SEL_ENVIAR = "button[aria-label='Enviar mensagem']"  # botão enviar (dentro do iframe)
SEL_MSG_BOT = "div[shape][aria-live=polite]:not(:has-text('está digitando'))"  # bolhas do bot (as do usuário têm aria-live=off)
SEL_DIGITANDO = "div[shape][aria-live=polite]:has-text('está digitando')"  # "Bot está digitando"; "" se não existir
SEL_JANELA = "iframe[title='janela de mensagens' i]"  # elemento da página que será fotografado

# Login (só usado se o site redirecionar para a tela de entrada); credenciais vêm do venv/.env.
LOGIN_EMAIL = os.getenv("LOGIN_EMAIL", "")
LOGIN_SENHA = os.getenv("LOGIN_SENHA", "")
URL_LOGIN = "**/auth/**"                         # padrão da URL da tela de login
SEL_LOGIN_EMAIL = "[data-testid='email-input']"
SEL_LOGIN_SENHA = "[data-testid='password-input']"
SEL_LOGIN_ENVIAR = "[data-testid='submit-button']"

# O bot pede nome e confirmação antes de aceitar perguntas; enviadas a cada conversa, sem registrar.
MENSAGENS_INICIAIS = [
    m.strip().format(nome=os.getenv("NOME", ""))
    for m in os.getenv("MENSAGENS_INICIAIS", "").split("|")
    if m.strip()
]  # vêm do venv/.env (NOME e MENSAGENS_INICIAIS)
# Bolha final igual a um destes textos curtos = resposta ainda não chegou.
TEXTOS_INTERMEDIARIOS = ["um momento", "estou verificando", "aguarde"]
# Trechos (minúsculos) que indicam transbordo para atendente humano; ajuste conforme o que o bot do site diz.
# Só a transferência de fato: frases como "vou transferir" vêm ANTES da coleta de dados, e a conversa continua.
# Evite palavras soltas como "atendente": o bot as usa ao perguntar o motivo, antes de transferir.
TEXTOS_TRANSBORDO = [
    "transbordado",          # mensagem de transferência do ambiente de testes ("TU FOSTE TRANSBORDADO")
    "fila de atendimento",
]
# Respostas automáticas às perguntas de coleta do bot (ex.: "qual o número do pedido?"), para a conversa
# chegar até o fim. (trecho minúsculo procurado no fim da resposta, mensagem enviada); vale o primeiro que casar.
# São gravadas no resultado como "[auto] ..." com nº 3.1, 3.2... (depois da pergunta 3). Dados fictícios.
RESPOSTAS_AUTOMATICAS = [
    ('"sim" ou "não"', "Sim"),
    ("aceita a transferência", "Sim"),
    ("seguir por aqui", "Quero falar com um agente"),
    ("seguir comigo", "Quero falar com um agente"),
    ("número do pedido", "1234567890"),
    ("id do agendamento", "123456"),
    ("mês de apuração", "Outubro de 2026"),
    ("data e o horário", "Ontem, por volta das 14h"),
    ("relato detalhado", "O atendente recusou os pacotes dizendo que não constavam no sistema, mesmo com as etiquetas válidas."),
    ("foto", "Não tenho a imagem agora, podem seguir sem ela?"),
    ("print", "Não tenho a imagem agora, podem seguir sem ela?"),
]
MAX_AUTOMATICAS = 6  # limite de respostas automáticas seguidas, para não entrar em laço
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
    return len(t) < 160 and any(x in t for x in TEXTOS_INTERMEDIARIOS)  # ex.: "Vou verificar ... Aguarde um momento, por favor."


def transbordou(texto):
    t = texto.lower()
    return any(x in t for x in TEXTOS_TRANSBORDO)


def resposta_automatica(texto):
    """Mensagem para responder a uma pergunta de coleta no fim de `texto`, ou None."""
    fim = texto.lower()[-250:]
    return next((msg for trecho, msg in RESPOSTAS_AUTOMATICAS if trecho in fim), None)


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


def entrar(page):
    """Se o site mandou para a tela de login, entra com LOGIN_EMAIL/LOGIN_SENHA e volta para o site."""
    try:  # o desafio do Cloudflare pode atrasar o redirecionamento
        page.wait_for_url(URL_LOGIN, timeout=15_000)
    except Exception:
        return  # sem tela de login: o site já está acessível
    if not (LOGIN_EMAIL and LOGIN_SENHA):
        raise RuntimeError("o site exige login: preencha LOGIN_EMAIL e LOGIN_SENHA em venv/.env")
    page.fill(SEL_LOGIN_EMAIL, LOGIN_EMAIL)
    page.fill(SEL_LOGIN_SENHA, LOGIN_SENHA)
    page.click(SEL_LOGIN_ENVIAR)
    page.wait_for_url(lambda url: "/auth/" not in url, timeout=TIMEOUT_MS)


def conversar(page, n_linha, perguntas, registrar):
    pasta = PASTA_PRINTS / f"linha_{n_linha}"
    shutil.rmtree(pasta, ignore_errors=True)
    pasta.mkdir(parents=True, exist_ok=True)

    chat = page.frame_locator(SEL_IFRAME)
    msgs = chat.locator(SEL_MSG_BOT)
    falha = None
    try:
        page.goto(URL_DO_SITE, wait_until="domcontentloaded")
        entrar(page)
        if SEL_ABRIR_CHAT:
            botao = page.frame_locator(SEL_IFRAME_ABRIR) if SEL_IFRAME_ABRIR else page
            botao.locator(SEL_ABRIR_CHAT).first.click()
        expect(chat.locator(SEL_INPUT)).to_be_visible(timeout=TIMEOUT_MS)
        # deixa a saudação inicial (se houver) terminar antes de contar as mensagens
        esperar_estavel(page, lambda: (tuple(textos_bot(msgs)), digitando(chat)), TIMEOUT_MS / 1000)
        for texto in MENSAGENS_INICIAIS:
            trocar_mensagem(page, chat, msgs, texto)
    except Exception as e:
        falha = f"falha ao abrir o chat: {str(e).splitlines()[0]}"

    def passo(num, texto, rotulo=None):
        """Envia `texto`, tira o print e registra; devolve a resposta (ou None se falhou/transbordou)."""
        nonlocal falha
        caminho = pasta / f"pergunta_{num}.png"
        resposta, status, tempo = "", "OK", ""
        try:
            resposta, tempo = trocar_mensagem(page, chat, msgs, texto)
            if transbordou(resposta):
                status = "TRANSBORDO"
                falha = f"transbordo para atendente na pergunta {num}"
                print(f"  ⚠ AVISO: linha {n_linha}, pergunta {num}: o bot transbordou para atendente humano")
        except Exception as e:
            status, resposta = "ERRO", str(e).splitlines()[0] if str(e) else type(e).__name__
            falha = f"pergunta {num} falhou"
        try:
            tirar_print(page, msgs, caminho)
            caminho_txt = str(caminho)
        except Exception as e:
            caminho_txt = f"(print falhou: {str(e).splitlines()[0]})"
        registrar(n_linha, num, rotulo or texto, resposta, caminho_txt, status, tempo)
        return None if falha else resposta

    for num, pergunta in perguntas:
        if falha:
            registrar(n_linha, num, pergunta, f"NÃO EXECUTADA ({falha})", "", "ERRO", "")
            continue
        resposta = passo(num, pergunta)
        # responde sozinho às perguntas de coleta (nº do pedido, "Sim ou Não"...) até o bot seguir adiante
        for k in range(1, MAX_AUTOMATICAS + 1):
            auto = resposta and resposta_automatica(resposta)
            if not auto:
                break
            resposta = passo(round(num + k / 10, 1), auto, f"[auto] {auto}")


def digitando(chat):
    return bool(SEL_DIGITANDO) and chat.locator(SEL_DIGITANDO).first.is_visible()


def celula_segura(v):
    """Remove caracteres ilegais, limita o tamanho e impede que '=...' vire fórmula."""
    if isinstance(v, str):
        v = ILLEGAL_CHARACTERS_RE.sub("", v)[:MAX_CELULA]
    return v


CABECALHO = ["Linha", "Nº pergunta", "Pergunta", "Resposta do bot", "Print", "Status", "Tempo de resposta (s)"]


def posicao_janela(slot, total):
    """(x, y, largura, altura) da janela `slot` numa grade que divide a tela em `total` partes."""
    colunas = math.ceil(math.sqrt(total))
    linhas = math.ceil(total / colunas)
    largura, altura = TELA_LARGURA // colunas, TELA_ALTURA // linhas
    return (slot % colunas) * largura, (slot // colunas) * altura, largura, altura


def executar_linha(n_linha, perguntas, registrar, janela=None):
    """Roda uma conversa em um navegador próprio; nunca levanta exceção.
    `janela` = (x, y, largura, altura) posiciona o navegador na tela (só aparece com HEADLESS = False)."""
    feitas = set()

    def registrar_linha(*valores):
        feitas.add(valores[1])
        registrar(*valores)

    try:
        with sync_playwright() as p:
            args = [f"--window-position={janela[0]},{janela[1]}", f"--window-size={janela[2]},{janela[3]}"] if janela else []
            browser = p.chromium.launch(headless=HEADLESS, slow_mo=SLOW_MO_MS, args=args)
            try:
                # sessão nova por linha; com janela fixa, a página acompanha o tamanho da janela
                context = browser.new_context(
                    no_viewport=bool(janela),
                    record_video_dir=str(PASTA_SAIDA / "videos" / f"linha_{n_linha}") if GRAVAR_VIDEO else None,
                    record_video_size={"width": janela[2], "height": janela[3]} if janela and GRAVAR_VIDEO else None,
                )
                context.set_default_timeout(TIMEOUT_MS)
                if GRAVAR_TRACE:
                    context.tracing.start(screenshots=True, snapshots=True)
                try:
                    conversar(context.new_page(), n_linha, perguntas, registrar_linha)
                finally:
                    if GRAVAR_TRACE:
                        context.tracing.stop(path=str(PASTA_SAIDA / "traces" / f"linha_{n_linha}.zip"))
                    context.close()
            finally:
                browser.close()
    except Exception as e:  # ex.: navegador não abriu
        motivo = str(e).splitlines()[0] if str(e) else type(e).__name__
        for num, pergunta in perguntas:
            if num not in feitas:
                registrar(n_linha, num, pergunta, f"NÃO EXECUTADA (erro inesperado: {motivo})", "", "ERRO", "")


def trabalhador(fila, vagas, n_vagas, n_linha, perguntas):
    """Roda em um processo filho: manda cada resultado para o processo principal, que é o único a gravar o Excel.
    Pega uma vaga da grade de janelas enquanto roda, para duas conversas simultâneas não ficarem no mesmo lugar."""
    vaga = vagas.get()
    try:
        janela = None if HEADLESS else posicao_janela(vaga, n_vagas)
        executar_linha(n_linha, perguntas, lambda *valores: fila.put(valores), janela)
    finally:
        vagas.put(vaga)


def salvar(resultados):
    wb = Workbook()
    ws = wb.active
    ws.append(CABECALHO)
    for r in sorted(resultados, key=lambda r: (r[0], r[1])):
        ws.append(r)
    for linha in ws.iter_rows(min_row=2):
        for c in linha:
            if isinstance(c.value, str):
                c.data_type = "s"  # texto literal, nunca fórmula
    try:
        wb.save(SAIDA)  # salva a cada pergunta: nada se perde se o script cair
    except OSError as e:
        print(f"  aviso: não consegui salvar {SAIDA} ({e}); feche o arquivo se estiver aberto")


def main(paralelo=1):
    PASTA_SAIDA.mkdir(exist_ok=True)
    linhas = list(ler_conversas(ENTRADA))
    resultados = []

    def registrar(*valores):
        resultados.append([celula_segura(v) for v in valores])
        salvar(resultados)
        print(f"linha {valores[0]} pergunta {valores[1]}: {valores[5]}")

    if paralelo <= 1 or len(linhas) <= 1:
        for n_linha, perguntas in linhas:
            executar_linha(n_linha, perguntas, registrar)
    else:
        ctx = multiprocessing.get_context("spawn")
        n_vagas = min(paralelo, len(linhas))
        with ctx.Manager() as manager, ctx.Pool(n_vagas) as pool:
            fila, vagas = manager.Queue(), manager.Queue()
            for v in range(n_vagas):
                vagas.put(v)
            tarefas = [pool.apply_async(trabalhador, (fila, vagas, n_vagas, n, perg)) for n, perg in linhas]
            while True:
                try:
                    registrar(*fila.get(timeout=0.5))
                except Empty:
                    if all(t.ready() for t in tarefas):
                        break
            while not fila.empty():
                registrar(*fila.get())
            for t in tarefas:
                t.get()  # propaga erro de processo filho, se houver
    print(f"Concluído: {len(resultados)} registros em {SAIDA}")
    transbordos = sorted({r[0] for r in resultados if r[5] == "TRANSBORDO"})
    if transbordos:
        print(f"⚠ AVISO: transbordo para atendente humano nas linhas: {', '.join(map(str, transbordos))}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Bot de testes do chat de atendimento")
    parser.add_argument("--paralelo", type=int, default=PARALELO, metavar="N",
                        help="quantas conversas (linhas da planilha) rodar ao mesmo tempo (padrão: %(default)s)")
    main(parser.parse_args().paralelo)

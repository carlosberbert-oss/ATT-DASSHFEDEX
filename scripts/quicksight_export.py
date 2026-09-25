#!/usr/bin/env python3
"""
Exporta o CSV do visual "Details In Transit" do QuickSight e grava
na aba "Base de dados" da planilha Dash_fedex.

Fluxo automatizado (o mesmo que era feito na mão):
  1. Login no QuickSight (conta + usuário + senha, sem MFA)
  2. Abre o dashboard Logistic KPI's
  3. Vai na aba "3PL In Transit"
  4. Clica no menu do visual "Details In Transit" e exporta CSV
  5. Espera o arquivo ficar pronto e baixa
  6. Escreve o conteúdo na aba "Base de dados"

Variáveis de ambiente necessárias:
  QS_ACCOUNT        nome da conta QuickSight (ex: advanceanalytics)
  QS_USER           usuário/email
  QS_PASSWORD       senha
  GOOGLE_CREDS_JSON conteúdo do JSON da service account do Google
  SHEET_ID          id da planilha (opcional, tem default)

Rodar local:
  pip install -r requirements.txt
  playwright install chromium
  python scripts/quicksight_export.py --headed     # pra ver o navegador
"""

import argparse
import csv
import io
import json
import os
import sys
import time
from pathlib import Path

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ── Configuração ────────────────────────────────────────────────

QS_REGION = "us-east-1"
DASHBOARD_ID = "e14ad331-ba9b-4d41-82a5-b05f9075da13"   # Logistic KPI's
ABA_DASHBOARD = "3PL In Transit"
NOME_VISUAL = "Details In Transit"

SHEET_ID_PADRAO = "1RMM3Niaf3E9WVCN3NdfU1Cb4Fuz4VSJjaLUPBdrFldI"
ABA_DESTINO = "Base de dados"

# O export é assíncrono: o QuickSight prepara o arquivo e só depois
# dispara o download. Em bases grandes isso leva um tempo.
TIMEOUT_DOWNLOAD_MS = 300_000   # 5 min
TIMEOUT_PADRAO_MS = 60_000

DIR_SAIDA = Path("saida")
DIR_DEBUG = Path("debug")

# Textos de menu aceitos — o runner pode renderizar em outro idioma
# que não o português, então cobrimos as três variações.
# Confirmado no dashboard: o menu tem "Exportar para CSV" e
# "Exportar para o Excel". As outras variações cobrem o caso do
# runner renderizar a interface em outro idioma.
TEXTOS_EXPORTAR = [
    "Exportar para CSV",
    "Exportar em CSV", "Baixar CSV",
    "Export to CSV", "Download CSV",
    "Exportar a CSV", "Descargar CSV",
]


def log(msg):
    print(f"[{time.strftime('%H:%M:%S')}] {msg}", flush=True)


def salvar_debug(page, nome):
    """Screenshot + HTML quando algo dá errado — essencial pra
    descobrir o que mudou quando a AWS mexe no layout."""
    DIR_DEBUG.mkdir(exist_ok=True)
    try:
        page.screenshot(path=str(DIR_DEBUG / f"{nome}.png"), full_page=True)
        (DIR_DEBUG / f"{nome}.html").write_text(page.content(), encoding="utf-8")
        log(f"Debug salvo em debug/{nome}.png e .html")
    except Exception as e:
        log(f"Não consegui salvar debug: {e}")


# ── Etapa 1: login ──────────────────────────────────────────────

def fazer_login(page, conta, usuario, senha):
    url = f"https://{QS_REGION}.quicksight.aws.amazon.com/sn/auth/signin"
    log(f"Abrindo {url}")
    page.goto(url, wait_until="domcontentloaded", timeout=TIMEOUT_PADRAO_MS)

    # Passo 1 — nome da conta
    log("Preenchendo nome da conta")
    campo_conta = page.locator("input").first
    campo_conta.wait_for(state="visible", timeout=TIMEOUT_PADRAO_MS)
    campo_conta.fill(conta)
    page.keyboard.press("Enter")

    # Passo 2 — usuário
    log("Preenchendo usuário")
    page.wait_for_timeout(2500)
    campo_user = page.locator("input[type='text'], input[type='email']").first
    campo_user.wait_for(state="visible", timeout=TIMEOUT_PADRAO_MS)
    campo_user.fill(usuario)
    page.keyboard.press("Enter")

    # Passo 3 — senha (redireciona pra signin.aws)
    log("Aguardando tela de senha")
    campo_senha = page.locator("input[type='password']").first
    campo_senha.wait_for(state="visible", timeout=TIMEOUT_PADRAO_MS)
    campo_senha.fill(senha)
    page.keyboard.press("Enter")

    # Confirma que entrou
    log("Aguardando entrar no QuickSight")
    page.wait_for_url("**/sn/account/**", timeout=TIMEOUT_PADRAO_MS)
    page.wait_for_load_state("networkidle", timeout=TIMEOUT_PADRAO_MS)
    log("Login concluído")


# ── Etapa 2: abrir o dashboard e a aba certa ────────────────────

def fechar_popups(page):
    """O QuickSight abre um modal de novidades ("Estamos ocupados criando")
    por cima do dashboard. Ele intercepta todos os cliques, então precisa
    sair da frente antes de qualquer coisa."""
    page.wait_for_timeout(1500)

    tentativas = [
        ("botão Entrar",    lambda: page.get_by_role("button", name="Entrar").first),
        ("botão Enter",     lambda: page.get_by_role("button", name="Enter").first),
        ("botão Entrar/ES", lambda: page.get_by_role("button", name="Entrar").first),
        ("X de fechar",     lambda: page.locator("[aria-label='Fechar'], [aria-label='Close']").first),
    ]

    for nome, obter in tentativas:
        try:
            el = obter()
            if el.count() > 0 and el.is_visible(timeout=2000):
                log(f"Fechando popup pelo {nome}")
                el.click(timeout=5000)
                page.wait_for_timeout(1500)
                return True
        except Exception:
            continue

    # Último recurso: Esc costuma fechar modais
    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(1000)
    except Exception:
        pass

    return False


def garantir_aba_correta(page):
    """O link do dashboard já abre na aba 3PL In Transit. Só clica se,
    por algum motivo, outra aba estiver selecionada — clicar numa aba
    já ativa dá erro de elemento interceptado."""
    selecionada = page.locator("[data-automation-id='selectedTab_sheet_name']")

    if selecionada.count() > 0:
        contexto = selecionada.first.get_attribute("data-automation-context") or ""
        if ABA_DASHBOARD.lower() in contexto.lower():
            log(f"Aba '{ABA_DASHBOARD}' já está selecionada")
            return

    log(f"Clicando na aba '{ABA_DASHBOARD}'")
    aba = page.get_by_text(ABA_DASHBOARD, exact=True).first
    aba.wait_for(state="visible", timeout=TIMEOUT_PADRAO_MS)
    aba.click()
    page.wait_for_timeout(5000)


def abrir_dashboard(page, conta):
    url = (f"https://{QS_REGION}.quicksight.aws.amazon.com"
           f"/sn/account/{conta}/dashboards/{DASHBOARD_ID}")
    log("Abrindo dashboard")
    page.goto(url, wait_until="domcontentloaded", timeout=TIMEOUT_PADRAO_MS)

    # O dashboard demora a montar; espera antes de procurar o popup
    page.wait_for_timeout(8000)
    fechar_popups(page)

    garantir_aba_correta(page)

    # Os visuais carregam de forma assíncrona
    page.wait_for_timeout(5000)
    try:
        page.wait_for_load_state("networkidle", timeout=30_000)
    except PWTimeout:
        log("networkidle não chegou, seguindo mesmo assim")

    fechar_popups(page)   # às vezes aparece de novo depois do carregamento
    log("Dashboard carregado")


# ── Etapa 3: exportar o CSV ─────────────────────────────────────

def localizar_visual(page):
    """Acha o visual 'Details In Transit'.

    O título é renderizado como "Details" + "In Transit" em negrito, ou
    seja, dois elementos separados no HTML — procurar pela frase inteira
    não casa. Por isso a busca é por "Details", que é único nessa tela
    (a outra tabela se chama "In Transit Con Internal Cases").
    """
    log("Procurando o visual 'Details In Transit' (rolando a página)")

    for tentativa in range(20):
        candidatos = page.get_by_text("Details", exact=False)

        for i in range(min(candidatos.count(), 5)):
            el = candidatos.nth(i)
            try:
                if not el.is_visible(timeout=1000):
                    continue
                texto = (el.inner_text(timeout=2000) or "").strip()
                # Confirma que é o título certo e não outro "Details"
                if "details" in texto.lower() and len(texto) < 60:
                    el.scroll_into_view_if_needed(timeout=5000)
                    page.wait_for_timeout(2000)
                    log(f"Visual encontrado: '{texto}' (após {tentativa} rolagem(ns))")
                    return el
            except Exception:
                continue

        page.mouse.wheel(0, 700)
        page.wait_for_timeout(1000)

    salvar_debug(page, "visual_nao_encontrado")
    raise RuntimeError(
        "Não achei o visual 'Details In Transit' mesmo rolando a página — veja debug/"
    )


def _menu_de_exportacao_apareceu(page):
    """Confirma que o menu aberto é mesmo o de exportação."""
    for texto in TEXTOS_EXPORTAR:
        if page.get_by_text(texto, exact=False).count() > 0:
            return True
    # "Exportar para o Excel" fica ao lado do CSV no mesmo menu
    return page.get_by_text("Exportar para o Excel", exact=False).count() > 0


def abrir_menu_visual(page, titulo):
    """Abre o menu (⋮) do visual.

    A barra de ícones só aparece no hover e tem vários botões (favoritar,
    maximizar, ordenar, ⋮). Em vez de apostar em qual é o certo, tentamos
    os candidatos da direita para a esquerda e verificamos, a cada clique,
    se o menu que abriu contém a opção de exportar. Se não for o menu
    certo, fecha com Esc e tenta o próximo.
    """
    caixa = titulo.bounding_box()
    if not caixa:
        return False

    titulo.hover()
    page.wait_for_timeout(1000)
    page.mouse.move(caixa["x"] + 900, caixa["y"] - 15)
    page.wait_for_timeout(1200)

    y_titulo = caixa["y"]

    # Coleta os botões visíveis na faixa horizontal do título
    candidatos = []
    botoes = page.locator("button, [role='button']")
    for i in range(botoes.count()):
        b = botoes.nth(i)
        try:
            if not b.is_visible(timeout=200):
                continue
            cb = b.bounding_box()
            if not cb:
                continue
            if abs(cb["y"] - y_titulo) > 60:
                continue
            # Ignora elementos largos demais pra ser um ícone de barra
            if cb["width"] > 80:
                continue
            candidatos.append((cb["x"], b))
        except Exception:
            continue

    if not candidatos:
        salvar_debug(page, "barra_icones_nao_apareceu")
        return False

    # Da direita para a esquerda — o ⋮ costuma ser o último
    candidatos.sort(key=lambda t: t[0], reverse=True)
    log(f"{len(candidatos)} ícone(s) na barra do visual; testando um a um")

    for pos, (x, botao) in enumerate(candidatos[:6]):
        try:
            titulo.hover()
            page.wait_for_timeout(400)
            botao.click(timeout=4000)
            page.wait_for_timeout(1200)

            if _menu_de_exportacao_apareceu(page):
                log(f"Menu de exportação aberto (ícone em x={x:.0f})")
                return True

            # Menu errado — fecha e tenta o próximo
            log(f"Ícone em x={x:.0f} não abriu o menu de exportação")
            page.keyboard.press("Escape")
            page.wait_for_timeout(600)

        except Exception as e:
            log(f"Ícone em x={x:.0f} falhou: {e}")
            try:
                page.keyboard.press("Escape")
                page.wait_for_timeout(400)
            except Exception:
                pass

    salvar_debug(page, "nenhum_icone_abriu_menu")
    return False


def clicar_exportar_csv(page):
    for texto in TEXTOS_EXPORTAR:
        item = page.get_by_text(texto, exact=False)
        if item.count() > 0:
            log(f"Clicando em '{texto}'")
            item.first.click()
            return True
    return False


def exportar_csv(page):
    fechar_popups(page)
    titulo = localizar_visual(page)

    if not abrir_menu_visual(page, titulo):
        salvar_debug(page, "menu_visual_nao_encontrado")
        raise RuntimeError(
            "Não consegui abrir o menu do visual. "
            "O layout do QuickSight provavelmente mudou — veja debug/"
        )

    log("Menu aberto — clicando em exportar e aguardando o arquivo")
    with page.expect_download(timeout=TIMEOUT_DOWNLOAD_MS) as info:
        if not clicar_exportar_csv(page):
            salvar_debug(page, "opcao_exportar_nao_encontrada")
            raise RuntimeError(
                "Menu abriu mas não achei a opção de exportar CSV — veja debug/"
            )
        log("Clique feito. O QuickSight prepara o arquivo antes de baixar "
            "(pode levar alguns minutos)")

    download = info.value
    DIR_SAIDA.mkdir(exist_ok=True)
    destino = DIR_SAIDA / "details_in_transit.csv"
    download.save_as(str(destino))

    tamanho_kb = destino.stat().st_size / 1024
    log(f"CSV baixado: {destino} ({tamanho_kb:.0f} KB)")
    return destino


# ── Etapa 4: escrever no Google Sheets ──────────────────────────

def escrever_no_sheets(caminho_csv, sheet_id, creds_json):
    import gspread
    from google.oauth2.service_account import Credentials

    log("Conectando no Google Sheets")
    escopos = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(
        json.loads(creds_json), scopes=escopos
    )
    gc = gspread.authorize(creds)

    planilha = gc.open_by_key(sheet_id)
    try:
        aba = planilha.worksheet(ABA_DESTINO)
    except gspread.WorksheetNotFound:
        raise RuntimeError(f'Aba "{ABA_DESTINO}" não existe na planilha')

    # Lê o CSV. O QuickSight exporta em UTF-8 com BOM.
    with open(caminho_csv, "r", encoding="utf-8-sig", newline="") as f:
        linhas = list(csv.reader(f))

    if not linhas:
        raise RuntimeError("CSV veio vazio — abortando pra não apagar a aba")

    log(f"CSV tem {len(linhas)} linha(s) e {len(linhas[0])} coluna(s)")

    # Só limpa depois de ter os dados em mãos — se algo falhar antes,
    # a aba continua com o conteúdo anterior em vez de ficar vazia.
    log("Limpando a aba e gravando os dados novos")
    aba.clear()
    aba.update(values=linhas, range_name="A1")

    log(f'Aba "{ABA_DESTINO}" atualizada com sucesso')
    return len(linhas)


# ── Principal ───────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--headed", action="store_true",
                        help="mostra o navegador (útil pra depurar)")
    parser.add_argument("--so-baixar", action="store_true",
                        help="baixa o CSV mas não escreve no Sheets")
    args = parser.parse_args()

    conta = os.environ.get("QS_ACCOUNT", "advanceanalytics")
    usuario = os.environ.get("QS_USER")
    senha = os.environ.get("QS_PASSWORD")
    sheet_id = os.environ.get("SHEET_ID", SHEET_ID_PADRAO)
    creds_json = os.environ.get("GOOGLE_CREDS_JSON")

    if not usuario or not senha:
        sys.exit("Faltam as variáveis QS_USER e QS_PASSWORD")
    if not args.so_baixar and not creds_json:
        sys.exit("Falta a variável GOOGLE_CREDS_JSON")

    with sync_playwright() as p:
        navegador = p.chromium.launch(
            headless=not args.headed,
            args=["--disable-blink-features=AutomationControlled"],
        )
        contexto = navegador.new_context(
            viewport={"width": 1920, "height": 1080},
            accept_downloads=True,
            locale="pt-BR",
        )
        page = contexto.new_page()

        try:
            fazer_login(page, conta, usuario, senha)
            abrir_dashboard(page, conta)
            caminho = exportar_csv(page)

            if args.so_baixar:
                log("Modo --so-baixar: não escrevi no Sheets")
            else:
                total = escrever_no_sheets(caminho, sheet_id, creds_json)
                log(f"Concluído — {total} linha(s) na planilha")

        except PWTimeout as e:
            salvar_debug(page, "timeout")
            sys.exit(f"Timeout: {e}")
        except Exception as e:
            salvar_debug(page, "erro")
            sys.exit(f"Erro: {e}")
        finally:
            contexto.close()
            navegador.close()


if __name__ == "__main__":
    main()

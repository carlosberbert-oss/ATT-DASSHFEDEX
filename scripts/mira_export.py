#!/usr/bin/env python3
"""
Baixa o relatório de rastreamento da Mira e grava numa aba da Dash_fedex.

Fluxo automatizado (o mesmo que era feito na mão):
  1. Login em cliente.mira.com.br (e-mail + senha, sem MFA)
  2. Abre /rastrear/remetente com o período na própria URL
  3. Menu ⋯ → "Baixar CSV (1 NF por linha)"
  4. Escreve o conteúdo na aba MiraRelatorio

O site limita o período a 30 dias — o script respeita isso.

Variáveis de ambiente:
  MIRA_USER         e-mail de login
  MIRA_PASSWORD     senha
  GOOGLE_CREDS_JSON conteúdo do JSON da service account
  SHEET_ID          id da planilha (opcional, tem default)
  MIRA_DIAS         quantos dias buscar (opcional, padrão 30)

Rodar local:
  python scripts/mira_export.py --headed --so-baixar
"""

import argparse
import csv
import json
import os
import sys
import time
from datetime import date, timedelta
from pathlib import Path

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ── Configuração ────────────────────────────────────────────────

MIRA_BASE = "https://cliente.mira.com.br"
SHEET_ID_PADRAO = "1RMM3Niaf3E9WVCN3NdfU1Cb4Fuz4VSJjaLUPBdrFldI"
ABA_DESTINO = "MiraRelatorio"

DIAS_PADRAO = 30          # o site recusa períodos maiores
TIMEOUT_PADRAO_MS = 60_000
TIMEOUT_DOWNLOAD_MS = 180_000

DIR_SAIDA = Path("saida")
DIR_DEBUG = Path("debug")

# A opção "1 NF por linha" é a que queremos: sem ela, várias notas
# ficam agrupadas numa linha só e a conciliação por NF não funciona.
TEXTOS_BAIXAR = [
    "Baixar CSV (1 NF por linha)",
    "Baixar CSV (1 NF por linha)".replace("(", "").replace(")", ""),
    "1 NF por linha",
]
TEXTO_BAIXAR_FALLBACK = "Baixar CSV"


def log(msg):
    print(f"[{time.strftime('%H:%M:%S')}] {msg}", flush=True)


def salvar_debug(page, nome):
    DIR_DEBUG.mkdir(exist_ok=True)
    try:
        page.screenshot(path=str(DIR_DEBUG / f"mira_{nome}.png"), full_page=True)
        (DIR_DEBUG / f"mira_{nome}.html").write_text(page.content(), encoding="utf-8")
        log(f"Debug salvo em debug/mira_{nome}.png e .html")
    except Exception as e:
        log(f"Não consegui salvar debug: {e}")


# ── Etapa 1: login ──────────────────────────────────────────────

def fazer_login(page, usuario, senha):
    log(f"Abrindo {MIRA_BASE}/login")
    page.goto(f"{MIRA_BASE}/login", wait_until="domcontentloaded",
              timeout=TIMEOUT_PADRAO_MS)

    log("Preenchendo e-mail")
    campo_email = page.locator("input[type='email'], input[type='text']").first
    campo_email.wait_for(state="visible", timeout=TIMEOUT_PADRAO_MS)
    campo_email.fill(usuario)

    log("Preenchendo senha")
    campo_senha = page.locator("input[type='password']").first
    campo_senha.wait_for(state="visible", timeout=TIMEOUT_PADRAO_MS)
    campo_senha.fill(senha)

    page.get_by_role("button", name="Entrar").first.click()

    log("Aguardando entrar")
    page.wait_for_url("**/dashboard**", timeout=TIMEOUT_PADRAO_MS)
    page.wait_for_load_state("networkidle", timeout=TIMEOUT_PADRAO_MS)
    log("Login concluído")


# ── Etapa 2: abrir o relatório no período ───────────────────────

def abrir_relatorio(page, dias):
    fim = date.today()
    inicio = fim - timedelta(days=dias - 1)

    # O período vai direto na URL — não precisa mexer no seletor de datas
    url = (f"{MIRA_BASE}/rastrear/remetente/"
           f"{inicio.isoformat()}/{fim.isoformat()}?group=false")

    log(f"Abrindo relatório de {inicio.strftime('%d/%m/%Y')} "
        f"a {fim.strftime('%d/%m/%Y')}")
    page.goto(url, wait_until="domcontentloaded", timeout=TIMEOUT_PADRAO_MS)

    # A tabela carrega de forma assíncrona ("Carregando rastreamento...")
    page.wait_for_timeout(6000)
    try:
        page.wait_for_load_state("networkidle", timeout=60_000)
    except PWTimeout:
        log("networkidle não chegou, seguindo mesmo assim")

    # Espera o carregamento sumir
    for _ in range(20):
        if page.get_by_text("Carregando rastreamento", exact=False).count() == 0:
            break
        page.wait_for_timeout(2000)

    page.wait_for_timeout(2000)
    log("Relatório carregado")


# ── Etapa 3: baixar o CSV ───────────────────────────────────────

def abrir_menu_opcoes(page):
    """O menu fica no ⋯ do canto superior direito, ao lado do
    'Guia do rastreamento'."""
    candidatos = [
        lambda: page.get_by_role("button", name="⋯").first,
        lambda: page.locator("[aria-label*='opç' i], [aria-label*='option' i]").first,
        lambda: page.locator("button:has-text('⋯')").first,
    ]

    for obter in candidatos:
        try:
            el = obter()
            if el.count() > 0 and el.is_visible(timeout=2000):
                el.click(timeout=5000)
                page.wait_for_timeout(1200)
                if page.get_by_text("Baixar CSV", exact=False).count() > 0:
                    log("Menu de opções aberto")
                    return True
        except Exception:
            continue

    # Alternativa geométrica: o ⋯ é o botão mais à direita do cabeçalho
    log("Tentando achar o menu pela posição no cabeçalho")
    botoes = page.locator("button, [role='button']")
    melhor, melhor_x = None, -1

    for i in range(botoes.count()):
        b = botoes.nth(i)
        try:
            if not b.is_visible(timeout=200):
                continue
            cb = b.bounding_box()
            if not cb or cb["y"] > 140:      # só o topo da página
                continue
            if cb["width"] > 80:             # ignora botões largos
                continue
            if cb["x"] > melhor_x:
                melhor_x, melhor = cb["x"], b
        except Exception:
            continue

    if melhor is not None:
        try:
            melhor.click(timeout=5000)
            page.wait_for_timeout(1200)
            if page.get_by_text("Baixar CSV", exact=False).count() > 0:
                log(f"Menu aberto pelo botão em x={melhor_x:.0f}")
                return True
        except Exception:
            pass

    salvar_debug(page, "menu_opcoes_nao_abriu")
    return False


def clicar_baixar(page):
    for texto in TEXTOS_BAIXAR:
        item = page.get_by_text(texto, exact=False)
        if item.count() > 0:
            log(f"Clicando em '{texto}'")
            item.first.click()
            return True

    # Se a opção "1 NF por linha" não existir, usa a comum
    item = page.get_by_text(TEXTO_BAIXAR_FALLBACK, exact=False)
    if item.count() > 0:
        log("Opção '1 NF por linha' não encontrada — usando 'Baixar CSV'")
        item.first.click()
        return True

    return False


def baixar_csv(page):
    if not abrir_menu_opcoes(page):
        raise RuntimeError("Não consegui abrir o menu de opções — veja debug/")

    log("Clicando em baixar e aguardando o arquivo")
    with page.expect_download(timeout=TIMEOUT_DOWNLOAD_MS) as info:
        if not clicar_baixar(page):
            salvar_debug(page, "opcao_baixar_nao_encontrada")
            raise RuntimeError("Menu abriu mas não achei a opção de baixar — veja debug/")

    download = info.value
    DIR_SAIDA.mkdir(exist_ok=True)
    destino = DIR_SAIDA / "mira_relatorio.csv"
    download.save_as(str(destino))

    log(f"CSV baixado: {destino} ({destino.stat().st_size / 1024:.0f} KB)")
    return destino


# ── Etapa 4: escrever no Google Sheets ──────────────────────────

def escrever_no_sheets(caminho_csv, sheet_id, creds_json):
    import gspread
    from google.oauth2.service_account import Credentials

    log("Conectando no Google Sheets")
    creds = Credentials.from_service_account_info(
        json.loads(creds_json),
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    gc = gspread.authorize(creds)
    planilha = gc.open_by_key(sheet_id)

    try:
        aba = planilha.worksheet(ABA_DESTINO)
    except gspread.WorksheetNotFound:
        log(f'Aba "{ABA_DESTINO}" não existe — criando')
        aba = planilha.add_worksheet(title=ABA_DESTINO, rows=1000, cols=30)

    # O CSV da Mira pode vir em UTF-8 com BOM e separado por ; ou ,
    with open(caminho_csv, "r", encoding="utf-8-sig", newline="") as f:
        amostra = f.read(4096)
        f.seek(0)
        sep = ";" if amostra.count(";") > amostra.count(",") else ","
        linhas = list(csv.reader(f, delimiter=sep))

    if not linhas:
        raise RuntimeError("CSV veio vazio — abortando pra não apagar a aba")

    log(f"CSV tem {len(linhas)} linha(s) e {len(linhas[0])} coluna(s) (separador '{sep}')")
    log(f"Cabeçalho: {linhas[0]}")

    aba.clear()
    aba.update(values=linhas, range_name="A1")
    log(f'Aba "{ABA_DESTINO}" atualizada')
    return len(linhas)


# ── Principal ───────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--headed", action="store_true")
    parser.add_argument("--so-baixar", action="store_true")
    parser.add_argument("--dias", type=int, default=None,
                        help=f"quantos dias buscar (padrão {DIAS_PADRAO}, máximo 30)")
    args = parser.parse_args()

    usuario = os.environ.get("MIRA_USER")
    senha = os.environ.get("MIRA_PASSWORD")
    sheet_id = os.environ.get("SHEET_ID", SHEET_ID_PADRAO)
    creds_json = os.environ.get("GOOGLE_CREDS_JSON")

    dias = args.dias or int(os.environ.get("MIRA_DIAS", DIAS_PADRAO))
    if dias > 30:
        log("A Mira limita o período a 30 dias — ajustando")
        dias = 30

    if not usuario or not senha:
        sys.exit("Faltam as variáveis MIRA_USER e MIRA_PASSWORD")
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
            fazer_login(page, usuario, senha)
            abrir_relatorio(page, dias)
            caminho = baixar_csv(page)

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

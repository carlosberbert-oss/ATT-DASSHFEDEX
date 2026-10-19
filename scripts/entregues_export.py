#!/usr/bin/env python3
"""
PEDIDOS ENTREGUES — atualização automática da aba BASE.

Roda todo dia às 12h (disparo externo pelo cron-job.org, igual ao bot do
QuickSight) e faz o que hoje é feito na mão:

  1. QuickSight → dashboard "Cadena de Suministro" → aba "On Time":
     coloca a "Fecha de Entrega" = ONTEM nos dois campos e exporta a tabela
     RAW DATA ("Total Volume to be delivered by week and % On Time").
     Cada linha do export é UM ITEM (o DOCNAME se repete quando o pedido
     tem mais de um item). O QTY é ignorado.
  2. Zecore (Shipping Core, um por pedido): data de criação e, para cada
     item (SCI), a NF e o endereço completo.
  3. Rastreio da transportadora pela NF: Jamef pela API, Fitlog e Mira pelo
     SSW → data de entrega da transportadora e previsão. Inhouse: em branco.
  4. Calcula transportadora, On time / Late time, UF, Região, SEMANA, MÊS e
     DIAS.
  5. Acrescenta na BASE só os SCIs que ainda não estão lá (trava contra
     duplicidade). "Baixa API" e "Motivo do atraso" ficam em branco — são
     preenchidos à mão.

Variáveis de ambiente:
  QS_ACCOUNT, QS_USER, QS_PASSWORD      QuickSight (as mesmas do outro bot)
  GOOGLE_CREDS_JSON                      service account do Google
  ZECORE_USER, ZECORE_PASSWORD           login do Zecore (o mesmo do bot de XML)
  JAMEF_API_USER, JAMEF_API_PASS         API da Jamef (as do Apps Script)
  ENTREGUES_SHEET_ID                     opcional (tem default)

Uso:
  python scripts/entregues_export.py                    # dia anterior, completo
  python scripts/entregues_export.py --data 2026-10-13  # reprocessa um dia
  python scripts/entregues_export.py --csv arquivo.csv  # pula o QuickSight
  python scripts/entregues_export.py --simular          # não grava na planilha
  python scripts/entregues_export.py --headed           # mostra o navegador
"""

import argparse
import csv
import datetime as dt
import json
import os
import re
import sys
import time
import unicodedata
from pathlib import Path

import requests

# ── Configuração ────────────────────────────────────────────────

QS_REGION = "us-east-1"
DASHBOARD_ID = "2f4c72a5-0568-4f16-8b27-86dd73eb9a0b"     # Cadena de Suministro
ABA_DASHBOARD = "On Time"
TITULO_VISUAL = "Total Volume to be delivered"            # tabela da seção RAW DATA
CONTROLE_DATA = "Fecha de Entrega"

SHEET_ID_PADRAO = "1GXxR37E_lCvJXDVE3nDP-IxMiXNxv5XmMorXe8vMaYk"   # PEDIDOS ENTREGUES
ABA_BASE = "BASE"

ZECORE_URL = "https://zecore.zebrands.mx"
DASHFEDEX_ID_PADRAO = "1RMM3Niaf3E9WVCN3NdfU1Cb4Fuz4VSJjaLUPBdrFldI"   # onde ficam os relatórios Jamef e Mira
ABA_JAMEF_REL = "JamefRelatorio"     # relatório do Portal do Cliente (mesclado: guarda o histórico)
ABA_MIRA_REL = "MiraRelatorio"       # relatório da Mira (CSV oficial)
URL_SSW_SERVICO = "https://ssw.inf.br/2/servico"   # download do XML do CT-e (Fitlog)
COL_VALOR_FRETE, COL_VALOR_NF = "VALOR FRETE", "VALOR NF"
COL_LAT, COL_LNG = "LAT", "LNG"          # coordenadas do endereço (para os alfinetes do mapa)
COLUNAS_NOVAS = (COL_VALOR_FRETE, COL_VALOR_NF, COL_LAT, COL_LNG)
NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
NOMINATIM_UA = "PainelLogisticaLuuna/1.0 (logistica@zeb.mx)"   # o Nominatim exige identificar quem consulta
CNPJ = "42418313000104"
JAMEF_API = "https://api.jamef.com.br"
SSW_URL = "https://ssw.inf.br/2/resultSSW"
SSW_SENHAS = {"FITLOG": "0104", "Mira": ""}

# Transportadoras aceitas (as do filtro "Carrier" do QuickSight)
TRANSPORTADORAS = {"JAMEF", "FITLOG", "Mira", "INHOUSE"}

TIMEOUT_DOWNLOAD_MS = 300_000
TIMEOUT_PADRAO_MS = 60_000
DIR_SAIDA = Path("saida")
DIR_DEBUG = Path("debug")

MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
         "agosto", "setembro", "outubro", "novembro", "dezembro"]

REGIAO_UF = {
    "AC": "Norte", "AP": "Norte", "AM": "Norte", "PA": "Norte", "RO": "Norte", "RR": "Norte", "TO": "Norte",
    "AL": "Nordeste", "BA": "Nordeste", "CE": "Nordeste", "MA": "Nordeste", "PB": "Nordeste",
    "PE": "Nordeste", "PI": "Nordeste", "RN": "Nordeste", "SE": "Nordeste",
    "DF": "Centro-Oeste", "GO": "Centro-Oeste", "MT": "Centro-Oeste", "MS": "Centro-Oeste",
    "ES": "Sudeste", "MG": "Sudeste", "RJ": "Sudeste", "SP": "Sudeste",
    "PR": "Sul", "RS": "Sul", "SC": "Sul",
}

# Faixas oficiais de CEP por UF (5 primeiros dígitos)
FAIXAS_CEP = [
    (1000, 19999, "SP"), (20000, 28999, "RJ"), (29000, 29999, "ES"), (30000, 39999, "MG"),
    (40000, 48999, "BA"), (49000, 49999, "SE"), (50000, 56999, "PE"), (57000, 57999, "AL"),
    (58000, 58999, "PB"), (59000, 59999, "RN"), (60000, 63999, "CE"), (64000, 64999, "PI"),
    (65000, 65999, "MA"), (66000, 68899, "PA"), (68900, 68999, "AP"), (69000, 69299, "AM"),
    (69300, 69399, "RR"), (69400, 69899, "AM"), (69900, 69999, "AC"), (70000, 72799, "DF"),
    (72800, 72999, "GO"), (73000, 73699, "DF"), (73700, 76799, "GO"), (76800, 76999, "RO"),
    (77000, 77999, "TO"), (78000, 78899, "MT"), (79000, 79999, "MS"), (80000, 87999, "PR"),
    (88000, 89999, "SC"), (90000, 99999, "RS"),
]


def log(msg):
    print(f"[{time.strftime('%H:%M:%S')}] {msg}", flush=True)


def sem_acento(s):
    return unicodedata.normalize("NFD", str(s or "")).encode("ascii", "ignore").decode().upper().strip()


# ── 1. QuickSight ───────────────────────────────────────────────

def exportar_quicksight(dia, headed=False):
    """Baixa o CSV da tabela RAW DATA com a Fecha de Entrega = dia."""
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
    # Reaproveita o login e a abertura de menus do bot que já funciona
    sys.path.insert(0, str(Path(__file__).parent))
    import quicksight_export as qs

    conta = os.environ.get("QS_ACCOUNT", "advanceanalytics")
    usuario, senha = os.environ.get("QS_USER"), os.environ.get("QS_PASSWORD")
    if not usuario or not senha:
        sys.exit("Faltam as variáveis QS_USER e QS_PASSWORD")

    with sync_playwright() as p:
        navegador = p.chromium.launch(headless=not headed, args=["--disable-blink-features=AutomationControlled"])
        contexto = navegador.new_context(viewport={"width": 1920, "height": 1080}, accept_downloads=True, locale="pt-BR")
        page = contexto.new_page()
        try:
            qs.fazer_login(page, conta, usuario, senha)

            url = f"https://{QS_REGION}.quicksight.aws.amazon.com/sn/account/{conta}/dashboards/{DASHBOARD_ID}"
            log("Abrindo o dashboard Cadena de Suministro")
            page.goto(url, wait_until="domcontentloaded", timeout=TIMEOUT_PADRAO_MS)
            page.wait_for_timeout(8000)
            qs.fechar_popups(page)
            abrir_aba(page, qs)
            definir_data(page, dia, qs)
            log("Conferindo os demais filtros")
            conferir_filtros(page, qs)
            return baixar_visual(page, qs)
        except PWTimeout as e:
            qs.salvar_debug(page, "entregues_timeout")
            sys.exit(f"Timeout no QuickSight: {e}")
        except Exception as e:
            qs.salvar_debug(page, "entregues_erro")
            sys.exit(f"Erro no QuickSight: {e}")
        finally:
            contexto.close()
            navegador.close()


def abrir_aba(page, qs):
    selecionada = page.locator("[data-automation-id='selectedTab_sheet_name']")
    if selecionada.count() > 0:
        ctx = selecionada.first.get_attribute("data-automation-context") or ""
        if ABA_DASHBOARD.lower() == ctx.strip().lower():
            log(f"Aba '{ABA_DASHBOARD}' já está selecionada")
            return
    log(f"Clicando na aba '{ABA_DASHBOARD}'")
    aba = page.get_by_text(ABA_DASHBOARD, exact=True).first
    aba.wait_for(state="visible", timeout=TIMEOUT_PADRAO_MS)
    aba.click()
    page.wait_for_timeout(6000)
    qs.fechar_popups(page)


def definir_data(page, dia, qs):
    """Coloca o mesmo dia nos dois campos do controle 'Fecha de Entrega'
    (formato YYYY/MM/DD). Os outros controles — Country, Reschedule Reason,
    Carrier, Shipping State — ficam como estão salvos para o usuário do bot;
    o resultado é conferido depois do download.

    Estrutura do QuickSight (confirmada no HTML):
      div[data-automation-id=sheet_control][data-automation-context="Fecha de Entrega"]
        input[data-automation-id=date_picker_0]   → Data de início
        input[data-automation-id=date_picker_1]   → Data de término
    """
    texto = dia.strftime("%Y/%m/%d")
    log(f"Definindo '{CONTROLE_DATA}' = {texto} a {texto}")
    controle = page.locator(f"[data-automation-id='sheet_control'][data-automation-context='{CONTROLE_DATA}']").first

    # O painel "Controles" pode vir recolhido — abre se precisar
    try:
        controle.wait_for(state="visible", timeout=15_000)
    except Exception:
        log("Controles recolhidos — abrindo o painel")
        try:
            page.get_by_text("Controles", exact=True).first.click(timeout=5000)
        except Exception:
            page.get_by_text("Controls", exact=True).first.click(timeout=5000)
        page.wait_for_timeout(2000)
        controle.wait_for(state="visible", timeout=TIMEOUT_PADRAO_MS)

    # Primeiro a data de TÉRMINO, depois a de INÍCIO: se o início for
    # preenchido antes, por um instante ele fica depois do término antigo,
    # o campo vira "inválido" e o QuickSight abre um aviso por cima.
    for i, nome in ((1, "término"), (0, "início")):
        campo = controle.locator(f"input[data-automation-id='date_picker_{i}']").first
        if campo.count() == 0:
            qs.salvar_debug(page, "campos_data_nao_encontrados")
            raise RuntimeError(f"Não achei o campo de data de {nome} do controle '{CONTROLE_DATA}'")
        fechar_popovers(page)
        # focus + teclado em vez de clique: um calendário/aviso aberto por
        # cima não atrapalha
        campo.focus()
        page.keyboard.press("Control+A")
        page.keyboard.press("Backspace")
        page.keyboard.type(texto, delay=60)
        page.keyboard.press("Tab")
        page.wait_for_timeout(1500)
        fechar_popovers(page)
        valor = (campo.input_value() or "").strip()
        if valor != texto:
            campo.fill(texto)
            campo.dispatch_event("change")
            campo.dispatch_event("blur")
            page.wait_for_timeout(1500)
            fechar_popovers(page)
            valor = (campo.input_value() or "").strip()
        log(f"  data de {nome}: {valor}")
        if valor != texto:
            qs.salvar_debug(page, "data_nao_aceita")
            raise RuntimeError(f"O campo de data de {nome} ficou com '{valor}' em vez de '{texto}'")

    # Conferência final dos dois campos
    for i, nome in ((0, "início"), (1, "término")):
        v = (controle.locator(f"input[data-automation-id='date_picker_{i}']").first.input_value() or "").strip()
        if v != texto:
            qs.salvar_debug(page, "data_nao_aceita")
            raise RuntimeError(f"Depois de preencher, a data de {nome} está '{v}' (esperado '{texto}')")

    page.keyboard.press("Escape")
    log("Aguardando os visuais atualizarem")
    page.wait_for_timeout(10000)
    try:
        page.wait_for_load_state("networkidle", timeout=30_000)
    except Exception:
        pass


# Filtros que precisam estar assim no QuickSight (o bot confere antes de exportar).
# "exato": tem que ser exatamente esses valores; senão, basta conter todos.
FILTROS_ESPERADOS = {
    "Country":           {"valores": ["Brazil"], "exato": True},
    "Reschedule Reason": {"valores": ["NULO", "String vazia", "Failed Attempt"], "exato": True},
    "Carrier":           {"valores": ["Inhouse_Brasil", "JAMEF", "Mira Transportes",
                                      "FITLOG TRANSPORTES E LOGISTICA LTDA"], "exato": False},
    "Shipping State":    {"valores": ["Delivered"], "exato": True},
}


def fechar_popovers(page):
    """Fecha calendário/aviso do Material UI que fique aberto por cima
    dos campos (ele intercepta os cliques)."""
    for _ in range(3):
        try:
            pop = page.locator(".MuiPopover-root")
            if pop.count() == 0 or not pop.first.is_visible():
                return
        except Exception:
            return
        page.keyboard.press("Escape")
        page.wait_for_timeout(600)


def conferir_filtros(page, qs):
    """Lê o valor mostrado em cada controle e confere com o esperado.

    Estrutura (confirmada no HTML do Carrier):
      [data-automation-id=sheet_control_value][data-automation-context="<nome>"]
      → texto "Inhouse_Brasil, JAMEF, Mira Transportes, FITLOG ..."

    Se um filtro estiver diferente, o bot PARA sem exportar — exportar com
    filtro errado gravaria dados errados na BASE. A correção é ajustar o
    filtro no QuickSight logado com o usuário do bot (ele fica salvo)."""
    problemas = []
    for nome, regra in FILTROS_ESPERADOS.items():
        el = page.locator(f"[data-automation-id='sheet_control_value'][data-automation-context='{nome}']").first
        try:
            el.wait_for(state="attached", timeout=10_000)
            texto = (el.inner_text(timeout=5000) or "").strip()
        except Exception:
            log(f"  filtro '{nome}': não encontrado na tela — conferência pulada")
            continue
        atuais = {sem_acento(v) for v in re.split(r"\s*,\s*", texto) if v.strip()}
        esperados = {sem_acento(v) for v in regra["valores"]}
        faltam = esperados - atuais
        sobram = atuais - esperados
        ok = not faltam and (not regra["exato"] or not sobram)
        log(f"  filtro '{nome}': {texto or '(vazio)'} {'✓' if ok else '✗'}")
        if not ok:
            detalhe = []
            if faltam:
                detalhe.append("faltam " + ", ".join(sorted(faltam)))
            if sobram and regra["exato"]:
                detalhe.append("sobram " + ", ".join(sorted(sobram)))
            problemas.append(f"{nome} ({'; '.join(detalhe)})")
    if problemas:
        qs.salvar_debug(page, "filtros_diferentes")
        raise RuntimeError("Filtros do QuickSight diferentes do esperado: " + " | ".join(problemas) +
                           ". Ajuste no QuickSight logado com o usuário do bot e rode de novo. Nada foi gravado.")


def baixar_visual(page, qs):
    """Exporta a tabela RAW DATA.

    Estrutura (confirmada no HTML):
      bloco  = [data-automation-id=grid-block] que contém o título
               h2[data-automation-id=analysis_visual_title_label]
               [data-automation-context="Total Volume to be delivered by week and % On Time"]
      células= [data-automation-id^=sn-table-cell-]
      opção  = li[data-automation-id=dashboard_visual_dropdown_export] ("Exportar para CSV")

    Ordem: 1) espera a tabela terminar de carregar (aparecerem células ou o
    aviso de "sem dados"); 2) só então abre o menu ⋮ e exporta. Abrir o menu
    antes da tabela carregar deixa "Exportar para CSV" desabilitado."""
    log(f"Procurando a tabela '{TITULO_VISUAL}'")
    titulo_sel = f"[data-automation-id='analysis_visual_title_label'][data-automation-context^='{TITULO_VISUAL}']"
    sel_bloco = f"[data-automation-id='grid-block']:has({titulo_sel})"
    celulas_sel = "[data-automation-id^='sn-table-cell-']"

    def achar_bloco():
        for _ in range(25):
            b = page.locator(sel_bloco).first
            try:
                if b.count() > 0:
                    b.scroll_into_view_if_needed(timeout=5000)
                    return b
            except Exception:
                pass
            page.mouse.wheel(0, 700)
            page.wait_for_timeout(900)
        return None

    bloco = achar_bloco()
    if bloco is None:
        qs.salvar_debug(page, "tabela_raw_data_nao_encontrada")
        raise RuntimeError(f"Não achei a tabela '{TITULO_VISUAL}'")

    # 1) espera os dados da tabela (depois da troca da data ela recarrega)
    log("Esperando a tabela carregar com o dia escolhido")
    vazio_textos = ("Nenhum dado", "Sem dados", "No data", "Não há dados", "No hay datos")
    inicio, estado = time.time(), None
    while time.time() - inicio < 120:
        bloco = page.locator(sel_bloco).first
        try:
            if bloco.locator(celulas_sel).count() > 0:
                estado = "dados"
                break
            texto = bloco.inner_text(timeout=3000) or ""
            if any(t.lower() in texto.lower() for t in vazio_textos):
                estado = "vazio"
                break
        except Exception:
            pass
        page.wait_for_timeout(3000)
    if estado == "vazio":
        log("A tabela RAW DATA está vazia para esse dia — nenhum item entregue")
        return None
    if estado is None:
        qs.salvar_debug(page, "tabela_nao_carregou")
        raise RuntimeError("A tabela RAW DATA não terminou de carregar em 2 minutos — veja debug/")
    log("  tabela carregada")
    page.wait_for_timeout(3000)

    opcao = page.locator("li[data-automation-id='dashboard_visual_dropdown_export']")

    def opcao_visivel():
        try:
            for k in range(opcao.count()):
                o = opcao.nth(k)
                if o.is_visible():
                    return o
        except Exception:
            pass
        return None

    def habilitada(o):
        try:
            return o is not None and (o.get_attribute("aria-disabled") or "false") != "true" and \
                   "disabled" not in (o.get_attribute("class") or "").lower()
        except Exception:
            return False

    def passar_mouse(b):
        try:
            caixa = b.bounding_box(timeout=5000)
        except Exception:
            caixa = None
        if caixa:
            page.mouse.move(caixa["x"] + caixa["width"] / 2, caixa["y"] + 20)
            page.wait_for_timeout(400)
            page.mouse.move(caixa["x"] + caixa["width"] - 60, caixa["y"] + 15)
            page.wait_for_timeout(900)

    # Botões que NÃO são o menu (maximizar, ordenar, filtros etc.)
    nao_menu = ("maxim", "expand", "orden", "sort", "filtr", "foco", "focus", "favorit")

    def abrir_menu():
        b = achar_bloco()
        if b is None:
            return None
        passar_mouse(b)
        # 1º: o botão de menu do visual pelo identificador, se existir
        candidatos = []
        direto = b.locator("[data-automation-id='analysis_visual_dropdown_menu_button'], button[aria-label*='Menu'], button[aria-label*='menu'], button[aria-label*='opções' i], button[aria-label*='options' i]")
        for k in range(direto.count()):
            candidatos.append((99999, direto.nth(k)))
        # 2º: demais botões pequenos do bloco, da direita para a esquerda
        botoes = b.locator("button")
        outros = []
        for k in range(botoes.count()):
            bt = botoes.nth(k)
            try:
                if not bt.is_visible(timeout=300):
                    continue
                rot = (bt.get_attribute("aria-label") or bt.get_attribute("title") or "").lower()
                if any(n in rot for n in nao_menu):
                    continue
                cx = bt.bounding_box()
                if cx and cx["width"] <= 60:
                    outros.append((cx["x"], bt))
            except Exception:
                continue
        outros.sort(key=lambda t: t[0], reverse=True)
        for _, bt in candidatos + outros[:8]:
            try:
                passar_mouse(b)
                bt.click(timeout=4000)
                page.wait_for_timeout(1200)
                o = opcao_visivel()
                if o is not None:
                    return o
                page.keyboard.press("Escape")
                page.wait_for_timeout(500)
            except Exception:
                continue
        return None

    # 2) abre o menu e exporta (até 6 tentativas, ~1 min)
    alvo = None
    for tentativa in range(6):
        alvo = abrir_menu()
        if alvo is not None and habilitada(alvo):
            break
        if alvo is not None:
            log("  'Exportar para CSV' ainda desabilitado — fechando o menu e esperando")
            page.mouse.click(5, 5)          # fecha o menu clicando fora (sem Esc)
        else:
            log(f"  menu da tabela não abriu (tentativa {tentativa + 1})")
        page.wait_for_timeout(10_000)
        alvo = None
    if alvo is None:
        qs.salvar_debug(page, "menu_raw_data")
        raise RuntimeError("Não consegui abrir o menu da tabela RAW DATA com 'Exportar para CSV' habilitado — veja debug/")

    with page.expect_download(timeout=TIMEOUT_DOWNLOAD_MS) as info:
        alvo.click()
        log("Exportando — o QuickSight prepara o arquivo antes de baixar")
    DIR_SAIDA.mkdir(exist_ok=True)
    destino = DIR_SAIDA / "entregues_raw_data.csv"
    info.value.save_as(str(destino))
    log(f"CSV baixado: {destino} ({destino.stat().st_size / 1024:.0f} KB)")
    return destino


# ── Leitura e conferência do CSV ────────────────────────────────

def data_iso(v):
    """'2026-10-09 00:00:00' ou '09/10/2026' → date"""
    s = str(v or "").strip()
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return dt.date(int(m[1]), int(m[2]), int(m[3]))
    m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m:
        return dt.date(int(m[3]), int(m[2]), int(m[1]))
    return None


def transportadora(arrangement):
    a = sem_acento(arrangement)
    if a.startswith("JAMEF"):
        return "JAMEF"
    if a.startswith("FITLOG"):
        return "FITLOG"
    if a.startswith("MIRA"):
        return "Mira"
    if a.startswith("INHOUSE"):
        return "INHOUSE"
    return ""


def ler_csv(caminho, dia):
    with open(caminho, "r", encoding="utf-8-sig", newline="") as f:
        linhas = list(csv.DictReader(f))
    log(f"Export tem {len(linhas)} linha(s)")
    itens, descartes = [], {}
    for r in linhas:
        r = {k.strip(): (v or "").strip() for k, v in r.items() if k}
        entrega = data_iso(r.get("DELIVERED_DATE"))
        car = transportadora(r.get("ARRANGEMENT"))
        motivo = None
        if not r.get("SCI"):
            motivo = "sem SCI"
        elif entrega != dia:
            motivo = f"entrega em {entrega} (esperado {dia})"
        elif sem_acento(r.get("SHIPPING_STATE")) != "DELIVERED":
            motivo = f"Shipping State {r.get('SHIPPING_STATE')}"
        elif car not in TRANSPORTADORAS:
            motivo = f"transportadora fora da lista ({r.get('ARRANGEMENT')[:30]})"
        if motivo:
            descartes[motivo] = descartes.get(motivo, 0) + 1
            continue
        r["_entrega"], r["_carrier"] = entrega, car
        itens.append(r)
    for m, q in descartes.items():
        log(f"  descartadas {q} linha(s): {m}")
    # Trava: se quase tudo veio de outro dia, o filtro de data não pegou
    if linhas and len(itens) == 0:
        raise RuntimeError("Nenhuma linha do export é do dia esperado — o filtro de data não foi aplicado. Nada foi gravado.")
    return itens


# ── 2. Zecore (Shipping Core) ───────────────────────────────────

class Zecore:
    """Sessão no Zecore para ler o Shipping Core em formato de dados (JSON).

    Entra pelo login padrão do Frappe (usuário + senha, o mesmo do bot de
    baixas). Se o servidor recusar esse caminho, faz o login pela tela —
    exatamente como o zecore_atualizar_entregas.py — e reaproveita a sessão
    do navegador."""

    def __init__(self, usuario, senha, headed=False):
        self.s = requests.Session()
        self.s.headers.update({"Accept": "application/json"})
        try:
            r = self.s.post(f"{ZECORE_URL}/api/method/login", data={"usr": usuario, "pwd": senha}, timeout=60)
            if r.status_code == 200 and self._sessao_ok():
                log("Login no Zecore ok")
                return
            log(f"Login direto no Zecore recusado (HTTP {r.status_code}) — tentando pela tela")
        except requests.RequestException as e:
            log(f"Login direto no Zecore falhou ({e}) — tentando pela tela")
        self._login_pela_tela(usuario, senha, headed)
        if not self._sessao_ok():
            raise RuntimeError("Não consegui abrir sessão no Zecore (nem direto, nem pela tela)")
        log("Login no Zecore ok (pela tela)")

    def _sessao_ok(self):
        try:
            r = self.s.get(f"{ZECORE_URL}/api/method/frappe.auth.get_logged_user", timeout=30)
            return r.status_code == 200 and (r.json().get("message") or "") not in ("", "Guest")
        except Exception:
            return False

    def _login_pela_tela(self, usuario, senha, headed):
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            nav = p.chromium.launch(headless=not headed, args=["--disable-blink-features=AutomationControlled"])
            ctx = nav.new_context(viewport={"width": 1600, "height": 1000}, locale="pt-BR")
            page = ctx.new_page()
            try:
                page.goto(f"{ZECORE_URL}/login", wait_until="domcontentloaded", timeout=TIMEOUT_PADRAO_MS)
                campo_user = page.locator("#login_email, input[name='email'], input[type='email']").first
                campo_user.wait_for(state="visible", timeout=TIMEOUT_PADRAO_MS)
                campo_user.fill(usuario)
                page.locator("#login_password, input[type='password']").first.fill(senha)
                page.keyboard.press("Enter")
                page.wait_for_url("**/app**", timeout=TIMEOUT_PADRAO_MS)
                for c in ctx.cookies():
                    if "zebrands" in (c.get("domain") or ""):
                        self.s.cookies.set(c["name"], c["value"], domain=c["domain"], path=c.get("path", "/"))
            finally:
                ctx.close()
                nav.close()

    def shipping_core(self, docname):
        tipo = "Free Sales Order Request" if docname.upper().startswith("FREE-") else "Sales Order"
        nome = f"{tipo}-{docname}-shipping"
        for tentativa in range(3):
            try:
                r = self.s.get(f"{ZECORE_URL}/api/resource/Shipping Core/{requests.utils.quote(nome, safe='')}", timeout=60)
                if r.status_code == 404:
                    return None
                if r.status_code == 200:
                    return r.json().get("data")
                log(f"  {docname}: HTTP {r.status_code}, tentando de novo")
            except requests.RequestException as e:
                log(f"  {docname}: {e}, tentando de novo")
            time.sleep(2 * (tentativa + 1))
        return None


def total_da_invoice(zc, nome, cache):
    """Valor total da NF no Zecore (Sales Invoice → grand_total)."""
    if not nome:
        return None
    if nome in cache:
        return cache[nome]
    v = None
    try:
        r = zc.s.get(f"{ZECORE_URL}/api/resource/Sales Invoice/{requests.utils.quote(nome, safe='')}", timeout=60)
        if r.status_code == 200:
            d = r.json().get("data") or {}
            v = d.get("grand_total") if d.get("grand_total") is not None else d.get("rounded_total")
            v = float(v) if v is not None else None
    except Exception:
        v = None
    cache[nome] = v
    time.sleep(0.1)
    return v


def nf_da_chave(texto):
    """A chave da NF-e (44 dígitos) traz o número da nota nas posições 26-34.
    Ex.: .../35260742418313000104550010001064711006560589.xml → 106471"""
    m = re.search(r"(\d{44})", str(texto or ""))
    return str(int(m.group(1)[25:34])) if m else ""


def dados_zecore(itens, zc):
    """Busca um Shipping Core por pedido e devolve, por SCI:
    criação, NF e endereço."""
    por_sci, pedidos = {}, sorted({r["DOCNAME"] for r in itens})
    log(f"Consultando {len(pedidos)} pedido(s) no Zecore")
    faltando = []
    for i, doc in enumerate(pedidos, 1):
        sc = zc.shipping_core(doc)
        if not sc:
            faltando.append(doc)
            continue
        criacao = data_iso(sc.get("creation"))
        for it in sc.get("items") or []:
            nf = re.sub(r"^\d+-", "", str(it.get("invoice_number") or "").strip())   # "1-106471" → "106471"
            if not nf:
                nf = nf_da_chave(it.get("sales_invoice_xml"))   # plano B: número dentro da chave da NF-e
            end = re.sub(r"-Shipping-\d+$", "", str(it.get("address") or "").strip())
            por_sci[it.get("name")] = {"criacao": criacao, "nf": nf, "endereco": end,
                                       "invoice": str(it.get("sales_invoice") or "").strip()}
        if i % 20 == 0:
            log(f"  {i}/{len(pedidos)} pedidos")
        time.sleep(0.15)
    if faltando:
        log(f"  {len(faltando)} pedido(s) sem Shipping Core: {', '.join(faltando[:5])}{'…' if len(faltando) > 5 else ''}")
    return por_sci


# ── 3. Rastreio das transportadoras ─────────────────────────────

def br_data(s):
    """'27/04/26 10:15', '2026-04-27T10:15', '27/04/2026' → date"""
    s = str(s or "").strip()
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return dt.date(int(m[1]), int(m[2]), int(m[3]))
    m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{2,4})", s)
    if m:
        a = int(m[3]) + (2000 if len(m[3]) == 2 else 0)
        return dt.date(a, int(m[2]), int(m[1]))
    return None


def eh_entrega(texto):
    t = sem_acento(texto)
    return ("ENTREGA REALIZADA" in t or "ENTREGUE" in t or "ENTREGA EFETUADA" in t) and "NAO" not in t and "PREJUDICADA" not in t


class Jamef:
    def __init__(self, usuario, senha):
        r = requests.post(f"{JAMEF_API}/auth/v1/login", json={"username": usuario, "password": senha},
                          headers={"accept": "application/json"}, timeout=60)
        if r.status_code != 200:
            raise RuntimeError(f"Login na API Jamef falhou (HTTP {r.status_code}): {r.text[:200]}")
        self.token = r.json()["dado"][0]["accessToken"]
        log("Login na API Jamef ok")

    def consultar(self, nf):
        url = (f"{JAMEF_API}/consulta/v1/rastreamento?documentoPagadorFrete={CNPJ}"
               f"&documentoRemetente={CNPJ}&numeroNotaFiscal={nf}")
        for tentativa in range(3):
            try:
                r = requests.get(url, headers={"accept": "application/json", "Authorization": f"Bearer {self.token}"}, timeout=60)
                if r.status_code in (429,) or r.status_code >= 500:
                    time.sleep(3 * (tentativa + 1))
                    continue
                if r.status_code != 200:
                    return {}
                doc = ((r.json().get("dado") or [{}])[0].get("rastreamento") or [None])[0]
                if not doc:
                    return {}
                eventos = doc.get("eventosRastreio") or []          # mais recente primeiro
                entrega = next((br_data(e.get("data")) for e in eventos if eh_entrega(e.get("status"))), None)
                prev = br_data((doc.get("frete") or {}).get("previsaoEntrega"))
                return {"entrega": entrega, "previsao": prev}
            except (requests.RequestException, ValueError, KeyError, IndexError):
                time.sleep(3 * (tentativa + 1))
        return {}


def ssw_lote(nfs, senha):
    """Consulta várias NFs de uma vez no SSW. Devolve {nf: {entrega, previsao}}.
    O SSW em lote traz o último evento de cada NF — para NF entregue, é a entrega."""
    payload = {"cnpj": CNPJ, "NR": "\r\n".join(nfs), "chave": senha}
    r = requests.post(SSW_URL, data=payload, timeout=90, headers={
        "User-Agent": "Mozilla/5.0", "Origin": "https://ssw.inf.br", "Referer": "https://ssw.inf.br/2/rastreamento?"})
    r.raise_for_status()
    html = r.text
    saida = {}
    for bloco in re.split(r"<tr[^>]+onclick=", html, flags=re.I)[1:]:
        m_nf = re.search(r"<label[^>]*class=[\"']?rastreamento[\"']?[^>]*>([\s\S]*?)</label>", bloco, re.I)
        if not m_nf:
            continue
        tokens = re.findall(r"\d{4,}", m_nf.group(1))
        m_dt = re.search(r"<p[^>]*class=[\"']?tdb[\"']?[^>]*>[\s\S]*?<br>\s*(\d{2}/\d{2}/\d{2,4})", bloco, re.I)
        m_st = re.search(r"<p[^>]*class=[\"']?titulo[\"']?[^>]*>\s*([^<]+)", bloco, re.I)
        m_pv = re.search(r"[Pp]revis[aã]o de entrega[:\s]+(\d{2}/\d{2}/\d{2,4})", bloco)
        situacao = m_st.group(1).strip() if m_st else ""
        dados = {"entrega": br_data(m_dt.group(1)) if (m_dt and eh_entrega(situacao)) else None,
                 "previsao": br_data(m_pv.group(1)) if m_pv else None}
        for t in tokens:
            saida[t] = dados
            saida[str(int(t))] = dados
    return saida


def rastrear(itens, por_sci, jamef):
    """Devolve {(transportadora, nf): {entrega, previsao}}."""
    nfs = {}
    for r in itens:
        z = por_sci.get(r["SCI"]) or {}
        if z.get("nf") and r["_carrier"] != "INHOUSE":
            nfs.setdefault(r["_carrier"], set()).add(z["nf"])
    out = {}
    if nfs.get("JAMEF") and jamef:
        log(f"Rastreando {len(nfs['JAMEF'])} NF(s) Jamef pela API")
        for nf in sorted(nfs["JAMEF"]):
            out[("JAMEF", nf)] = jamef.consultar(nf)
            time.sleep(0.3)
    for car in ("FITLOG", "Mira"):
        lista = sorted(nfs.get(car, []))
        if not lista:
            continue
        log(f"Rastreando {len(lista)} NF(s) {car} pelo SSW")
        for i in range(0, len(lista), 100):
            try:
                res = ssw_lote(lista[i:i + 100], SSW_SENHAS[car])
            except requests.RequestException as e:
                log(f"  SSW {car} falhou: {e}")
                continue
            for nf in lista[i:i + 100]:
                out[(car, nf)] = res.get(nf) or res.get(str(int(nf))) or {}
            time.sleep(0.5)
    return out


# ── 3b. Valor do frete e valor da NF ────────────────────────────
# Mesma origem que a Dash fedex usa para as colunas V e W da Rastreamento:
#   • Jamef  → relatório do Portal do Cliente: coluna X (valor do CT-e = frete)
#              e W (valor da mercadoria); NF na coluna BC
#   • Mira   → relatório da Mira: colunas "frete" e "valorMercadoria"
#   • Fitlog → XML do CT-e no SSW (serviço id=52): vTPrest (frete) e vCarga (valor)
#   • Inhouse→ sem frete de transportadora; valor da NF vem do Zecore
# Os relatórios da Jamef e da Mira exigem login (MFA pelo Gmail / sessão da
# Mira) e já são baixados todo dia pelas automações da Dash fedex; o bot lê
# esses relatórios — não a coluna V da Rastreamento, que perde a nota quando
# ela é entregue.

def valor_num(v):
    """'1.399,00', '186,15', '1,399.00' ou número → float (None se vazio)."""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"R\$|\s", "", str(v))
    if not s:
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    elif "." in s:
        partes = s.split(".")
        if len(partes) > 2 or len(partes[-1]) == 3:
            s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return None


def nf_numero(v):
    d = re.sub(r"\D", "", re.sub(r"-\d+$", "", str(v or "").strip()))
    return str(int(d)) if d else ""


def valores_relatorios(gc):
    """{(transportadora, nf): (frete, valor)} dos relatórios Jamef e Mira."""
    out = {}
    try:
        planilha = gc.open_by_key(os.environ.get("DASHFEDEX_SHEET_ID", DASHFEDEX_ID_PADRAO))
    except Exception as e:
        log(f"  não consegui abrir a planilha dos relatórios ({e}) — frete Jamef/Mira em branco")
        return out
    # Jamef: NF = BC (55), valor = W (23), frete = X (24)
    try:
        linhas = planilha.worksheet(ABA_JAMEF_REL).get_all_values()[1:]
        for l in linhas:
            if len(l) < 55:
                continue
            nf = nf_numero(l[54])
            if nf:
                out[("JAMEF", nf)] = (valor_num(l[23]), valor_num(l[22]))
        log(f"  relatório Jamef: {sum(1 for k in out if k[0] == 'JAMEF')} NF(s)")
    except Exception as e:
        log(f"  relatório Jamef indisponível ({e})")
    # Mira: colunas pelo nome do cabeçalho
    try:
        todas = planilha.worksheet(ABA_MIRA_REL).get_all_values()
        cab = [c.strip().lower() for c in (todas[0] if todas else [])]
        def col(*nomes):
            for n_ in nomes:
                if n_.lower() in cab:
                    return cab.index(n_.lower())
            for n_ in nomes:
                for i, c in enumerate(cab):
                    if n_.lower() in c:
                        return i
            return -1
        i_nf, i_fr, i_vl = col("notaFiscal", "nota fiscal", "nf"), col("frete"), col("valorMercadoria", "valor mercadoria")
        q = 0
        for l in todas[1:]:
            if i_nf < 0 or i_nf >= len(l):
                continue
            ult = (l[i_nf].split() or [""])[-1]          # "1 000110117" → 110117
            nf = nf_numero(ult)
            if nf:
                out[("Mira", nf)] = (valor_num(l[i_fr]) if 0 <= i_fr < len(l) else None,
                                     valor_num(l[i_vl]) if 0 <= i_vl < len(l) else None)
                q += 1
        log(f"  relatório Mira: {q} NF(s)")
    except Exception as e:
        log(f"  relatório Mira indisponível ({e})")
    return out


def xml_cte_fitlog(nf, senha):
    """Baixa o XML do CT-e da Fitlog pelo SSW e devolve (frete, valor) ou None.
    Mesmo caminho do Apps Script (_baixarXmlCteFitlog): formulário id=52 →
    link 'ssw…?id=…' → arquivo (XML ou ZIP)."""
    import io, zipfile
    ua = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131.0.0.0 Safari/537.36"
    ses = requests.Session()
    resp = ses.post(URL_SSW_SERVICO, data={"action": "xml", "id": "52", "cnpj": CNPJ, "password": senha, "nronf": nf},
                    headers={"User-Agent": ua, "Origin": "https://ssw.inf.br", "Referer": "https://ssw.inf.br/2/servico?id=52"}, timeout=60)
    for _ in range(4):
        if resp.status_code != 200:
            return None
        corpo = resp.content
        xml = None
        if corpo[:2] == b"PK":
            with zipfile.ZipFile(io.BytesIO(corpo)) as z:
                for nome in z.namelist():
                    t = z.read(nome).decode("utf-8", "ignore")
                    if "<vTPrest>" in t:
                        xml = t
                        break
        else:
            t = corpo.decode("utf-8", "ignore")
            if "<vTPrest>" in t:
                xml = t
        if xml:
            chaves = re.findall(r"<infNFe>\s*<chave>(\d{44})</chave>", xml)
            nfs = [str(int(c[25:34])) for c in chaves]
            if nfs and str(int(nf)) not in nfs:
                return None                       # CT-e de outra NF
            fr = re.search(r"<vTPrest>([^<]*)</vTPrest>", xml)
            vl = re.search(r"<vCarga>([^<]*)</vCarga>", xml)
            return (float(fr.group(1)) if fr else None, float(vl.group(1)) if vl else None)
        html = corpo.decode("utf-8", "ignore")
        m = (re.search(r"href=[\"'](https?://[^\"']*ssw\d+\?id=[^\"']+)[\"']", html, re.I)
             or re.search(r"href=[\"']([^\"']*ssw\d+\?id=[^\"']+)[\"']", html, re.I)
             or re.search(r"(/cgi-local/ssw\d+\?id=[0-9A-Za-z]+)", html)
             or re.search(r"(ssw\d+\?id=[0-9A-Za-z]+)", html))
        if not m:
            return None
        link = m.group(1).replace("&amp;", "&")
        if not link.lower().startswith("http"):
            link = "https://ssw.inf.br" + link if link.startswith("/") else (
                "https://ssw.inf.br/cgi-local/" + link if re.match(r"ssw\d+\?", link) else "https://ssw.inf.br/2/" + link)
        resp = ses.get(link, headers={"User-Agent": ua, "Referer": URL_SSW_SERVICO}, timeout=60)
    return None


def buscar_valores(itens, por_sci, gc, zc):
    """{SCI: (frete, valor_nf)}"""
    log("Buscando valor do frete e valor da NF")
    rel = valores_relatorios(gc) if gc else {}
    senha_fitlog = os.environ.get("FITLOG_SENHA_XML") or SSW_SENHAS["FITLOG"]
    xml_cache, inv_cache, out = {}, {}, {}
    fitlog_nfs = sorted({(por_sci.get(r["SCI"]) or {}).get("nf") for r in itens if r["_carrier"] == "FITLOG"} - {None, ""})
    if fitlog_nfs:
        log(f"  Fitlog: baixando o XML do CT-e de {len(fitlog_nfs)} NF(s) no SSW")
        for nf in fitlog_nfs:
            try:
                xml_cache[nf] = xml_cte_fitlog(nf, senha_fitlog)
            except Exception as e:
                xml_cache[nf] = None
            time.sleep(0.4)
    cont = {}
    for r in itens:
        z = por_sci.get(r["SCI"]) or {}
        nf, car = z.get("nf", ""), r["_carrier"]
        frete = valor = None
        if car == "FITLOG" and xml_cache.get(nf):
            frete, valor = xml_cache[nf]
        elif car in ("JAMEF", "Mira") and (car, nf) in rel:
            frete, valor = rel[(car, nf)]
        if valor is None and zc is not None:                 # Inhouse ou sem relatório: valor da NF pelo Zecore
            valor = total_da_invoice(zc, z.get("invoice"), inv_cache)
        out[r["SCI"]] = (frete, valor)
        c = cont.setdefault(car, [0, 0, 0])
        c[0] += 1
        c[1] += 1 if frete is not None else 0
        c[2] += 1 if valor is not None else 0
    log("  frete / valor NF por transportadora: " + " · ".join(f"{k} {v[1]}/{v[2]} de {v[0]}" for k, v in sorted(cont.items())))
    return out


# ── 3c. Coordenadas do endereço ─────────────────────────────────
# Uma vez por endereço, aqui no robô — assim o mapa do painel já abre com
# cada alfinete no lugar certo, para todo mundo, sem esperar a busca no
# navegador. OpenStreetMap (Nominatim): no máximo 1 consulta por segundo.

def consulta_endereco(txt):
    s = str(txt or "").strip()
    if not re.search(r"[A-Za-z]", s):
        return None, None
    m = re.search(r"(\d{5})-?(\d{3})(?!\d)", s)
    cep = f"{m.group(1)}-{m.group(2)}" if m else ""
    s = re.sub(r"N[úu]mero\s+interior:[^,]*,?", " ", s, flags=re.I)
    s = re.sub(r"\d{5}-?\d{3}(?!\d)", " ", s)
    partes = [p.strip() for p in s.split(",") if p.strip()]
    # complemento atrapalha a busca (apto, bloco, sala…). Só sai quando é
    # mesmo complemento: depois do número da rua, ou um trecho que começa com
    # ele — "Rua Casa Branca" e o bairro "Casa Verde" ficam como estão.
    compl = r"(APTO?|APARTAMENTO|APT|AP|BLOCO|BL|SALA|LOJA|TORRE|CONJ|CONJUNTO|FUNDOS|CASA\s*\d|LOTE\s*\d|QD\s*\d|QUADRA\s*\d)"
    if partes:
        m1 = re.match(r"^(.*?\d+)\s+" + compl + r"\b.*$", partes[0], flags=re.I)
        if m1:
            partes[0] = m1.group(1)
    partes = [p for i, p in enumerate(partes) if i == 0 or not re.match(r"^" + compl + r"\b", p, flags=re.I)]
    if partes:
        mm = re.match(r"^(\d+)\s+(.*)$", partes[0])
        if mm:
            partes[0] = f"{mm.group(2)}, {mm.group(1)}"
    q = re.sub(r"\s{2,}", " ", ", ".join(partes) + (f", {cep}" if cep else "") + ", Brasil")
    return q, cep


class Geocoder:
    def __init__(self):
        self.s = requests.Session()
        self.s.headers.update({"User-Agent": NOMINATIM_UA, "Accept": "application/json", "Accept-Language": "pt-BR"})
        self.ultimo, self.cache, self.feitas = 0.0, {}, 0

    def _busca(self, params):
        espera = 1.1 - (time.time() - self.ultimo)
        if espera > 0:
            time.sleep(espera)
        self.ultimo = time.time()
        self.feitas += 1
        try:
            r = self.s.get(NOMINATIM_URL, params={**params, "format": "json", "limit": 1, "countrycodes": "br"}, timeout=30)
            if r.status_code == 200 and r.json():
                h = r.json()[0]
                return round(float(h["lat"]), 6), round(float(h["lon"]), 6)
        except Exception:
            pass
        return None

    def coordenadas(self, endereco):
        q, cep = consulta_endereco(endereco)
        if not q:
            return None
        if q in self.cache:
            return self.cache[q]
        g = self._busca({"q": q})
        if g is None and cep:                    # plano B: só pelo CEP
            g = self._busca({"postalcode": cep, "country": "Brasil"})
        self.cache[q] = g
        return g


def geocodificar_linhas(linhas, tempo_max_s=600):
    """Preenche LAT/LNG nas linhas novas (uma consulta por endereço)."""
    geo, t0, ok = Geocoder(), time.time(), 0
    enderecos = sorted({l["CEP"] for l in linhas if l.get("CEP")})
    log(f"Coordenadas: {len(enderecos)} endereço(s) para localizar")
    achados = {}
    for e in enderecos:
        if time.time() - t0 > tempo_max_s:
            log("  tempo máximo das coordenadas atingido — o resto fica para o modo completar")
            break
        achados[e] = geo.coordenadas(e)
    for l in linhas:
        g = achados.get(l.get("CEP"))
        if g:
            l[COL_LAT], l[COL_LNG] = f"{g[0]:.6f}", f"{g[1]:.6f}"     # texto com ponto: não vira data/número errado
            ok += 1
    log(f"  {ok} de {len(linhas)} item(ns) com coordenadas ({geo.feitas} consulta(s))")


def completar_coordenadas(gc, sheet_id, desde, ate, tempo_max_min, simular=False):
    """Preenche LAT/LNG nas linhas da BASE que têm endereço e ainda não têm coordenada."""
    import gspread.utils as gu
    t0 = time.time()
    aba = gc.open_by_key(sheet_id).worksheet(ABA_BASE)
    todas = aba.get_all_values()
    cab = [c.strip() for c in todas[0]]
    novas_cols = [c for c in COLUNAS_NOVAS if c not in cab]
    if novas_cols and not simular:
        if aba.col_count < len(cab) + len(novas_cols):
            aba.add_cols(len(cab) + len(novas_cols) - aba.col_count)
        aba.update(values=[novas_cols], range_name=gu.rowcol_to_a1(1, len(cab) + 1))
        log(f"colunas criadas na BASE: {', '.join(novas_cols)}")
    cab += novas_cols
    ix = {c: i for i, c in enumerate(cab)}
    g = lambda l, c: (l[ix[c]].strip() if ix[c] < len(l) else "")
    alvo = {}
    for n_lin, l in enumerate(todas[1:], start=2):
        d = data_iso(g(l, "DELIVERED_DATE"))
        if not d or (desde and d < desde) or (ate and d > ate):
            continue
        if g(l, COL_LAT) and g(l, COL_LNG):
            continue
        e = g(l, "CEP")
        if e and re.search(r"[A-Za-z]", e):
            alvo.setdefault(e, []).append((n_lin, d))
    ordem = sorted(alvo, key=lambda e: max(d for _, d in alvo[e]), reverse=True)   # mais recentes primeiro
    log(f"{sum(len(v) for v in alvo.values())} linha(s) sem coordenada · {len(ordem)} endereço(s) diferentes")
    geo, atualiz, k = Geocoder(), [], 0
    col_lat = gu.rowcol_to_a1(1, ix[COL_LAT] + 1)[:-1]
    col_lng = gu.rowcol_to_a1(1, ix[COL_LNG] + 1)[:-1]
    for k, e in enumerate(ordem):
        if time.time() - t0 > tempo_max_min * 60:
            log(f"Tempo máximo atingido em {k}/{len(ordem)} endereços — rode de novo para continuar")
            break
        c = geo.coordenadas(e)
        if c:
            for n_lin, _ in alvo[e]:
                atualiz.append({"range": f"{col_lat}{n_lin}", "values": [[f"{c[0]:.6f}"]]})
                atualiz.append({"range": f"{col_lng}{n_lin}", "values": [[f"{c[1]:.6f}"]]})
        if (k + 1) % 100 == 0:
            log(f"  {k + 1}/{len(ordem)} endereços consultados")
    log(f"{len(atualiz) // 2} linha(s) com coordenadas novas")
    if simular or not atualiz:
        return
    for i in range(0, len(atualiz), 400):
        aba.batch_update(atualiz[i:i + 400], value_input_option="RAW")
        time.sleep(1)
    log("Coordenadas gravadas na BASE")


# ── 4. Montagem das linhas da BASE ──────────────────────────────

def cep_de(texto):
    m = re.findall(r"(\d{5})-?(\d{3})(?!\d)", str(texto or ""))
    return (m[-1][0] + m[-1][1]) if m else ""


def uf_do_cep(cep8):
    if not cep8:
        return ""
    p5 = int(cep8[:5])
    for a, b, uf in FAIXAS_CEP:
        if a <= p5 <= b:
            return uf
    return ""


def fmt(d):
    return d.strftime("%d/%m/%Y") if d else ""


def montar_linha(r, z, tr, val=(None, None)):
    entrega, prometido = r["_entrega"], data_iso(r.get("ORG_DELIVERY_DATE"))
    criacao = z.get("criacao")
    uf = uf_do_cep(cep_de(z.get("endereco")))
    pct = str(r.get("% On Time", "")).replace("%", "").replace(",", ".").strip()
    try:
        on_time = float(pct) > 0
    except ValueError:
        on_time = sem_acento(pct).startswith("ON")
    return {
        "DELIVERED_DATE": fmt(entrega),
        "ORG_DELIVERY_DATE": fmt(prometido),
        "REPORTING_AGGREGATOR_NEW": r.get("REPORTING_AGGREGATOR_NEW", ""),
        "DATA DE CRIAÇÃO": fmt(criacao),
        "S_AND_OP_AGGREGATOR": r.get("S_AND_OP_AGGREGATOR", ""),
        "POS": r.get("POS", ""),
        "DOCNAME": r.get("DOCNAME", ""),
        "SCI": r.get("SCI", ""),
        "SHIPPING_STATE": r.get("SHIPPING_STATE", ""),
        "% On Time": "On time" if on_time else "Late time",
        "NF": z.get("nf", ""),
        "Transportadora": r["_carrier"],
        "Data prevista Zecore": fmt(prometido),
        "Data baixada Zecore": fmt(entrega),
        "Data prevista transportadora": fmt(tr.get("previsao")),
        "Data de entrega Transportadora": fmt(tr.get("entrega")),
        "Baixa API": "",            # manual
        "Motivo do atraso": "",     # manual
        "UF": uf,
        "Região": REGIAO_UF.get(uf, ""),
        "SEMANA": entrega.isocalendar()[1] if entrega else "",
        "MÊS": MESES[entrega.month - 1] if entrega else "",
        "DIAS": (entrega - criacao).days if (entrega and criacao) else "",
        "CEP": z.get("endereco", ""),
        COL_VALOR_FRETE: "" if val[0] is None else round(val[0], 2),
        COL_VALOR_NF: "" if val[1] is None else round(val[1], 2),
        COL_LAT: "", COL_LNG: "",
    }


# ── 5. Gravação na BASE ─────────────────────────────────────────

def cliente_google(creds_json):
    import gspread
    from google.oauth2.service_account import Credentials
    creds = Credentials.from_service_account_info(json.loads(creds_json), scopes=["https://www.googleapis.com/auth/spreadsheets"])
    return gspread.authorize(creds)


def gravar_base(linhas, sheet_id, gc, simular=False):
    aba = gc.open_by_key(sheet_id).worksheet(ABA_BASE)
    cab = [c.strip() for c in aba.row_values(1)]
    # Colunas novas de valores: cria no fim do cabeçalho se ainda não existirem
    novas_cols = [c for c in COLUNAS_NOVAS if c not in cab]
    if novas_cols:
        if simular:
            log(f"  (simulação) colunas que seriam criadas na BASE: {', '.join(novas_cols)}")
        else:
            if aba.col_count < len(cab) + len(novas_cols):
                aba.add_cols(len(cab) + len(novas_cols) - aba.col_count)
            import gspread.utils as gu
            ini = gu.rowcol_to_a1(1, len(cab) + 1)
            aba.update(values=[novas_cols], range_name=ini)
            log(f"  colunas criadas na BASE: {', '.join(novas_cols)}")
        cab += novas_cols
    if "SCI" not in cab:
        raise RuntimeError("A aba BASE não tem a coluna SCI no cabeçalho")
    col_sci = cab.index("SCI") + 1
    ja = {s.strip() for s in aba.col_values(col_sci)[1:] if s.strip()}
    novas = [l for l in linhas if l["SCI"] not in ja]
    log(f"{len(linhas)} item(ns) no export · {len(linhas) - len(novas)} já estavam na BASE · {len(novas)} novo(s)")
    if not novas:
        return 0
    faltam = [c for c in cab if c and c not in novas[0]]
    if faltam:
        log(f"  colunas da BASE sem valor automático (ficam em branco): {', '.join(faltam)}")
    valores = [[l.get(c, "") for c in cab] for l in novas]
    if simular:
        log("Modo --simular: nada foi gravado. Primeiras linhas:")
        for v in valores[:3]:
            log("  " + " | ".join(str(x) for x in v))
        return len(novas)
    aba.append_rows(valores, value_input_option="USER_ENTERED", insert_data_option="INSERT_ROWS")
    log(f"{len(novas)} item(ns) acrescentado(s) na BASE")
    return len(novas)


# ── Modo completar valores (histórico) ──────────────────────────
# Percorre a BASE e preenche VALOR FRETE / VALOR NF só onde estão em branco,
# com as mesmas fontes do dia a dia. Não mexe em nenhuma outra coluna e não
# acrescenta linhas. Pode rodar quantas vezes precisar: cada execução continua
# de onde a anterior parou (só olha o que ainda está vazio).

def _carrier_base(v):
    t = sem_acento(v)
    return "JAMEF" if t.startswith("JAMEF") else "FITLOG" if t.startswith("FITLOG") else \
           "Mira" if t.startswith("MIRA") else "INHOUSE" if t.startswith("INHOUSE") else ""


def completar_valores(gc, zc, sheet_id, desde, ate, tempo_max_min, simular=False):
    import gspread.utils as gu
    t0 = time.time()
    fim_ok = lambda: time.time() - t0 < tempo_max_min * 60
    aba = gc.open_by_key(sheet_id).worksheet(ABA_BASE)
    todas = aba.get_all_values()
    if not todas:
        log("BASE vazia")
        return
    cab = [c.strip() for c in todas[0]]
    novas_cols = [c for c in COLUNAS_NOVAS if c not in cab]
    if novas_cols and not simular:
        if aba.col_count < len(cab) + len(novas_cols):
            aba.add_cols(len(cab) + len(novas_cols) - aba.col_count)
        aba.update(values=[novas_cols], range_name=gu.rowcol_to_a1(1, len(cab) + 1))
        log(f"colunas criadas na BASE: {', '.join(novas_cols)}")
    cab += novas_cols
    ix = {c: i for i, c in enumerate(cab)}
    for obrig in ("DELIVERED_DATE", "Transportadora", "NF", "DOCNAME", "SCI"):
        if obrig not in ix:
            raise RuntimeError(f"A BASE não tem a coluna {obrig}")
    i_fr, i_vl = ix[COL_VALOR_FRETE], ix[COL_VALOR_NF]
    g = lambda l, c: (l[ix[c]].strip() if ix[c] < len(l) else "")

    # 1) linhas que precisam de valor, dentro do período (mais recentes primeiro)
    alvo = []
    for n_lin, l in enumerate(todas[1:], start=2):
        d = data_iso(g(l, "DELIVERED_DATE"))
        if not d or (desde and d < desde) or (ate and d > ate):
            continue
        car = _carrier_base(g(l, "Transportadora"))
        fr = l[i_fr].strip() if i_fr < len(l) else ""
        vl = l[i_vl].strip() if i_vl < len(l) else ""
        precisa_fr = car in ("JAMEF", "FITLOG", "Mira") and not fr
        precisa_vl = not vl
        if not (precisa_fr or precisa_vl):
            continue
        alvo.append({"linha": n_lin, "data": d, "car": car, "nf": nf_numero(g(l, "NF")), "doc": g(l, "DOCNAME"),
                     "sci": g(l, "SCI"), "fr": fr, "vl": vl, "precisa_fr": precisa_fr, "precisa_vl": precisa_vl})
    alvo.sort(key=lambda a: a["data"], reverse=True)
    periodo = f"{desde.strftime('%d/%m/%Y') if desde else 'início'} a {ate.strftime('%d/%m/%Y') if ate else 'hoje'}"
    log(f"{len(alvo)} linha(s) da BASE sem valor no período {periodo}")
    if not alvo:
        return

    # 2) relatórios Jamef e Mira (uma leitura só)
    rel = valores_relatorios(gc)

    # 3) para cada linha: frete/valor pela fonte da transportadora
    senha_fitlog = os.environ.get("FITLOG_SENHA_XML") or SSW_SENHAS["FITLOG"]
    xml_cache, sc_cache, inv_cache = {}, {}, {}
    resultados, parou = [], False
    for k, a in enumerate(alvo):
        if not fim_ok():
            parou = True
            break
        frete = valor = None
        if a["car"] == "FITLOG" and a["nf"]:
            if a["nf"] not in xml_cache:
                try:
                    xml_cache[a["nf"]] = xml_cte_fitlog(a["nf"], senha_fitlog)
                except Exception:
                    xml_cache[a["nf"]] = None
                time.sleep(0.4)
            if xml_cache[a["nf"]]:
                frete, valor = xml_cache[a["nf"]]
        elif a["car"] in ("JAMEF", "Mira") and (a["car"], a["nf"]) in rel:
            frete, valor = rel[(a["car"], a["nf"])]
        # valor da NF ainda vazio (Inhouse, ou NF fora do relatório): pelo Zecore
        if valor is None and a["precisa_vl"] and zc is not None and a["doc"]:
            if a["doc"] not in sc_cache:
                sc = zc.shipping_core(a["doc"])
                sc_cache[a["doc"]] = {it.get("name"): str(it.get("sales_invoice") or "") for it in ((sc or {}).get("items") or [])}
                time.sleep(0.15)
            valor = total_da_invoice(zc, sc_cache[a["doc"]].get(a["sci"]), inv_cache)
        novo_fr = a["fr"] if a["fr"] or frete is None else round(frete, 2)
        novo_vl = a["vl"] if a["vl"] or valor is None else round(valor, 2)
        if (novo_fr != a["fr"]) or (novo_vl != a["vl"]):
            resultados.append((a["linha"], novo_fr, novo_vl, a["car"]))
        if (k + 1) % 200 == 0:
            log(f"  {k + 1}/{len(alvo)} linhas olhadas · {len(resultados)} com valor novo")

    cont = {}
    for _, fr, vl, car in resultados:
        c = cont.setdefault(car, [0, 0, 0])
        c[0] += 1
        c[1] += 1 if fr != "" else 0
        c[2] += 1 if vl != "" else 0
    log(f"{len(resultados)} linha(s) com valor novo — " + " · ".join(f"{k}: {v[0]} (frete {v[1]}, NF {v[2]})" for k, v in sorted(cont.items())))
    if parou:
        log(f"Tempo máximo de {tempo_max_min} min atingido em {k}/{len(alvo)} linhas — rode de novo para continuar de onde parou")
    if simular:
        log("Modo --simular: nada foi gravado")
        for r in resultados[:5]:
            log(f"  linha {r[0]}: frete {r[1]} · NF {r[2]}")
        return

    # 4) grava só as duas colunas, em lotes
    col_fr, col_vl = gu.rowcol_to_a1(1, i_fr + 1)[:-1], gu.rowcol_to_a1(1, i_vl + 1)[:-1]
    atualiz = []
    for linha, fr, vl, _ in resultados:
        atualiz.append({"range": f"{col_fr}{linha}", "values": [[fr]]})
        atualiz.append({"range": f"{col_vl}{linha}", "values": [[vl]]})
    for i in range(0, len(atualiz), 400):
        aba.batch_update(atualiz[i:i + 400], value_input_option="USER_ENTERED")
        time.sleep(1)
    log(f"Gravado: {len(resultados)} linha(s) da BASE com frete/valor da NF")


# ── Principal ───────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", help="dia de entrega (AAAA-MM-DD). Padrão: ontem")
    ap.add_argument("--csv", help="usa este CSV em vez de exportar do QuickSight")
    ap.add_argument("--simular", action="store_true", help="faz tudo, mas não grava na planilha")
    ap.add_argument("--sem-rastreio", action="store_true", help="não consulta as transportadoras")
    ap.add_argument("--sem-valores", action="store_true", help="não busca frete e valor da NF")
    ap.add_argument("--sem-coordenadas", action="store_true", help="não localiza os endereços no mapa")
    ap.add_argument("--completar-coordenadas", action="store_true",
                    help="preenche LAT/LNG nas linhas antigas da BASE que têm endereço e estão sem coordenada")
    ap.add_argument("--completar-valores", action="store_true",
                    help="preenche VALOR FRETE/VALOR NF nas linhas antigas da BASE que estão em branco")
    ap.add_argument("--desde", help="com --completar-valores: entregas a partir de AAAA-MM-DD (padrão: 90 dias)")
    ap.add_argument("--ate", help="com --completar-valores: entregas até AAAA-MM-DD")
    ap.add_argument("--tempo-max", type=int, default=25, help="com --completar-valores: minutos máximos (padrão 25)")
    ap.add_argument("--headed", action="store_true", help="mostra o navegador")
    a = ap.parse_args()

    hoje = dt.datetime.now(dt.timezone(dt.timedelta(hours=-3))).date()   # horário de Brasília

    if a.completar_coordenadas:
        creds_json = os.environ.get("GOOGLE_CREDS_JSON")
        if not creds_json:
            sys.exit("Falta a variável GOOGLE_CREDS_JSON")
        desde = dt.date.fromisoformat(a.desde) if a.desde else hoje - dt.timedelta(days=90)
        ate = dt.date.fromisoformat(a.ate) if a.ate else None
        log("Modo completar coordenadas")
        completar_coordenadas(cliente_google(creds_json), os.environ.get("ENTREGUES_SHEET_ID", SHEET_ID_PADRAO), desde, ate, a.tempo_max, a.simular)
        log("Concluído")
        return

    if a.completar_valores:
        creds_json = os.environ.get("GOOGLE_CREDS_JSON")
        if not creds_json:
            sys.exit("Falta a variável GOOGLE_CREDS_JSON")
        zu, zp = os.environ.get("ZECORE_USER"), os.environ.get("ZECORE_PASSWORD")
        zc = Zecore(zu, zp, a.headed) if zu and zp else None
        if zc is None:
            log("Sem ZECORE_USER/ZECORE_PASSWORD — valor da NF do Inhouse fica em branco")
        desde = dt.date.fromisoformat(a.desde) if a.desde else hoje - dt.timedelta(days=90)
        ate = dt.date.fromisoformat(a.ate) if a.ate else None
        log("Modo completar valores")
        completar_valores(cliente_google(creds_json), zc, os.environ.get("ENTREGUES_SHEET_ID", SHEET_ID_PADRAO),
                          desde, ate, a.tempo_max, a.simular)
        log("Concluído")
        return
    dia = dt.date.fromisoformat(a.data) if a.data else hoje - dt.timedelta(days=1)
    log(f"Pedidos entregues em {dia.strftime('%d/%m/%Y')}")

    creds_json = os.environ.get("GOOGLE_CREDS_JSON")
    if not creds_json and not a.simular:
        sys.exit("Falta a variável GOOGLE_CREDS_JSON")

    caminho = Path(a.csv) if a.csv else exportar_quicksight(dia, a.headed)
    if caminho is None:
        log("Nenhum item entregue nesse dia. Nada a gravar.")
        return
    itens = ler_csv(caminho, dia)
    if not itens:
        log("Nenhum item entregue nesse dia. Nada a gravar.")
        return

    zu, zp = os.environ.get("ZECORE_USER"), os.environ.get("ZECORE_PASSWORD")
    if not zu or not zp:
        sys.exit("Faltam as variáveis ZECORE_USER e ZECORE_PASSWORD")
    zc = Zecore(zu, zp, a.headed)
    por_sci = dados_zecore(itens, zc)

    rastreio = {}
    if not a.sem_rastreio:
        ju, jp = os.environ.get("JAMEF_API_USER"), os.environ.get("JAMEF_API_PASS")
        jamef = None
        if ju and jp:
            try:
                jamef = Jamef(ju, jp)
            except Exception as e:
                log(f"Jamef API indisponível ({e}) — datas da Jamef ficam em branco")
        else:
            log("Sem JAMEF_API_USER/JAMEF_API_PASS — datas da Jamef ficam em branco")
        rastreio = rastrear(itens, por_sci, jamef)

    gc = cliente_google(creds_json) if creds_json else None
    valores = {} if a.sem_valores else buscar_valores(itens, por_sci, gc, zc)

    linhas, sem_zecore, cont = [], [], {}
    for r in itens:
        z = por_sci.get(r["SCI"]) or {}
        if not z:
            sem_zecore.append(f"{r['DOCNAME']}/{r['SCI']}")
        c = cont.setdefault(r["_carrier"], [0, 0])
        c[0] += 1
        c[1] += 1 if z.get("nf") else 0
        linhas.append(montar_linha(r, z, rastreio.get((r["_carrier"], z.get("nf", ""))) or {}, valores.get(r["SCI"], (None, None))))
    log("NF preenchida por transportadora: " + " · ".join(f"{k} {v[1]}/{v[0]}" for k, v in sorted(cont.items())))
    if sem_zecore:
        log(f"Atenção: {len(sem_zecore)} item(ns) sem dados do Zecore (NF/endereço em branco): {', '.join(sem_zecore[:5])}{'…' if len(sem_zecore) > 5 else ''}")

    if not a.sem_coordenadas:
        geocodificar_linhas(linhas)
    if gc is None:
        log("Sem GOOGLE_CREDS_JSON: nada foi gravado (simulação local)")
    else:
        gravar_base(linhas, os.environ.get("ENTREGUES_SHEET_ID", SHEET_ID_PADRAO), gc, a.simular)
    log("Concluído")


if __name__ == "__main__":
    main()

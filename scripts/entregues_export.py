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

    for i, nome in ((0, "início"), (1, "término")):
        campo = controle.locator(f"input[data-automation-id='date_picker_{i}']").first
        if campo.count() == 0:
            qs.salvar_debug(page, "campos_data_nao_encontrados")
            raise RuntimeError(f"Não achei o campo de data de {nome} do controle '{CONTROLE_DATA}'")
        campo.click()
        campo.press("Control+A")
        campo.press("Backspace")
        campo.type(texto, delay=60)
        campo.press("Enter")
        page.wait_for_timeout(1500)
        valor = (campo.input_value() or "").strip()
        if valor != texto:
            # 2ª tentativa com fill (o campo é um input de texto do Material UI)
            campo.fill(texto)
            campo.press("Tab")
            page.wait_for_timeout(1500)
            valor = (campo.input_value() or "").strip()
        log(f"  data de {nome}: {valor}")
        if valor != texto:
            qs.salvar_debug(page, "data_nao_aceita")
            raise RuntimeError(f"O campo de data de {nome} ficou com '{valor}' em vez de '{texto}'")

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
    log(f"Procurando a tabela '{TITULO_VISUAL}'")
    titulo = None
    for _ in range(25):
        cand = page.get_by_text(TITULO_VISUAL, exact=False)
        for i in range(min(cand.count(), 5)):
            el = cand.nth(i)
            try:
                if el.is_visible(timeout=800):
                    el.scroll_into_view_if_needed(timeout=5000)
                    page.wait_for_timeout(2000)
                    titulo = el
                    break
            except Exception:
                continue
        if titulo:
            break
        page.mouse.wheel(0, 700)
        page.wait_for_timeout(900)
    if not titulo:
        qs.salvar_debug(page, "tabela_raw_data_nao_encontrada")
        raise RuntimeError(f"Não achei a tabela '{TITULO_VISUAL}'")

    if not qs.abrir_menu_visual(page, titulo):
        qs.salvar_debug(page, "menu_raw_data")
        raise RuntimeError("Não consegui abrir o menu da tabela RAW DATA")

    with page.expect_download(timeout=TIMEOUT_DOWNLOAD_MS) as info:
        if not qs.clicar_exportar_csv(page):
            qs.salvar_debug(page, "exportar_raw_data")
            raise RuntimeError("Menu abriu, mas não achei 'Exportar para CSV'")
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
            end = re.sub(r"-Shipping-\d+$", "", str(it.get("address") or "").strip())
            por_sci[it.get("name")] = {"criacao": criacao, "nf": nf, "endereco": end}
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


def montar_linha(r, z, tr):
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
    }


# ── 5. Gravação na BASE ─────────────────────────────────────────

def gravar_base(linhas, sheet_id, creds_json, simular=False):
    import gspread
    from google.oauth2.service_account import Credentials

    creds = Credentials.from_service_account_info(json.loads(creds_json), scopes=["https://www.googleapis.com/auth/spreadsheets"])
    aba = gspread.authorize(creds).open_by_key(sheet_id).worksheet(ABA_BASE)
    cab = [c.strip() for c in aba.row_values(1)]
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


# ── Principal ───────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", help="dia de entrega (AAAA-MM-DD). Padrão: ontem")
    ap.add_argument("--csv", help="usa este CSV em vez de exportar do QuickSight")
    ap.add_argument("--simular", action="store_true", help="faz tudo, mas não grava na planilha")
    ap.add_argument("--sem-rastreio", action="store_true", help="não consulta as transportadoras")
    ap.add_argument("--headed", action="store_true", help="mostra o navegador")
    a = ap.parse_args()

    hoje = dt.datetime.now(dt.timezone(dt.timedelta(hours=-3))).date()   # horário de Brasília
    dia = dt.date.fromisoformat(a.data) if a.data else hoje - dt.timedelta(days=1)
    log(f"Pedidos entregues em {dia.strftime('%d/%m/%Y')}")

    creds_json = os.environ.get("GOOGLE_CREDS_JSON")
    if not creds_json and not a.simular:
        sys.exit("Falta a variável GOOGLE_CREDS_JSON")

    caminho = Path(a.csv) if a.csv else exportar_quicksight(dia, a.headed)
    itens = ler_csv(caminho, dia)
    if not itens:
        log("Nenhum item entregue nesse dia. Nada a gravar.")
        return

    zu, zp = os.environ.get("ZECORE_USER"), os.environ.get("ZECORE_PASSWORD")
    if not zu or not zp:
        sys.exit("Faltam as variáveis ZECORE_USER e ZECORE_PASSWORD")
    por_sci = dados_zecore(itens, Zecore(zu, zp, a.headed))

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

    linhas, sem_zecore = [], 0
    for r in itens:
        z = por_sci.get(r["SCI"]) or {}
        if not z:
            sem_zecore += 1
        linhas.append(montar_linha(r, z, rastreio.get((r["_carrier"], z.get("nf", ""))) or {}))
    if sem_zecore:
        log(f"Atenção: {sem_zecore} item(ns) sem dados do Zecore (NF/endereço em branco)")

    gravar_base(linhas, os.environ.get("ENTREGUES_SHEET_ID", SHEET_ID_PADRAO), creds_json, a.simular)
    log("Concluído")


if __name__ == "__main__":
    main()

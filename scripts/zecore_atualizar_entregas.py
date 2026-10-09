#!/usr/bin/env python3
"""
Atualiza no Zecore os pedidos que já foram entregues pela transportadora.

O PROBLEMA QUE ISSO RESOLVE
O Zecore carimba a data em que alguém clica, não a data real da entrega.
Como a atualização é manual, uma entrega das 23h só era registrada no dia
seguinte — e virava atraso no SLA sem ter sido. Rodando de forma
automática e frequente, a defasagem cai de dias para minutos.

O FLUXO (o mesmo que era feito na mão)
  1. Lê na Dash_fedex quem está entregue na transportadora mas ainda
     "Delivered to Carrier" no Zecore (só Fitlog e Mira — Jamef tem API)
  2. Abre shipping-core do pedido pela URL, direto
  3. Acha o item com Shipping State = "Delivered to Carrier"
  4. Abre a linha e copia o Master Guide + o link do Arrangement
  5. No Arrangement: Create → Delivery
  6. Na Delivery Task: apaga todas as linhas MENOS a do Master Guide
  7. Confere que sobrou exatamente 1 linha, e que é a certa
  8. Submit

SEGURANÇA
O Submit é irreversível, então ele só acontece depois de duas
verificações. Até lá tudo é rascunho e não afeta nada.

  --reconhecer  Navega e tira print de cada tela. NÃO MODIFICA NADA.
  --simular     Faz tudo menos o Submit. Deixa o rascunho de lado.
  --executar    Faz tudo, incluindo o Submit.

Sem nenhum desses, roda em --simular (o modo seguro é o padrão).

Variáveis de ambiente:
  ZECORE_USER, ZECORE_PASSWORD
  GOOGLE_CREDS_JSON, SHEET_ID (opcional)
"""

import argparse
import json
import os
import re
import sys
import time
import urllib.request
from datetime import datetime, timedelta, timezone
from pathlib import Path

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
except ImportError:
    # O modo --verificar roda antes de o navegador ser instalado
    sync_playwright = None

    class PWTimeout(Exception):
        pass

# ── Configuração ────────────────────────────────────────────────

ZECORE_BASE = "https://zecore.zebrands.mx"
SHEET_ID_PADRAO = "1RMM3Niaf3E9WVCN3NdfU1Cb4Fuz4VSJjaLUPBdrFldI"
ABA_RASTR = "Rastreamento"

# Colunas da aba Rastreamento (1 = A)
COL_NF            = 1    # A
COL_STATUS        = 2    # B  status da transportadora
COL_SALES_ORDER   = 5    # E
COL_STATUS_ZECORE = 10   # J
COL_CARRIER       = 11   # K

# Status da coluna B que contam como entregue. A comparação ignora
# acento, maiúscula e hífen: "MERCADORIA PRÉ-ENTREGUE (MOBILE)" casa com
# "PRE ENTREGUE". Pra acrescentar outro, é só incluir na lista.
STATUS_ENTREGUE = ["MERCADORIA ENTREGUE", "PRE ENTREGUE"]


def _norm_status(s):
    import unicodedata
    s = unicodedata.normalize("NFD", str(s or ""))
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", s.upper().replace("-", " ")).strip()


def _status_conta_como_entregue(status):
    st = _norm_status(status)
    return any(_norm_status(alvo) in st for alvo in STATUS_ENTREGUE)
STATUS_ZECORE_PENDENTE = "DELIVERED TO CARRIER"
CARRIERS = ["FITLOG", "MIRA"]          # Jamef fica de fora (tem API)

LIMITE_PADRAO = 1                      # local começa conservador; no GitHub roda com --limite 0 (sem limite)

# Não limita baixas: só interrompe a rodada quando vários pedidos SEGUIDOS
# dão problema — sinal de que algo mudou na tela do Zecore e todos vão
# falhar igual, deixando um monte de rascunho pendurado.
MAX_PROBLEMAS_SEGUIDOS = 3

TIMEOUT_MS = 45_000
DIR_DEBUG = Path("debug")
DIR_RECON = Path("reconhecimento")

# Histórico de baixas — fica no próprio repositório (o workflow faz commit
# dele no fim de cada execução). É a memória do robô: quem está aqui já
# teve baixa e é pulado sem nem abrir o Zecore.
ARQ_HISTORICO = Path("dados/baixas_realizadas.json")
DIAS_MANTER_HISTORICO = 90

# Horário de Brasília sem depender de base de fusos (o Windows não traz)
FUSO_BR = timezone(timedelta(hours=-3))


def agora_br():
    return datetime.now(FUSO_BR)


# ── NF ──────────────────────────────────────────────────────────

def _nf_normalizada(nf):
    """'115749-1' → '115749'. Vazio se não houver número."""
    s = re.sub(r"-\d+$", "", str(nf or "").strip())
    dig = "".join(ch for ch in s if ch.isdigit())
    return str(int(dig)) if dig else ""


def _nf_do_master_guide(master_guide):
    """O Master Guide é a chave da NF-e (44 dígitos), e o número da nota
    fica nas posições 26 a 34. Isso permite conferir se o item é mesmo da
    NF da planilha antes de dar baixa."""
    d = "".join(ch for ch in str(master_guide or "") if ch.isdigit())
    if len(d) != 44:
        return ""
    return str(int(d[25:34]))


def _chave(sales_order, nf):
    return f"{sales_order}|{nf}"


def _carrier_curto(texto):
    t = str(texto or "").upper()
    if "FITLOG" in t:
        return "Fitlog"
    if "MIRA" in t:
        return "Mira"
    return texto or "?"


# ── Histórico ───────────────────────────────────────────────────

def carregar_historico():
    if not ARQ_HISTORICO.exists():
        return {"baixas": {}}
    try:
        dados = json.loads(ARQ_HISTORICO.read_text(encoding="utf-8"))
        dados.setdefault("baixas", {})
        return dados
    except Exception as e:
        log(f"Histórico ilegível ({e}) — começando um novo")
        return {"baixas": {}}


def salvar_historico(hist):
    """Grava o histórico, descartando registros com mais de 90 dias."""
    limite = agora_br() - timedelta(days=DIAS_MANTER_HISTORICO)
    mantidos = {}
    for chave, reg in hist.get("baixas", {}).items():
        try:
            if datetime.fromisoformat(reg["data"]) >= limite:
                mantidos[chave] = reg
        except Exception:
            mantidos[chave] = reg
    hist["baixas"] = mantidos
    hist["atualizado_em"] = agora_br().isoformat(timespec="seconds")

    ARQ_HISTORICO.parent.mkdir(parents=True, exist_ok=True)
    ARQ_HISTORICO.write_text(
        json.dumps(hist, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def registrar_baixa(hist, pedido, r):
    nf = _nf_do_master_guide(r.get("master_guide")) or _nf_normalizada(pedido.get("nf"))
    hist["baixas"][_chave(pedido["sales_order"], nf)] = {
        "data": agora_br().isoformat(timespec="seconds"),
        "sales_order": pedido["sales_order"],
        "nf": nf,
        "carrier": _carrier_curto(
            pedido.get("carrier") if pedido.get("carrier") not in (None, "?")
            else r.get("arrangement")
        ),
        "status_transportadora": pedido.get("status"),
        "master_guide": r.get("master_guide"),
        "arrangement": r.get("arrangement"),
        "itens": r.get("mantidas"),
        "task": r.get("task"),
    }


def registrar_ja_baixado(hist, pedido, estados):
    """Guarda no histórico um pedido que já estava entregue no Zecore
    (baixa manual, por exemplo), pra não reabrir ele toda hora.

    Só grava se TODOS os itens estão 'Delivered', que é estado final.
    Qualquer outro estado (devolução etc.) pode mudar depois, então
    nesses casos o pedido continua sendo olhado nas próximas rodadas.
    """
    if not estados or any(e.strip().upper() != "DELIVERED" for e in estados):
        return False
    nf = _nf_normalizada(pedido.get("nf"))
    if not nf:
        return False
    hist["baixas"][_chave(pedido["sales_order"], nf)] = {
        "data": agora_br().isoformat(timespec="seconds"),
        "sales_order": pedido["sales_order"],
        "nf": nf,
        "carrier": _carrier_curto(pedido.get("carrier")),
        "origem": "ja_tinha_baixa",
    }
    return True


# ── Fila de remoção na planilha ─────────────────────────────────
# O robô NÃO apaga linhas da Rastreamento: o rastreio do Apps Script grava
# status pelo número da linha durante vários minutos, e uma linha apagada
# por fora no meio disso desloca todas as de baixo — o status iria pra
# linha errada. Em vez disso, ele anota o pedido na aba "Baixas Zecore",
# e o próprio Apps Script apaga, dentro da trava dele, no começo da
# rodada seguinte. Acrescentar linha em OUTRA aba não desloca nada.
ABA_FILA_BAIXAS = "Baixas Zecore"
CAB_FILA = ["Data/hora", "NF", "Sales Order", "Transportadora", "Origem",
            "Removido da Rastreamento em"]


class FilaBaixas:
    def __init__(self, sheet_id, creds_json):
        self.ws = None
        self.nfs = set()
        try:
            import gspread
            from google.oauth2.service_account import Credentials
            creds = Credentials.from_service_account_info(
                json.loads(creds_json),
                scopes=["https://www.googleapis.com/auth/spreadsheets"],
            )
            planilha = gspread.authorize(creds).open_by_key(sheet_id)
            try:
                self.ws = planilha.worksheet(ABA_FILA_BAIXAS)
            except gspread.WorksheetNotFound:
                self.ws = planilha.add_worksheet(title=ABA_FILA_BAIXAS, rows=1000,
                                                 cols=len(CAB_FILA))
                self.ws.update(values=[CAB_FILA], range_name="A1")
                log(f'Aba "{ABA_FILA_BAIXAS}" criada')
            for v in self.ws.col_values(2)[1:]:
                n = _nf_normalizada(v)
                if n:
                    self.nfs.add(n)
        except Exception as e:
            log(f'Não consegui abrir a aba "{ABA_FILA_BAIXAS}" ({e}) — '
                "as linhas não serão removidas da Rastreamento nesta rodada")
            self.ws = None

    def adicionar(self, itens):
        """itens: lista de (pedido, nf, origem). Ignora NF que já está na fila."""
        if not self.ws:
            return 0
        novos = []
        for pedido, nf, origem in itens:
            n = _nf_normalizada(nf)
            if not n or n in self.nfs:
                continue
            self.nfs.add(n)
            novos.append([
                agora_br().strftime("%d/%m/%Y %H:%M"), str(nf),
                pedido.get("sales_order", ""), _carrier_curto(pedido.get("carrier")),
                origem, "",
            ])
        if not novos:
            return 0
        try:
            # RAW pra o Sheets não transformar "115879-1" em data
            self.ws.append_rows(novos, value_input_option="RAW", table_range="A1")
            return len(novos)
        except Exception as e:
            log(f'Não consegui anotar na aba "{ABA_FILA_BAIXAS}": {e}')
            return 0


def enfileirar_do_historico(fila, todos, hist):
    """Pedidos que estão na planilha e que o ROBÔ já baixou numa rodada
    anterior também têm que sair da Rastreamento. Os que ele só encontrou
    já entregues (baixa manual) ficam de fora — só sai o que ele baixou."""
    ja = []
    for pd in todos:
        reg = hist["baixas"].get(_chave(pd["sales_order"], _nf_normalizada(pd["nf"])))
        if reg is not None and reg.get("origem") != "ja_tinha_baixa":
            ja.append((pd, pd["nf"], "baixa pelo robô"))
    n = fila.adicionar(ja) if ja else 0
    if n:
        log(f"{n} pedido(s) baixado(s) em rodada anterior anotado(s) pra sair da Rastreamento")
    return n


# ── Aviso no Chat ───────────────────────────────────────────────

def avisar_chat(webhook, baixas, problemas, ja_tinham_baixa, interrompida=False):
    """Manda o resumo da rodada pro espaço de baixas. Só envia se houve
    baixa ou problema — rodada sem nada a fazer fica em silêncio."""
    if not webhook:
        log("CHAT_WEBHOOK_BAIXAS não configurado — sem aviso no Chat")
        return
    if not baixas and not problemas:
        log("Nada a reportar no Chat")
        return

    linhas = [f"📦 *Baixas no Zecore — {agora_br().strftime('%d/%m %H:%M')}*", ""]

    if baixas:
        linhas.append(f"✅ *{len(baixas)} pedido(s) com baixa:*")
        for b in baixas:
            linhas.append(f"• NF {b['nf']} · {b['sales_order']} · {b['carrier']}"
                          + (" · _pré-entregue_" if b.get("pre") else ""))
        linhas.append("")

    if problemas:
        linhas.append(f"⚠️ *{len(problemas)} com problema — nada foi alterado neles:*")
        for pb in problemas:
            linhas.append(f"• {pb['sales_order']} — {pb['motivo'][:160]}")
        linhas.append("")

    if interrompida:
        linhas.append(f"🛑 *Rodada interrompida* depois de {MAX_PROBLEMAS_SEGUIDOS} problemas "
                      "seguidos — pode ser mudança na tela do Zecore. Os demais pedidos "
                      "ficam pra próxima rodada.")
        linhas.append("")

    if ja_tinham_baixa:
        linhas.append(f"_{ja_tinham_baixa} já estavam com baixa no Zecore e foram pulados._")

    if baixas:
        linhas.append("_As linhas dos pedidos com baixa saem da Rastreamento na próxima "
                      "rodada do rastreio (até 20 min). Os com problema continuam lá._")

    corpo = json.dumps({"text": "\n".join(linhas).strip()}).encode("utf-8")
    req = urllib.request.Request(
        webhook, data=corpo, method="POST",
        headers={"Content-Type": "application/json; charset=UTF-8"},
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            log(f"Resumo enviado ao Chat (HTTP {resp.status})")
    except Exception as e:
        log(f"Não consegui enviar ao Chat: {e}")


def log(msg):
    print(f"[{agora_br().strftime('%H:%M:%S')}] {msg}", flush=True)


def salvar_debug(page, nome, pasta=None):
    pasta = pasta or DIR_DEBUG
    pasta.mkdir(exist_ok=True)
    try:
        page.screenshot(path=str(pasta / f"{nome}.png"), full_page=True)
        (pasta / f"{nome}.html").write_text(page.content(), encoding="utf-8")
        log(f"Salvo: {pasta.name}/{nome}.png e .html")
    except Exception as e:
        log(f"Não consegui salvar {nome}: {e}")


# ── Lista de pedidos pendentes ──────────────────────────────────

def ler_pendentes(sheet_id, creds_json, limite):
    import gspread
    from google.oauth2.service_account import Credentials

    log("Lendo a Dash_fedex")
    creds = Credentials.from_service_account_info(
        json.loads(creds_json),
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
    )
    gc = gspread.authorize(creds)
    aba = gc.open_by_key(sheet_id).worksheet(ABA_RASTR)

    linhas = aba.get_all_values()
    pendentes = []

    for i, linha in enumerate(linhas[1:], start=2):
        def col(n):
            return linha[n - 1].strip() if len(linha) >= n else ""

        status = col(COL_STATUS).upper()
        carrier = col(COL_CARRIER).upper()
        zecore = col(COL_STATUS_ZECORE).upper()
        sales_order = col(COL_SALES_ORDER)

        if not _status_conta_como_entregue(status):
            continue
        if not any(c in carrier for c in CARRIERS):
            continue
        if STATUS_ZECORE_PENDENTE not in zecore:
            continue
        if not sales_order:
            continue

        pendentes.append({
            "linha": i,
            "nf": col(COL_NF),
            "sales_order": sales_order,
            "carrier": carrier,
            "status": col(COL_STATUS),
        })

    log(f"{len(pendentes)} pedido(s) pendente(s) de atualização na planilha")
    return pendentes


# ── Login ───────────────────────────────────────────────────────

def fazer_login(page, usuario, senha):
    log("Fazendo login no Zecore")
    page.goto(f"{ZECORE_BASE}/login", wait_until="domcontentloaded", timeout=TIMEOUT_MS)

    campo_user = page.locator("#login_email, input[name='email'], input[type='email']").first
    campo_user.wait_for(state="visible", timeout=TIMEOUT_MS)
    campo_user.click()
    campo_user.fill(usuario)

    campo_senha = page.locator("#login_password, input[type='password']").first
    campo_senha.click()
    campo_senha.fill(senha)

    page.keyboard.press("Enter")
    page.wait_for_url("**/app**", timeout=TIMEOUT_MS)
    page.wait_for_load_state("networkidle", timeout=TIMEOUT_MS)
    log("Login concluído")


# ── Etapa: abrir o pedido e pegar Master Guide + Arrangement ────

def abrir_pedido(page, sales_order):
    url = f"{ZECORE_BASE}/app/shipping-core/Sales%20Order-{sales_order}-shipping"
    log(f"Abrindo {sales_order}")
    page.goto(url, wait_until="domcontentloaded", timeout=TIMEOUT_MS)
    page.wait_for_timeout(4000)
    try:
        page.wait_for_load_state("networkidle", timeout=20_000)
    except PWTimeout:
        pass


def localizar_itens_pendentes(page, espera_s=20):
    """Lê o Shipping State de cada item do pedido.

    Devolve (pendentes, estados_vistos):
      pendentes      → linhas ainda em 'Delivered to Carrier'
      estados_vistos → o estado de TODOS os itens, pra explicar no log

    A grade do Frappe às vezes termina de desenhar depois que a página
    "carregou". Por isso espera até aparecer pelo menos um item com
    Shipping State, em vez de olhar uma vez só — senão um pedido
    pendente pode ser pulado como se já tivesse baixa.
    """
    fim = time.time() + espera_s
    while True:
        pendentes, estados = [], []
        linhas = page.locator(".grid-row[data-idx]")

        for i in range(linhas.count()):
            linha = linhas.nth(i)
            try:
                estado = linha.locator("[data-fieldname='shipping_state'] .static-area")
                if estado.count() == 0:
                    continue
                texto = (estado.first.inner_text(timeout=2000) or "").strip()
                if not texto:
                    continue
                estados.append(texto)
                if texto.upper() == STATUS_ZECORE_PENDENTE:
                    pendentes.append((linha, i + 1))
            except Exception:
                continue

        if estados or time.time() >= fim:
            return pendentes, estados
        page.wait_for_timeout(1000)


def _fechar_item(page):
    """Fecha o painel de edição de uma linha (atalho ESC do Frappe)."""
    page.keyboard.press("Escape")
    page.wait_for_timeout(1200)


def _link_do_item(linha, doctype):
    """Procura o link SÓ dentro da linha do item. No Frappe o formulário
    de edição abre dentro da própria linha, então tudo que é daquele item
    está ali dentro."""
    try:
        el = linha.locator(f"a[data-doctype='{doctype}']")
        if el.count() > 0:
            return (el.first.get_attribute("data-name") or "").strip(), el.first
    except Exception:
        pass
    return "", None


def extrair_dados_item(page, linha):
    """Abre a linha e lê o Master Guide e o Arrangement DAQUELE item.

    Antes a busca era na página inteira, pegando o primeiro link que
    aparecesse. Só que a tabela de itens também mostra o Master Guide de
    cada linha como link — então, num pedido com várias NFs, abrir o
    terceiro item e ler "o primeiro da página" devolvia o Master Guide do
    PRIMEIRO item. O robô achava que nenhum item era da NF procurada.
    """
    # Master Guide da célula da própria linha, antes de abrir
    mg_linha, _ = _link_do_item(linha, "Shipping Master Tracker")

    linha.locator(".btn-open-row").first.click()
    page.wait_for_timeout(2500)

    # Depois de aberto: formulário dentro da mesma linha
    mg_form, _ = _link_do_item(linha, "Shipping Master Tracker")
    if not mg_form:
        try:
            el = page.locator(".grid-row-open a[data-doctype='Shipping Master Tracker']")
            if el.count() > 0:
                mg_form = (el.first.get_attribute("data-name") or "").strip()
        except Exception:
            pass

    master_guide = mg_linha or mg_form
    if not master_guide:
        raise RuntimeError("Não achei o Master Guide deste item")
    if mg_linha and mg_form and mg_linha != mg_form:
        raise RuntimeError(
            f"Master Guide da linha ({mg_linha}) diferente do formulário ({mg_form}) — "
            "não vou arriscar"
        )

    # Arrangement: só aparece no formulário aberto
    arrangement_nome, arranjo = _link_do_item(linha, "Arrangement")
    if arranjo is None:
        try:
            el = page.locator(".grid-row-open a[data-doctype='Arrangement']")
            if el.count() == 0:
                # Último recurso: na página toda, mas só se houver UM
                # (com um único item aberto, é o dele)
                el = page.locator("a[data-doctype='Arrangement']")
                if el.count() != 1:
                    el = None
            if el is not None and el.count() > 0:
                arranjo = el.first
                arrangement_nome = (arranjo.get_attribute("data-name") or "").strip()
        except Exception:
            arranjo = None
    if arranjo is None or not arrangement_nome:
        raise RuntimeError("Não achei o Arrangement deste item")

    arrangement_href = arranjo.get_attribute("href") or ""

    return {
        "master_guide": master_guide,
        "arrangement": arrangement_nome,
        "arrangement_url": ZECORE_BASE + arrangement_href,
    }


# ── Etapa: criar a Delivery Task ────────────────────────────────

def criar_delivery_task(page, arrangement_url):
    log("Abrindo o Arrangement")
    page.goto(arrangement_url, wait_until="domcontentloaded", timeout=TIMEOUT_MS)
    page.wait_for_timeout(4000)

    log("Create → Delivery")
    botao_create = page.locator("button[data-toggle='dropdown']:has-text('Create')").first
    botao_create.wait_for(state="visible", timeout=TIMEOUT_MS)
    botao_create.click()
    page.wait_for_timeout(1200)

    opcao = page.locator("a.dropdown-item[data-label='Delivery']").first
    opcao.wait_for(state="visible", timeout=TIMEOUT_MS)
    opcao.click()

    page.wait_for_timeout(5000)
    try:
        page.wait_for_load_state("networkidle", timeout=20_000)
    except PWTimeout:
        pass

    log(f"Delivery Task aberta: {page.url}")

    # Espera as linhas da grade de Items aparecerem antes de seguir
    for _ in range(15):
        try:
            linhas = _grade_items(page).locator(".grid-row[data-idx]")
            if linhas.count() > 0:
                log(f"Grade de Items pronta ({linhas.count()} linha(s))")
                page.wait_for_timeout(1500)
                return
        except Exception:
            pass
        page.wait_for_timeout(1000)

    log("A grade demorou a ficar pronta — seguindo mesmo assim")


# ── Etapa: deixar só a linha certa ──────────────────────────────

def _grade_items(page):
    """Devolve o container da grade de Items.

    A página tem duas grades com a mesma estrutura: Items (editável) e
    Resumed Task (somente leitura). Pegar a errada faz tudo falhar com
    "element is not visible", porque a de baixo tem os checkboxes
    desabilitados.

    Distinguimos de duas formas, e as duas têm que concordar:
      1. O campo: a de Items é data-fieldname="items"
      2. O cabeçalho: a de Items diz "Tracker Code"; a Resumed Task
         diz "Tracker Code/LPN"
    """
    grade = page.locator("[data-fieldname='items'] .form-grid").first

    if grade.count() > 0 and _grade_e_de_items(grade):
        return grade

    # Se o campo não bastou, procura pelo cabeçalho entre todas as grades
    todas = page.locator(".form-grid")
    for i in range(todas.count()):
        g = todas.nth(i)
        if _grade_e_de_items(g):
            return g

    raise RuntimeError(
        "Não consegui identificar a grade de Items (a de cima, editável)"
    )


def _grade_e_de_items(grade):
    """Confirma pelo cabeçalho: Items tem "Tracker Code",
    Resumed Task tem "Tracker Code/LPN"."""
    try:
        cab = grade.locator(
            ".grid-heading-row [data-fieldname='tracker_code'] .static-area"
        )
        if cab.count() == 0:
            return False
        texto = (cab.first.inner_text(timeout=2000) or "").strip()
        return texto == "Tracker Code"
    except Exception:
        return False


def _linhas_da_grade(page):
    """Devolve [(indice, tracker_code, locator)] da tabela Items."""
    resultado = []
    linhas = _grade_items(page).locator(".grid-row[data-idx]")

    for i in range(linhas.count()):
        linha = linhas.nth(i)
        try:
            campo = linha.locator("[data-fieldname='tracker_code'] .static-area")
            if campo.count() == 0:
                continue
            tracker = (campo.first.inner_text(timeout=2000) or "").strip()
            if tracker and tracker.lower() != "tracker code":
                resultado.append((i, tracker, linha))
        except Exception:
            continue

    return resultado


def _itens_do_documento(page):
    """Lê a tabela de itens direto do documento aberto (cur_frm.doc.items).

    É o dado real do Frappe, não o que está desenhado na tela — então não
    depende de grade renderizada, paginação ou checkbox habilitado.
    """
    return page.evaluate("""() => {
        if (typeof cur_frm === 'undefined' || !cur_frm || !cur_frm.doc) return null;
        return {
            doctype: cur_frm.doctype,
            name: cur_frm.docname,
            docstatus: cur_frm.doc.docstatus,
            itens: (cur_frm.doc.items || []).map(r => ({
                tracker: (r.tracker_code || '').trim(),
                item: r.item_code || '',
                idx: r.idx
            }))
        };
    }""")


def _aguardar_documento(page, tentativas=20):
    """Espera o cur_frm ser a Delivery Task e ter itens carregados."""
    for _ in range(tentativas):
        doc = _itens_do_documento(page)
        if doc and doc.get("doctype") == "Delivery Task" and doc.get("itens"):
            return doc
        page.wait_for_timeout(1000)
    return _itens_do_documento(page)


def apagar_outras_linhas(page, master_guide):
    """Deixa na Delivery Task só as linhas do Master Guide informado.

    Altera a tabela pela API de cliente do Frappe (cur_frm) em vez de
    clicar em checkboxes. Na sessão do robô o Frappe desenha os
    checkboxes da grade como desabilitados; a API é o que os botões da
    própria tela chamam por baixo, então é o caminho mais confiável.
    """
    doc = _aguardar_documento(page)

    if not doc:
        salvar_debug(page, "documento_nao_carregou")
        raise RuntimeError("Não consegui ler o documento aberto (cur_frm)")

    log(f"Documento: {doc['doctype']} {doc['name']} (docstatus {doc['docstatus']})")

    if doc["doctype"] != "Delivery Task":
        raise RuntimeError(
            f"O documento aberto é {doc['doctype']}, não Delivery Task. "
            "Não vou mexer."
        )

    if doc["docstatus"] != 0:
        raise RuntimeError(
            f"O documento não está em rascunho (docstatus {doc['docstatus']}). "
            "Não vou mexer."
        )

    itens = doc["itens"]
    alvos = [i for i in itens if i["tracker"] == master_guide]
    outras = [i for i in itens if i["tracker"] != master_guide]

    log(f"A Delivery Task tem {len(itens)} linha(s)")

    # ── Verificação 1: o Master Guide precisa estar na lista ──
    if not alvos:
        raise RuntimeError(
            f"O Master Guide {master_guide} não está nesta Delivery Task. "
            "Nada foi alterado."
        )

    log(f"{len(alvos)} linha(s) do Master Guide — serão mantidas")

    if not outras:
        log("Só tem as linhas certas — nada a remover")
        return None, len(alvos)

    log(f"Removendo {len(outras)} linha(s) de outros pedidos")

    resultado = page.evaluate("""(master) => {
        const antes = cur_frm.doc.items.length;
        const manter = cur_frm.doc.items.filter(
            r => (r.tracker_code || '').trim() === master
        );
        if (manter.length === 0) return { erro: 'nenhuma linha a manter' };

        manter.forEach((r, i) => { r.idx = i + 1; });
        cur_frm.doc.items = manter;
        cur_frm.refresh_field('items');
        cur_frm.dirty();

        return { antes: antes, depois: cur_frm.doc.items.length };
    }""", master_guide)

    if not resultado or resultado.get("erro"):
        salvar_debug(page, "remocao_falhou")
        raise RuntimeError(
            f"A remoção não funcionou: {resultado}. Nada foi salvo."
        )

    log(f"Tabela ajustada: {resultado['antes']} → {resultado['depois']} linha(s)")
    page.wait_for_timeout(1500)

    return None, len(alvos)


def conferir_antes_do_submit(page, master_guide, esperadas):
    """Última checagem antes do passo irreversível, lendo o documento.

    Toda linha que sobrou tem que ser do Master Guide esperado, e a
    quantidade tem que bater com a que contamos antes de remover.
    """
    doc = _itens_do_documento(page)
    if not doc:
        raise RuntimeError("Não consegui ler o documento. NÃO vou submeter.")

    itens = doc["itens"]

    if not itens:
        raise RuntimeError("Não sobrou nenhuma linha. NÃO vou submeter.")

    intrusas = [i["tracker"] for i in itens if i["tracker"] != master_guide]
    if intrusas:
        raise RuntimeError(
            f"Sobraram {len(intrusas)} linha(s) de outro Master Guide "
            f"({intrusas[:3]}). NÃO vou submeter."
        )

    if len(itens) != esperadas:
        raise RuntimeError(
            f"Sobraram {len(itens)} linha(s), esperava {esperadas}. "
            "NÃO vou submeter."
        )

    log(f"Confirmado: {len(itens)} linha(s), todas do Master Guide {master_guide}")
    return True


def salvar(page):
    """Depois de apagar linhas o documento fica "Not Saved" e o botão
    principal vira Save. Só depois de salvar o Submit reaparece.

    Save e Submit são o MESMO elemento (button.primary-action), mudando
    só o data-label. E o texto tem um <span> no meio ("S<span>a</span>ve"),
    então buscar por texto é frágil — usamos o data-label.
    """
    botao = page.locator("button.primary-action[data-label='Save']").first

    if botao.count() == 0 or not botao.is_visible():
        log("Não apareceu botão Save — o documento já deve estar salvo")
        return

    log("Salvando as alterações")
    botao.click()
    page.wait_for_timeout(3000)

    try:
        page.wait_for_load_state("networkidle", timeout=15_000)
    except PWTimeout:
        pass

    # Confirma que virou Submit — se continuar em Save, algo não salvou
    submit = page.locator("button.primary-action[data-label='Submit']").first
    if submit.count() == 0:
        salvar_debug(page, "save_nao_virou_submit")
        raise RuntimeError(
            "Depois de salvar o botão não virou Submit — veja debug/"
        )

    log("Salvo")


def _clicar_nao(nao, page):
    """Recusa a confirmação. Tenta o botão secundário do modal e, se não
    achar, fecha com Esc — o importante é NÃO confirmar."""
    try:
        if nao.count() > 0:
            nao.click(timeout=4000)
            page.wait_for_timeout(1000)
            return
    except Exception:
        pass
    page.keyboard.press("Escape")
    page.wait_for_timeout(1000)


def _yes_visivel(page, timeout_ms):
    """Devolve o botão Yes que está VISÍVEL agora, ou None.

    O Frappe não remove as janelas de confirmação da página — só esconde.
    Depois da primeira, o DOM tem dois botões Yes: o da janela antiga
    (escondido) e o da nova. Pegar o primeiro da lista pegava o escondido.
    Por isso procuramos entre todos o que está de fato visível.
    """
    fim = time.time() + timeout_ms / 1000.0
    while time.time() < fim:
        botoes = page.locator("button.btn-modal-primary").filter(has_text="Yes")
        for i in range(botoes.count()):
            b = botoes.nth(i)
            try:
                if b.is_visible():
                    return b
            except Exception:
                continue
        page.wait_for_timeout(300)
    return None


def submeter(page, esperadas):
    """Submete a Delivery Task — o passo irreversível.

    Depois do Submit aparecem duas confirmações:
        "Permanently Submit 4c3df7f3b0?"          [No] [Yes]
        "1 items to be marked as DELIVERED"       [No] [Yes]

    Clica Yes nas duas sem conferir o conteúdo: a conferência de verdade
    já foi feita antes, lendo o documento depois do Save — quando ele é
    recarregado do servidor. Se sobrou a linha certa ali, estas janelas
    refletem isso.

    A única exceção é uma janela que não seja nenhuma dessas duas: aí
    não confirma, porque não sabe o que está aceitando.
    """
    log("Clicando em Submit")
    botao = page.locator("button.primary-action[data-label='Submit']").first
    if botao.count() == 0:
        salvar_debug(page, "submit_nao_encontrado")
        raise RuntimeError("Não achei o botão Submit — veja debug/")
    botao.click()

    for etapa in range(4):
        sim = _yes_visivel(page, 15_000 if etapa == 0 else 8_000)
        if sim is None:
            break   # não apareceu mais nenhuma janela

        modal = sim.locator("xpath=ancestor::div[contains(@class,'modal-content')][1]")
        texto = " ".join((modal.inner_text(timeout=5000) or "").split())

        conhecida = ("permanently submit" in texto.lower()
                     or "to be marked as" in texto.lower())

        if not conhecida:
            _clicar_nao(modal.locator("button.btn-modal-secondary").first, page)
            salvar_debug(page, "janela_desconhecida")
            raise RuntimeError(
                f"Apareceu uma confirmação inesperada. Cliquei em No. "
                f"Texto: {texto[:150]}"
            )

        log(f"Confirmando: {texto}")
        sim.click()

        # Espera essa janela sumir antes de procurar a próxima
        for _ in range(20):
            page.wait_for_timeout(250)
            try:
                if not sim.is_visible():
                    break
            except Exception:
                break
        page.wait_for_timeout(800)

    # ── Confirma que foi aceito ──
    for _ in range(20):
        page.wait_for_timeout(1000)
        try:
            d = _itens_do_documento(page)
            if d and d.get("docstatus") == 1:
                log("Submetido — documento confirmado pelo sistema")
                return
        except Exception:
            pass

    salvar_debug(page, "submit_sem_confirmacao")
    raise RuntimeError(
        "Confirmei o Submit mas não consegui verificar que foi aceito. "
        "Confira o documento antes de rodar de novo — veja debug/"
    )


# ── Processa um pedido ──────────────────────────────────────────

def processar(page, pedido, modo):
    so = pedido["sales_order"]
    log(f"━━━ {so} (NF {pedido['nf']}, {pedido['carrier']}) ━━━")

    abrir_pedido(page, so)

    if modo == "reconhecer":
        salvar_debug(page, f"1_shipping_core_{so}", DIR_RECON)

    itens, estados = localizar_itens_pendentes(page)

    if not estados:
        # Nem um item apareceu: a tela não carregou (ou o pedido não existe
        # nessa URL). Isso é problema, não "já tinha baixa".
        salvar_debug(page, f"itens_nao_carregaram_{so}")
        log("Os itens do pedido não apareceram na tela — veja debug/")
        return {"ok": False, "motivo": "a tela do pedido não carregou os itens (veja debug/)"}

    log(f"Itens do pedido: {', '.join(estados)}")

    if not itens:
        log("Nenhum item em 'Delivered to Carrier' — o pedido já teve baixa no Zecore")
        return {"ok": False, "motivo": "sem item pendente", "estados": estados}

    # Escolhe o item da NF da planilha. O número da nota está dentro do
    # Master Guide, então dá pra conferir sem depender de outra tela.
    # Sem NF conhecida (modo --pedido), usa o primeiro pendente.
    nf_esperada = _nf_normalizada(pedido.get("nf"))
    dados = None

    for linha, idx in itens:
        d = extrair_dados_item(page, linha)
        nf_item = _nf_do_master_guide(d["master_guide"])

        if not nf_esperada or nf_item == nf_esperada:
            dados = d
            log(f"Item pendente na linha {idx} (NF {nf_item or '?'})")
            break

        log(f"Item da linha {idx} é da NF {nf_item or '?'}, "
            f"não da {nf_esperada} — procurando o próximo")
        _fechar_item(page)

    if dados is None:
        return {
            "ok": False,
            "motivo": f"nenhum item pendente da NF {nf_esperada} neste pedido",
        }

    log(f"Master Guide: {dados['master_guide']}")
    log(f"Arrangement: {dados['arrangement']}")

    if modo == "reconhecer":
        salvar_debug(page, f"2_item_{so}", DIR_RECON)
        page.goto(dados["arrangement_url"], wait_until="domcontentloaded")
        page.wait_for_timeout(4000)
        salvar_debug(page, f"3_arrangement_{so}", DIR_RECON)
        log("Modo reconhecimento: parei aqui, nada foi criado")
        return {"ok": True, "modo": "reconhecer", **dados}

    criar_delivery_task(page, dados["arrangement_url"])
    url_task = page.url

    _, esperadas = apagar_outras_linhas(page, dados["master_guide"])
    conferir_antes_do_submit(page, dados["master_guide"], esperadas)

    salvar(page)

    # Confere de novo depois de salvar — a tela recarrega e é barato
    conferir_antes_do_submit(page, dados["master_guide"], esperadas)

    if modo == "simular":
        log("Modo simulação: NÃO vou submeter")
        log(f"Rascunho deixado em: {url_task}")
        return {
            "ok": True, "modo": "simular",
            "mantidas": esperadas, "rascunho": url_task, **dados
        }

    submeter(page, esperadas)
    return {
        "ok": True, "modo": "executar",
        "mantidas": esperadas, "task": url_task, **dados
    }


# ── Principal ───────────────────────────────────────────────────

def main():
    p = argparse.ArgumentParser()
    p.add_argument("--reconhecer", action="store_true",
                   help="só navega e tira print — não modifica nada")
    p.add_argument("--simular", action="store_true",
                   help="faz tudo menos o Submit")
    p.add_argument("--executar", action="store_true",
                   help="faz tudo, incluindo o Submit (IRREVERSÍVEL)")
    p.add_argument("--headed", action="store_true")
    p.add_argument("--limite", type=int, default=LIMITE_PADRAO,
                   help="máximo de baixas por rodada (0 = sem limite)")
    p.add_argument("--verificar", action="store_true",
                   help="só conta os pendentes da planilha, sem abrir o Zecore")
    p.add_argument("--pedido", help="processa só este Sales Order")
    p.add_argument("--nf", help="com --pedido: a NF da planilha (ex: 115541-1), "
                                "pra escolher o item certo em pedido com várias NFs")
    args = p.parse_args()

    # ── Modo verificar: só conta, pra o workflow decidir se instala o
    #    navegador. Rodada sem nada a fazer termina em segundos. ──
    if args.verificar:
        creds_json = os.environ.get("GOOGLE_CREDS_JSON")
        if not creds_json:
            sys.exit("Falta GOOGLE_CREDS_JSON")
        sheet_id = os.environ.get("SHEET_ID", SHEET_ID_PADRAO)
        hist = carregar_historico()
        todos = ler_pendentes(sheet_id, creds_json, None)
        pendentes = [
            pd for pd in todos
            if _chave(pd["sales_order"], _nf_normalizada(pd["nf"])) not in hist["baixas"]
        ]
        # Mesmo sem nada pra baixar, os já baixados precisam sair da Rastreamento
        if len(todos) != len(pendentes):
            enfileirar_do_historico(FilaBaixas(sheet_id, creds_json), todos, hist)
        log(f"{len(pendentes)} pedido(s) pra processar (fora os que já estão no histórico)")
        saida = os.environ.get("GITHUB_OUTPUT")
        if saida:
            with open(saida, "a", encoding="utf-8") as f:
                f.write(f"pendentes={len(pendentes)}\n")
        return

    if sync_playwright is None:
        sys.exit("O Playwright não está instalado (pip install playwright)")

    if args.reconhecer:
        modo = "reconhecer"
    elif args.executar:
        modo = "executar"
    else:
        modo = "simular"   # o padrão é o modo seguro

    log(f"Modo: {modo.upper()}")
    if modo == "executar":
        log("⚠️  O Submit é IRREVERSÍVEL")

    usuario = os.environ.get("ZECORE_USER")
    senha = os.environ.get("ZECORE_PASSWORD")
    creds_json = os.environ.get("GOOGLE_CREDS_JSON")
    sheet_id = os.environ.get("SHEET_ID", SHEET_ID_PADRAO)

    if not usuario or not senha:
        sys.exit("Faltam ZECORE_USER e ZECORE_PASSWORD")

    if args.pedido:
        pendentes = [{"linha": 0, "nf": args.nf or "?", "sales_order": args.pedido,
                      "carrier": "?", "status": "?"}]
        log(f"Pedido informado na linha de comando: {args.pedido}")
    else:
        if not creds_json:
            sys.exit("Falta GOOGLE_CREDS_JSON (ou use --pedido)")
        pendentes = ler_pendentes(sheet_id, creds_json, None)

    # Pula quem já está no histórico, sem nem abrir o Zecore
    hist = carregar_historico()
    fila = FilaBaixas(sheet_id, creds_json) if (modo == "executar" and creds_json) else None
    if not args.pedido:
        if fila:
            enfileirar_do_historico(fila, pendentes, hist)
        antes = len(pendentes)
        pendentes = [
            pd for pd in pendentes
            if _chave(pd["sales_order"], _nf_normalizada(pd["nf"])) not in hist["baixas"]
        ]
        if antes != len(pendentes):
            log(f"{antes - len(pendentes)} já estavam no histórico de baixas — pulados")

    if not pendentes:
        log("Nada a fazer")
        salvar_historico(hist)
        return

    resultados = []
    interrompida = False

    with sync_playwright() as pw:
        navegador = pw.chromium.launch(
            headless=not args.headed,
            args=["--disable-blink-features=AutomationControlled"],
        )
        contexto = navegador.new_context(
            viewport={"width": 1920, "height": 1080}, locale="pt-BR"
        )
        page = contexto.new_page()

        try:
            fazer_login(page, usuario, senha)

            # O limite conta só pedidos efetivamente processados. A coluna
            # "Status zecore" da planilha só atualiza na importação do
            # QuickSight, então entre uma e outra ela ainda lista pedidos
            # já atualizados — o robô abre, vê que não há item pendente e
            # pula. Esses não gastam o limite.
            feitos = 0
            seguidos = 0
            for pedido in pendentes:
                if args.limite and feitos >= args.limite:
                    log(f"Limite de {args.limite} atingido — parando")
                    break
                if seguidos >= MAX_PROBLEMAS_SEGUIDOS:
                    interrompida = True
                    log(f"{seguidos} problemas seguidos — interrompendo a rodada "
                        "(pode ser mudança na tela do Zecore)")
                    break
                try:
                    r = processar(page, pedido, modo)
                    r["sales_order"] = pedido["sales_order"]
                    r["pedido"] = pedido
                    resultados.append(r)
                    if r.get("ok"):
                        feitos += 1
                    problema = not r.get("ok") and r.get("motivo") != "sem item pendente"
                    seguidos = seguidos + 1 if problema else 0
                    # Grava a cada baixa, não só no fim: se der erro no
                    # meio, o que já foi feito não se perde
                    if r.get("ok") and r.get("modo") == "executar":
                        registrar_baixa(hist, pedido, r)
                        salvar_historico(hist)
                        if fila:
                            nf_fila = (pedido["nf"] if _nf_normalizada(pedido.get("nf"))
                                       else _nf_do_master_guide(r.get("master_guide")))
                            fila.adicionar([(pedido, nf_fila, "baixa pelo robô")])
                    elif (modo == "executar" and r.get("motivo") == "sem item pendente"
                          and registrar_ja_baixado(hist, pedido, r.get("estados"))):
                        salvar_historico(hist)
                        log("Anotado no histórico — não será reaberto nas próximas rodadas")
                except Exception as e:
                    seguidos += 1
                    log(f"✗ {pedido['sales_order']}: {e}")
                    salvar_debug(page, f"erro_{pedido['sales_order']}")
                    resultados.append({
                        "ok": False,
                        "sales_order": pedido["sales_order"],
                        "erro": str(e),
                    })

        finally:
            contexto.close()
            navegador.close()

    log("")
    log("━━━ Resumo ━━━")
    for r in resultados:
        if r.get("ok"):
            log(f"✓ {r['sales_order']} — {r.get('modo')}" +
                (f", {r.get('mantidas')} linha(s) mantida(s)"
                 if r.get("mantidas") is not None else ""))
            if r.get("rascunho"):
                log(f"   rascunho para limpar: {r['rascunho']}")
        elif r.get("motivo") == "sem item pendente":
            log(f"— {r['sales_order']} — já estava atualizado, pulado")
        else:
            log(f"✗ {r['sales_order']} — {r.get('erro') or r.get('motivo')}")

    salvar_historico(hist)

    # Resumo no Chat — só quando houve baixa de verdade
    if modo == "executar":
        baixas = []
        for r in resultados:
            if r.get("ok") and r.get("modo") == "executar":
                pd = r["pedido"]
                baixas.append({
                    "pre": "PRE ENTREGUE" in _norm_status(pd.get("status")),
                    "nf": _nf_do_master_guide(r.get("master_guide")) or _nf_normalizada(pd.get("nf")),
                    "sales_order": r["sales_order"],
                    "carrier": _carrier_curto(
                        pd.get("carrier") if pd.get("carrier") not in (None, "?")
                        else r.get("arrangement")
                    ),
                })
        problemas = [
            {"sales_order": r["sales_order"], "motivo": r.get("erro") or r.get("motivo") or "?"}
            for r in resultados
            if not r.get("ok") and r.get("motivo") != "sem item pendente"
        ]
        ja_tinham = sum(1 for r in resultados if r.get("motivo") == "sem item pendente")

        avisar_chat(os.environ.get("CHAT_WEBHOOK_BAIXAS"), baixas, problemas, ja_tinham, interrompida)


if __name__ == "__main__":
    main()

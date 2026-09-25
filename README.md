# Automação QuickSight → Dash_fedex

Substitui o processo manual de exportar o CSV do QuickSight e colar na
aba "Base de dados".

## Estrutura do repositório

```
.
├── scripts/
│   └── quicksight_export.py
├── .github/
│   └── workflows/
│       └── quicksight.yml
└── requirements.txt
```

Crie o repositório com essa estrutura e coloque cada arquivo no lugar
indicado (o `quicksight.yml` que veio junto vai em `.github/workflows/`).

---

## 1. Service account do Google

A planilha precisa aceitar escrita vinda do robô.

1. No Google Cloud Console, crie (ou reaproveite) uma **service account**
2. Gere uma chave JSON e baixe
3. Abra a **Dash_fedex** → Compartilhar → adicione o e-mail da service
   account (termina em `.iam.gserviceaccount.com`) como **Editor**

> Se você já tem uma service account do dashboard do Netlify, pode usar
> a mesma — só precisa dar acesso de editor a esta planilha.

---

## 2. Secrets no GitHub

No repositório: **Settings → Secrets and variables → Actions → New secret**

| Nome | Valor |
|---|---|
| `QS_USER` | o usuário do QuickSight |
| `QS_PASSWORD` | a senha |
| `GOOGLE_CREDS_JSON` | o conteúdo **inteiro** do JSON da service account |

---

## 3. Testar local antes de agendar

Vale rodar na sua máquina primeiro, com o navegador visível, pra
confirmar que os cliques estão certos:

```bash
pip install -r requirements.txt
playwright install chromium

export QS_USER="..."
export QS_PASSWORD="..."

# Só baixa o CSV, sem tocar na planilha
python scripts/quicksight_export.py --headed --so-baixar
```

Se o arquivo aparecer em `saida/details_in_transit.csv`, a parte difícil
funcionou. Aí testa o fluxo completo:

```bash
export GOOGLE_CREDS_JSON="$(cat caminho/para/credenciais.json)"
python scripts/quicksight_export.py --headed
```

---

## 4. Agendamento

O workflow está configurado pra rodar **07h e 13h (horário de Brasília),
de segunda a sexta**. O cron do GitHub usa UTC, por isso aparece `10` e
`16` no arquivo.

Pra mudar o horário, edite as linhas `cron` em `.github/workflows/quicksight.yml`.

Também dá pra rodar na hora pelo botão **Run workflow** na aba Actions.

---

## Quando quebrar

Automação de interface quebra quando o fornecedor mexe no layout — é
esperado, não é defeito. O script foi feito pra facilitar o conserto:

- Toda falha salva **screenshot e HTML** da tela no momento do erro
- No GitHub Actions esses arquivos ficam em **Artifacts** na execução que
  falhou (guardados por 14 dias)
- A mensagem de erro diz em qual etapa parou

Os pontos que mais tendem a quebrar, em ordem:

1. **O menu do visual** (`abrir_menu_visual`) — depende do `aria-label`
   "Opções de menu" e do hover revelar a barra de ícones
2. **O texto da opção de exportar** — se a AWS renomear, é só acrescentar
   o texto novo na lista `TEXTOS_EXPORTAR`
3. **Os campos de login** — hoje são o primeiro input de cada tela

Com o screenshot em mãos, ajustar costuma ser questão de trocar um
seletor.

---

## Detalhes que valem saber

**A aba só é limpa depois que o CSV chega.** Se o download falhar, a
"Base de dados" continua com os dados anteriores em vez de ficar vazia.

**O export do QuickSight é assíncrono.** Ele mostra "Trabalhando no seu
arquivo CSV" e só depois libera o download — por isso o timeout é de 5
minutos, bem acima do normal.

**O runner do GitHub usa IP da Microsoft.** Se a AWS de vocês tiver
restrição de IP no login, vai falhar. Dá pra descobrir logo na primeira
execução.

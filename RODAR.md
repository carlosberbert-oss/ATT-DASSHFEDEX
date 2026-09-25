# Como rodar (Windows / PowerShell)

Dentro da pasta do projeto:

```powershell
$env:QS_USER = "carlos.berbert@zeb.mx"
$env:QS_PASSWORD = Read-Host "Senha"
$env:GOOGLE_CREDS_JSON = Get-Content "CAMINHO\DO\SEU.json" -Raw

python scripts/quicksight_export.py --headed
```

As variáveis valem só na sessão atual do PowerShell — se fechar o
terminal, precisa definir de novo.

## Variações

- `--so-baixar` → baixa o CSV mas não escreve na planilha (bom pra testar)
- sem `--headed` → roda sem abrir o navegador (é como roda no GitHub Actions)

Se travar, olhe a pasta `debug/` — tem screenshot e HTML do erro.

## Agendamento automático

Roda às 08h e 18h (horário de Brasília), de segunda a sexta.
Configurado em `.github/workflows/quicksight.yml`.

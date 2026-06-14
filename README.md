# NK Contas a Pagar — Bot Telegram

Bot de gestão de contas a pagar para NK Soluções e NK Pré-Moldados.

## Funcionalidades

- Lançamento guiado (`/nova`) e rápido (`/cp`)
- Marcação de pagamento (`/pagar`)
- Listagem e resumos (dia / semana / mês / geral)
- Import em lote via `.xlsx`
- Alertas automáticos: 7d / 3d / 1d / no dia
- Resumo diário às 07h (horário de Brasília)

## Variáveis de ambiente (Railway)

| Variável | Valor |
|---|---|
| `BOT_TOKEN` | Token do @nk_cp_bot (BotFather) |
| `SHEET_ID` | `1ZXQo5V2NomogdETrKMNfN42GWXGrGWmSXdls1vdifD0` |
| `GOOGLE_CREDENTIALS` | Conteúdo completo do `credentials.json` |
| `MY_TELEGRAM_ID` | `647725027` |
| `NIXPACKS_PYTHON_VERSION` | `3.11` |

## Google Sheets

Planilha: **NK - Contas a Pagar**
ID: `1ZXQo5V2NomogdETrKMNfN42GWXGrGWmSXdls1vdifD0`

Service account: `nk-bot-sheets@nk-bot-498502.iam.gserviceaccount.com`

Abas criadas automaticamente na primeira execução:
- `Contas` — dados principais
- `Config` — listas de referência
- `Log` — histórico de ações

## Formato do XLSX para import

| Empresa | Categoria | Descrição | Credor | Valor | Vencimento |
|---|---|---|---|---|---|
| NK Soluções | DAS | Simples Nacional Jun | Receita Federal | 1500 | 20/06/2026 |

## Deploy Railway

1. Cria novo serviço no projeto existente
2. Conecta ao GitHub (`leobdiasg4-maker/nk-cp-bot`)
3. Adiciona todas as variáveis acima
4. Deploy automático

## Comando rápido

```
/cp Credor Valor DD/MM/AAAA Empresa Categoria
/cp Simples_Nacional 1500 20/06/2026 NK_Soluções DAS
```

Use `_` no lugar de espaços nos argumentos.

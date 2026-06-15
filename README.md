# NK Contas a Pagar

Sistema de gestão de contas a pagar para **NK Soluções** e **NK Pré-Moldados**.

Composto por três partes integradas:
- **Bot Telegram** `@nk_cp_bot` — lançamento e alertas no celular
- **Web** `leobdiasg4-maker.github.io/nk-cp-bot` — painel completo no navegador
- **API Flask** no Railway — backend de escrita segura no Google Sheets

---

## Funcionalidades

### Bot Telegram
- Lançamento guiado (`/nova`) com botões e atalhos de data
- Lançamento rápido por comando (`/cp`)
- Marcação de pagamento (`/pagar`) com geração automática de próxima parcela para contas recorrentes
- Listagem de pendentes (`/listar`)
- Resumos: `/resumo`, `/resumo_dia`, `/resumo_semana`, `/resumo_mes`
- Alertas automáticos: 7d / 3d / 1d / no dia
- Resumo diário às 07h (horário de Brasília)
- Import em lote via arquivo `.xlsx`
- Listas de empresa/categoria lidas dinamicamente da planilha

### Web
- Login com senha individual por usuário
- Dashboard com KPIs e alertas visuais
- Listagem com filtros por empresa, status e categoria
- Cadastro de nova conta
- Registro de pagamento com data, valor e conta bancária
- Geração automática de próxima parcela para recorrentes
- Aba **Configurações** para gerenciar empresas, categorias, contas e status
- Troca de senha individual

---

## Arquitetura

```
GitHub Pages (docs/index.html)
    │ leitura → Google Sheets API (API Key pública)
    │ escrita → Railway Flask API (service account)
    
Railway — Serviço 1: Bot Telegram (bot.py)
Railway — Serviço 2: API Flask (api.py)
    └── Google Sheets (planilha NK - Contas a Pagar)
```

---

## Variáveis de ambiente — Railway

### Serviço Bot (`bot.py`)
| Variável | Valor |
|---|---|
| `BOT_TOKEN` | Token do `@nk_cp_bot` (BotFather) |
| `SHEET_ID` | `1ZXQo5V2NomogdETrKMNfN42GWXGrGWmSXdls1vdifD0` |
| `GOOGLE_CREDENTIALS` | Conteúdo completo do `credentials.json` |
| `MY_TELEGRAM_ID` | `647725027` |
| `NIXPACKS_PYTHON_VERSION` | `3.11` |

### Serviço API (`api.py`)
| Variável | Valor |
|---|---|
| `SHEET_ID` | `1ZXQo5V2NomogdETrKMNfN42GWXGrGWmSXdls1vdifD0` |
| `GOOGLE_CREDENTIALS` | Conteúdo completo do `credentials.json` |
| `NIXPACKS_PYTHON_VERSION` | `3.11` |

Start command da API: `gunicorn api:app`

URL pública da API: `https://nk-cp-bot-production.up.railway.app`

---

## Google Sheets

**Planilha:** NK - Contas a Pagar  
**ID:** `1ZXQo5V2NomogdETrKMNfN42GWXGrGWmSXdls1vdifD0`  
**Service account:** `nk-bot-sheets@nk-bot-498502.iam.gserviceaccount.com`

### Abas criadas automaticamente
| Aba | Conteúdo |
|---|---|
| `Contas` | Dados principais de contas a pagar |
| `Config` | Listas de Empresa, Categoria, Conta, Status (formato Tipo\|Valor) |
| `Log` | Histórico de ações |
| `Usuarios` | Login, senha e nome dos usuários web |

### Estrutura da aba Config (formato vertical)
| Tipo | Valor |
|---|---|
| Empresa | NK Soluções |
| Empresa | NK Pré-Moldados |
| Categoria | DAS |
| Categoria | INSS |
| Conta | Nubank PJ |
| ... | ... |

---

## Web — GitHub Pages

**URL:** `https://leobdiasg4-maker.github.io/nk-cp-bot`  
**Fonte:** pasta `docs/` do repo  
**Arquivo:** `docs/.nojekyll` necessário para desativar Jekyll

### Configuração no index.html
```javascript
const SHEET_ID = '1ZXQo5V2NomogdETrKMNfN42GWXGrGWmSXdls1vdifD0';
const API_KEY  = 'SUA_API_KEY_GOOGLE';   // leitura pública
const API_URL  = 'https://nk-cp-bot-production.up.railway.app'; // escrita
```

### Usuários padrão (primeiro acesso)
| Login | Senha |
|---|---|
| leonardo | NK2026 |
| nicanor | NK2026 |

Cada usuário pode alterar a própria senha na aba **Minha Senha**.

---

## Bot — Comandos

### Lançamento
```
/nova                          — guiado passo a passo (com botões)
/cp Credor Valor Data Empresa Categoria  — rápido
```
Exemplo rápido:
```
/cp Simples_Nacional 1500 20/06/2026 NK_Soluções DAS
```
Use `_` no lugar de espaços. Atalhos de data disponíveis no modo guiado: Hoje / Em 7 dias / Em 15 dias / Em 30 dias.

### Pagamento
```
/pagar CP0001                        — marca pago hoje
/pagar CP0001 15/06/2026             — com data
/pagar CP0001 15/06/2026 1500        — com data e valor pago
```
Contas com `Recorrente = Sim` geram automaticamente a próxima parcela.

### Consulta
```
/listar          — pendentes ordenados por vencimento
/resumo          — visão geral de pendentes
/resumo_dia      — vencendo hoje
/resumo_semana   — próximos 7 dias
/resumo_mes      — mês atual
```

### Import XLSX
Envie um arquivo `.xlsx` com as colunas:

| Empresa | Categoria | Descrição | Credor | Valor | Vencimento |
|---|---|---|---|---|---|
| NK Soluções | DAS | Simples Nacional Jun | Receita Federal | 1500 | 20/06/2026 |

---

## Pendências conhecidas
- [ ] Bot ainda não deployado no Railway (apenas API está no ar)
- [ ] Valor salvo como texto em lançamentos antigos (novos já salvam como float)
- [ ] Inline buttons para funcionar no Telegram Web/Desktop

---

## Pessoas e acessos
- **Leonardo** — admin, Telegram ID `647725027`
- **Nicanor** — acesso web apenas (sem Telegram no bot por ora)
- **Service account:** `nk-bot-sheets@nk-bot-498502.iam.gserviceaccount.com`

import os
import json
import logging
import asyncio
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo

import gspread
from google.oauth2.service_account import Credentials
from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    ConversationHandler, ContextTypes, filters
)
import openpyxl

# ── Logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    format="%(asctime)s [%(levelname)s] %(message)s",
    level=logging.INFO,
)
log = logging.getLogger(__name__)

# ── Config ───────────────────────────────────────────────────────────────────
BOT_TOKEN      = os.environ["BOT_TOKEN"]
SHEET_ID       = os.environ["SHEET_ID"]
MY_TELEGRAM_ID = int(os.environ.get("MY_TELEGRAM_ID", "647725027"))
TZ             = ZoneInfo("America/Sao_Paulo")

EMPRESAS   = ["NK Soluções", "NK Pré-Moldados"]
CATEGORIAS = ["DAS", "INSS", "FGTS", "Folha", "Fornecedor", "Aluguel",
              "Empréstimo", "Honorários", "Outros"]
CONTAS     = ["Nubank PJ", "Sicoob", "C6", "PagBank",
              "Santander NK Soluções", "Santander NK Pré-Moldados",
              "Cora", "Dinheiro"]
STATUS_OPT = ["Pendente", "Pago", "Atrasado", "Parcial"]
FREQ_OPT   = ["Única", "Mensal", "Semanal", "Anual"]

# Colunas da aba Contas
COLS = [
    "ID", "Empresa", "Categoria", "Descrição", "Credor",
    "Valor (R$)", "Vencimento", "Status", "Data Pagamento",
    "Valor Pago", "Conta Bancária", "Recorrente", "Frequência",
    "Observação", "Lançado por", "Criado em"
]

# Estados do ConversationHandler (lançamento guiado)
(S_EMPRESA, S_CATEGORIA, S_DESCRICAO, S_CREDOR, S_VALOR,
 S_VENCIMENTO, S_RECORRENTE, S_FREQUENCIA, S_OBS, S_CONFIRMA) = range(10)

# ── Google Sheets ─────────────────────────────────────────────────────────────
def get_sheet():
    raw = os.environ["GOOGLE_CREDENTIALS"]
    if not raw.strip().startswith("{"):
        raw = "{" + raw + "}"
    info = json.loads(raw)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SHEET_ID)
    return sh

def ensure_sheets(sh):
    """Garante que as abas existam com cabeçalhos."""
    existing = [w.title for w in sh.worksheets()]

    if "Contas" not in existing:
        ws = sh.add_worksheet("Contas", rows=1000, cols=len(COLS))
        ws.append_row(COLS)
    if "Config" not in existing:
        ws = sh.add_worksheet("Config", rows=50, cols=4)
        ws.append_row(["Empresas", "Categorias", "Contas", "Status"])
        for i, (e, c, ct, s) in enumerate(zip(
            EMPRESAS + [""] * max(0, len(CATEGORIAS) - len(EMPRESAS)),
            CATEGORIAS + [""] * max(0, len(EMPRESAS) - len(CATEGORIAS)),
            CONTAS + [""] * max(0, len(CATEGORIAS) - len(CONTAS)),
            STATUS_OPT + [""] * max(0, len(CATEGORIAS) - len(STATUS_OPT)),
        )):
            ws.append_row([e, c, ct, s])
    if "Log" not in existing:
        ws = sh.add_worksheet("Log", rows=1000, cols=4)
        ws.append_row(["Timestamp", "Ação", "ID Conta", "Usuário"])

def next_id(ws):
    vals = ws.col_values(1)[1:]  # pula cabeçalho
    ids = [int(v.replace("CP", "")) for v in vals if v.startswith("CP")]
    return f"CP{(max(ids) + 1) if ids else 1:04d}"

def get_all_contas(ws):
    rows = ws.get_all_values()
    if len(rows) <= 1:
        return []
    headers = rows[0]
    return [dict(zip(headers, r)) for r in rows[1:]]

def append_conta(ws, data: dict):
    row = [data.get(c, "") for c in COLS]
    ws.append_row(row)

def update_status(ws, cp_id: str, new_status: str, data_pag: str = "", valor_pago: str = ""):
    rows = ws.get_all_values()
    for i, row in enumerate(rows[1:], start=2):
        if row[0] == cp_id:
            ws.update_cell(i, 8, new_status)
            if data_pag:
                ws.update_cell(i, 9, data_pag)
            if valor_pago:
                ws.update_cell(i, 10, valor_pago)
            return True
    return False

# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt_brl(val):
    try:
        return f"R$ {float(str(val).replace(',', '.')):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(val)

def parse_date(s: str):
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s.strip(), fmt).date()
        except ValueError:
            continue
    return None

def days_until(venc_str: str):
    d = parse_date(venc_str)
    if not d:
        return None
    return (d - date.today()).days

def status_emoji(status: str, days):
    if status == "Pago":
        return "✅"
    if days is None:
        return "❓"
    if days < 0:
        return "🔴"
    if days == 0:
        return "🚨"
    if days <= 3:
        return "🟠"
    if days <= 7:
        return "🟡"
    return "🟢"

def kb(options, columns=2):
    rows = [options[i:i+columns] for i in range(0, len(options), columns)]
    return ReplyKeyboardMarkup(rows, one_time_keyboard=True, resize_keyboard=True)

def is_authorized(update: Update):
    return update.effective_user.id == MY_TELEGRAM_ID

# ── Resumo ────────────────────────────────────────────────────────────────────
def build_resumo(contas, titulo: str) -> str:
    if not contas:
        return f"📋 *{titulo}*\n\nNenhuma conta encontrada."

    total_pend = sum(
        float(str(c.get("Valor (R$)", "0")).replace(",", ".") or 0)
        for c in contas if c.get("Status") in ("Pendente", "Atrasado", "Parcial")
    )
    atrasadas = [c for c in contas if days_until(c.get("Vencimento", "")) is not None and days_until(c.get("Vencimento", "")) < 0 and c.get("Status") != "Pago"]
    vence_hoje = [c for c in contas if days_until(c.get("Vencimento", "")) == 0 and c.get("Status") != "Pago"]
    vence_7    = [c for c in contas if days_until(c.get("Vencimento", "")) is not None and 1 <= days_until(c.get("Vencimento", "")) <= 7 and c.get("Status") != "Pago"]

    lines = [f"📋 *{titulo}*\n"]
    lines.append(f"💰 Total pendente: *{fmt_brl(total_pend)}*")
    lines.append(f"🔴 Atrasadas: {len(atrasadas)} | 🚨 Vencem hoje: {len(vence_hoje)} | 🟡 Próx. 7 dias: {len(vence_7)}\n")

    def bloco(label, lista):
        if not lista:
            return
        lines.append(f"*{label}*")
        for c in lista:
            d = days_until(c.get("Vencimento", ""))
            em = status_emoji(c.get("Status", ""), d)
            lines.append(
                f"{em} {c.get('ID','')} | {c.get('Empresa','')} | {c.get('Credor','')} "
                f"| {fmt_brl(c.get('Valor (R$)',0))} | {c.get('Vencimento','')}"
            )
        lines.append("")

    bloco("🔴 Atrasadas", atrasadas)
    bloco("🚨 Vencem hoje", vence_hoje)
    bloco("🟡 Próximos 7 dias", vence_7)

    return "\n".join(lines)

# ── Alertas automáticos ───────────────────────────────────────────────────────
async def job_alertas(context: ContextTypes.DEFAULT_TYPE):
    try:
        sh = get_sheet()
        ws = sh.worksheet("Contas")
        contas = get_all_contas(ws)
        hoje = date.today()

        alertas = []
        for c in contas:
            if c.get("Status") == "Pago":
                continue
            d = days_until(c.get("Vencimento", ""))
            if d in (0, 1, 3, 7):
                alertas.append((d, c))

        if not alertas:
            return

        msg = "⏰ *Alertas de Vencimento*\n\n"
        for d, c in sorted(alertas, key=lambda x: x[0]):
            em = status_emoji(c.get("Status", ""), d)
            prazo = "HOJE" if d == 0 else f"em {d} dia(s)"
            msg += (
                f"{em} *{c.get('ID','')}* — {c.get('Credor','')} ({c.get('Empresa','')})\n"
                f"   {fmt_brl(c.get('Valor (R$)',0))} — vence *{prazo}* ({c.get('Vencimento','')})\n\n"
            )

        await context.bot.send_message(
            chat_id=MY_TELEGRAM_ID,
            text=msg,
            parse_mode="Markdown"
        )
    except Exception as e:
        log.error("Erro job_alertas: %s", e)

async def job_resumo_diario(context: ContextTypes.DEFAULT_TYPE):
    try:
        sh = get_sheet()
        ws = sh.worksheet("Contas")
        contas = get_all_contas(ws)
        hoje = date.today()
        contas_ativas = [c for c in contas if c.get("Status") != "Pago"]
        msg = build_resumo(contas_ativas, f"Resumo Diário — {hoje.strftime('%d/%m/%Y')}")
        await context.bot.send_message(
            chat_id=MY_TELEGRAM_ID,
            text=msg,
            parse_mode="Markdown"
        )
    except Exception as e:
        log.error("Erro job_resumo_diario: %s", e)

# ── Comando /start ─────────────────────────────────────────────────────────────
async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_authorized(update):
        return
    await update.message.reply_text(
        "👋 *NK Contas a Pagar*\n\n"
        "Comandos disponíveis:\n"
        "*/nova* — Lançar conta (modo guiado)\n"
        "*/pagar* — Marcar conta como paga\n"
        "*/listar* — Listar contas pendentes\n"
        "*/resumo* — Resumo geral\n"
        "*/resumo\\_dia* — Resumo do dia\n"
        "*/resumo\\_semana* — Próximos 7 dias\n"
        "*/resumo\\_mes* — Mês atual\n"
        "*/ajuda* — Ver todos os comandos\n\n"
        "Ou use o comando rápido:\n"
        "`/cp Credor Valor DD/MM/AAAA Empresa Categoria`\n"
        "Ex: `/cp Simples\\_Nacional 1500 20/06/2026 NK\\_Soluções DAS`\n\n"
        "📎 Envie um arquivo `.xlsx` para importar em lote.",
        parse_mode="Markdown"
    )

# ── Lançamento guiado ─────────────────────────────────────────────────────────
async def cmd_nova(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_authorized(update):
        return ConversationHandler.END
    context.user_data.clear()
    await update.message.reply_text(
        "📝 *Nova conta a pagar*\n\nQual empresa?",
        reply_markup=kb(EMPRESAS),
        parse_mode="Markdown"
    )
    return S_EMPRESA

async def step_empresa(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["Empresa"] = update.message.text
    await update.message.reply_text("Categoria:", reply_markup=kb(CATEGORIAS))
    return S_CATEGORIA

async def step_categoria(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["Categoria"] = update.message.text
    await update.message.reply_text("Descrição (texto livre):", reply_markup=ReplyKeyboardRemove())
    return S_DESCRICAO

async def step_descricao(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["Descrição"] = update.message.text
    await update.message.reply_text("Credor (fornecedor/órgão):")
    return S_CREDOR

async def step_credor(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["Credor"] = update.message.text
    await update.message.reply_text("Valor (ex: 1500,00):")
    return S_VALOR

async def step_valor(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["Valor (R$)"] = update.message.text.replace("R$", "").strip()
    await update.message.reply_text("Data de vencimento (DD/MM/AAAA):")
    return S_VENCIMENTO

async def step_vencimento(update: Update, context: ContextTypes.DEFAULT_TYPE):
    d = parse_date(update.message.text)
    if not d:
        await update.message.reply_text("❌ Data inválida. Use DD/MM/AAAA:")
        return S_VENCIMENTO
    context.user_data["Vencimento"] = d.strftime("%d/%m/%Y")
    await update.message.reply_text("É recorrente?", reply_markup=kb(["Sim", "Não"]))
    return S_RECORRENTE

async def step_recorrente(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["Recorrente"] = update.message.text
    if update.message.text == "Sim":
        await update.message.reply_text("Frequência:", reply_markup=kb(FREQ_OPT))
        return S_FREQUENCIA
    context.user_data["Frequência"] = "Única"
    await update.message.reply_text("Observação (ou /pular):", reply_markup=ReplyKeyboardRemove())
    return S_OBS

async def step_frequencia(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["Frequência"] = update.message.text
    await update.message.reply_text("Observação (ou /pular):", reply_markup=ReplyKeyboardRemove())
    return S_OBS

async def step_obs(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["Observação"] = "" if update.message.text == "/pular" else update.message.text
    d = context.user_data
    resumo = (
        f"✅ *Confirmar lançamento?*\n\n"
        f"🏢 Empresa: {d.get('Empresa')}\n"
        f"📂 Categoria: {d.get('Categoria')}\n"
        f"📝 Descrição: {d.get('Descrição')}\n"
        f"🏦 Credor: {d.get('Credor')}\n"
        f"💰 Valor: {fmt_brl(d.get('Valor (R$)', 0))}\n"
        f"📅 Vencimento: {d.get('Vencimento')}\n"
        f"🔁 Recorrente: {d.get('Recorrente')} ({d.get('Frequência')})\n"
        f"📎 Obs: {d.get('Observação') or '—'}"
    )
    await update.message.reply_text(resumo, reply_markup=kb(["✅ Confirmar", "❌ Cancelar"]), parse_mode="Markdown")
    return S_CONFIRMA

async def step_confirma(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if "Confirmar" not in update.message.text:
        await update.message.reply_text("❌ Lançamento cancelado.", reply_markup=ReplyKeyboardRemove())
        return ConversationHandler.END
    try:
        sh = get_sheet()
        ws = sh.worksheet("Contas")
        cp_id = next_id(ws)
        now = datetime.now(TZ).strftime("%d/%m/%Y %H:%M")
        data = {**context.user_data,
                "ID": cp_id,
                "Status": "Pendente",
                "Lançado por": "Manual",
                "Criado em": now}
        append_conta(ws, data)
        await update.message.reply_text(
            f"✅ Conta *{cp_id}* lançada com sucesso!",
            reply_markup=ReplyKeyboardRemove(),
            parse_mode="Markdown"
        )
    except Exception as e:
        log.error("Erro ao salvar: %s", e)
        await update.message.reply_text(f"❌ Erro ao salvar: {e}", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END

async def cmd_cancelar(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("❌ Operação cancelada.", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END

# ── Comando rápido /cp ────────────────────────────────────────────────────────
async def cmd_cp(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Uso: /cp Credor Valor DD/MM/AAAA Empresa Categoria
    Ex:  /cp Simples_Nacional 1500 20/06/2026 NK_Soluções DAS
    """
    if not is_authorized(update):
        return
    args = context.args
    if len(args) < 5:
        await update.message.reply_text(
            "❌ Formato: `/cp Credor Valor DD/MM/AAAA Empresa Categoria`\n"
            "Ex: `/cp Simples_Nacional 1500 20/06/2026 NK_Soluções DAS`\n"
            "Use _ no lugar de espaços.",
            parse_mode="Markdown"
        )
        return
    credor    = args[0].replace("_", " ")
    valor     = args[1].replace(",", ".")
    venc_str  = args[2]
    empresa   = args[3].replace("_", " ")
    categoria = args[4].replace("_", " ")

    d = parse_date(venc_str)
    if not d:
        await update.message.reply_text("❌ Data inválida. Use DD/MM/AAAA.")
        return
    try:
        sh = get_sheet()
        ws = sh.worksheet("Contas")
        cp_id = next_id(ws)
        now = datetime.now(TZ).strftime("%d/%m/%Y %H:%M")
        data = {
            "ID": cp_id, "Empresa": empresa, "Categoria": categoria,
            "Descrição": f"{categoria} {credor}", "Credor": credor,
            "Valor (R$)": valor, "Vencimento": d.strftime("%d/%m/%Y"),
            "Status": "Pendente", "Recorrente": "Não", "Frequência": "Única",
            "Lançado por": "Comando", "Criado em": now
        }
        append_conta(ws, data)
        await update.message.reply_text(
            f"✅ *{cp_id}* lançado!\n"
            f"{credor} | {fmt_brl(valor)} | {d.strftime('%d/%m/%Y')} | {empresa}",
            parse_mode="Markdown"
        )
    except Exception as e:
        await update.message.reply_text(f"❌ Erro: {e}")

# ── /pagar ────────────────────────────────────────────────────────────────────
async def cmd_pagar(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Uso: /pagar CP0001 [DD/MM/AAAA] [valor_pago]
    """
    if not is_authorized(update):
        return
    args = context.args
    if not args:
        await update.message.reply_text("Uso: `/pagar CP0001` ou `/pagar CP0001 15/06/2026 1500`", parse_mode="Markdown")
        return
    cp_id    = args[0].upper()
    data_pag = args[1] if len(args) > 1 else date.today().strftime("%d/%m/%Y")
    val_pago = args[2] if len(args) > 2 else ""
    try:
        sh = get_sheet()
        ws = sh.worksheet("Contas")
        ok = update_status(ws, cp_id, "Pago", data_pag, val_pago)
        if ok:
            await update.message.reply_text(f"✅ *{cp_id}* marcado como pago em {data_pag}.", parse_mode="Markdown")
        else:
            await update.message.reply_text(f"❌ ID *{cp_id}* não encontrado.", parse_mode="Markdown")
    except Exception as e:
        await update.message.reply_text(f"❌ Erro: {e}")

# ── /listar ───────────────────────────────────────────────────────────────────
async def cmd_listar(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_authorized(update):
        return
    try:
        sh = get_sheet()
        ws = sh.worksheet("Contas")
        contas = get_all_contas(ws)
        pendentes = [c for c in contas if c.get("Status") != "Pago"]
        if not pendentes:
            await update.message.reply_text("✅ Nenhuma conta pendente!")
            return
        pendentes.sort(key=lambda c: parse_date(c.get("Vencimento", "")) or date.max)
        lines = ["📋 *Contas Pendentes*\n"]
        for c in pendentes[:20]:
            d = days_until(c.get("Vencimento", ""))
            em = status_emoji(c.get("Status", ""), d)
            lines.append(
                f"{em} `{c.get('ID','')}` {c.get('Credor','')} | "
                f"{fmt_brl(c.get('Valor (R$)',0))} | {c.get('Vencimento','')} | {c.get('Empresa','')}"
            )
        if len(pendentes) > 20:
            lines.append(f"\n_...e mais {len(pendentes)-20} contas. Use /resumo para visão completa._")
        await update.message.reply_text("\n".join(lines), parse_mode="Markdown")
    except Exception as e:
        await update.message.reply_text(f"❌ Erro: {e}")

# ── Resumos ───────────────────────────────────────────────────────────────────
async def _send_resumo(update, contas, titulo):
    msg = build_resumo(contas, titulo)
    await update.message.reply_text(msg, parse_mode="Markdown")

async def cmd_resumo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_authorized(update):
        return
    sh = get_sheet()
    ws = sh.worksheet("Contas")
    contas = [c for c in get_all_contas(ws) if c.get("Status") != "Pago"]
    await _send_resumo(update, contas, "Resumo Geral")

async def cmd_resumo_dia(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_authorized(update):
        return
    sh = get_sheet()
    ws = sh.worksheet("Contas")
    hoje = date.today().strftime("%d/%m/%Y")
    contas = [c for c in get_all_contas(ws) if c.get("Vencimento") == hoje and c.get("Status") != "Pago"]
    await _send_resumo(update, contas, f"Vencendo hoje — {hoje}")

async def cmd_resumo_semana(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_authorized(update):
        return
    sh = get_sheet()
    ws = sh.worksheet("Contas")
    contas = [c for c in get_all_contas(ws)
              if days_until(c.get("Vencimento", "")) is not None
              and 0 <= days_until(c.get("Vencimento", "")) <= 7
              and c.get("Status") != "Pago"]
    await _send_resumo(update, contas, "Próximos 7 dias")

async def cmd_resumo_mes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_authorized(update):
        return
    sh = get_sheet()
    ws = sh.worksheet("Contas")
    hoje = date.today()
    contas = []
    for c in get_all_contas(ws):
        d = parse_date(c.get("Vencimento", ""))
        if d and d.year == hoje.year and d.month == hoje.month and c.get("Status") != "Pago":
            contas.append(c)
    await _send_resumo(update, contas, f"Mês atual — {hoje.strftime('%m/%Y')}")

# ── Import XLSX ───────────────────────────────────────────────────────────────
async def handle_xlsx(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_authorized(update):
        return
    doc = update.message.document
    if not doc.file_name.endswith(".xlsx"):
        await update.message.reply_text("❌ Envie um arquivo .xlsx.")
        return
    await update.message.reply_text("⏳ Processando arquivo...")
    try:
        f = await context.bot.get_file(doc.file_id)
        path = f"/tmp/{doc.file_name}"
        await f.download_to_drive(path)

        wb = openpyxl.load_workbook(path)
        ws_xl = wb.active
        headers = [str(c.value).strip() if c.value else "" for c in next(ws_xl.iter_rows(min_row=1, max_row=1))]

        sh = get_sheet()
        ws = sh.worksheet("Contas")
        now = datetime.now(TZ).strftime("%d/%m/%Y %H:%M")
        count = 0

        for row in ws_xl.iter_rows(min_row=2, values_only=True):
            if not any(row):
                continue
            r = dict(zip(headers, row))
            cp_id = next_id(ws)
            venc = r.get("Vencimento") or r.get("vencimento") or ""
            if hasattr(venc, "strftime"):
                venc = venc.strftime("%d/%m/%Y")
            data = {
                "ID": cp_id,
                "Empresa":     r.get("Empresa") or r.get("empresa") or "",
                "Categoria":   r.get("Categoria") or r.get("categoria") or "Outros",
                "Descrição":   r.get("Descrição") or r.get("descricao") or r.get("Descricao") or "",
                "Credor":      r.get("Credor") or r.get("credor") or "",
                "Valor (R$)":  str(r.get("Valor") or r.get("Valor (R$)") or "0").replace(",", "."),
                "Vencimento":  str(venc),
                "Status":      r.get("Status") or "Pendente",
                "Recorrente":  r.get("Recorrente") or "Não",
                "Frequência":  r.get("Frequência") or r.get("Frequencia") or "Única",
                "Observação":  r.get("Observação") or r.get("Observacao") or "",
                "Lançado por": "Upload",
                "Criado em":   now,
            }
            append_conta(ws, data)
            count += 1

        await update.message.reply_text(f"✅ {count} conta(s) importada(s) com sucesso!")
    except Exception as e:
        log.error("Erro import xlsx: %s", e)
        await update.message.reply_text(f"❌ Erro ao importar: {e}")

# ── /ajuda ────────────────────────────────────────────────────────────────────
async def cmd_ajuda(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_authorized(update):
        return
    await update.message.reply_text(
        "*📚 Comandos NK Contas a Pagar*\n\n"
        "*Lançamento*\n"
        "/nova — guiado passo a passo\n"
        "`/cp Credor Valor Data Empresa Cat` — rápido\n\n"
        "*Pagamento*\n"
        "`/pagar CP0001` — marca como pago hoje\n"
        "`/pagar CP0001 15/06/2026 1500` — com data e valor\n\n"
        "*Consulta*\n"
        "/listar — pendentes ordenados por vencimento\n"
        "/resumo — visão geral\n"
        "/resumo\\_dia — vence hoje\n"
        "/resumo\\_semana — próximos 7 dias\n"
        "/resumo\\_mes — mês atual\n\n"
        "*Import*\n"
        "Envie um `.xlsx` com colunas: Empresa, Categoria, Descrição, Credor, Valor, Vencimento\n\n"
        "*Alertas automáticos*\n"
        "🟡 7 dias | 🟠 3 dias | 🔴 1 dia | 🚨 no dia\n"
        "📋 Resumo diário às 07h",
        parse_mode="Markdown"
    )

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    # Inicializa planilha
    try:
        sh = get_sheet()
        ensure_sheets(sh)
        log.info("Google Sheets conectado.")
    except Exception as e:
        log.error("Erro ao conectar Sheets: %s", e)

    app = Application.builder().token(BOT_TOKEN).build()

    # ConversationHandler para /nova
    conv = ConversationHandler(
        entry_points=[CommandHandler("nova", cmd_nova)],
        states={
            S_EMPRESA:    [MessageHandler(filters.TEXT & ~filters.COMMAND, step_empresa)],
            S_CATEGORIA:  [MessageHandler(filters.TEXT & ~filters.COMMAND, step_categoria)],
            S_DESCRICAO:  [MessageHandler(filters.TEXT & ~filters.COMMAND, step_descricao)],
            S_CREDOR:     [MessageHandler(filters.TEXT & ~filters.COMMAND, step_credor)],
            S_VALOR:      [MessageHandler(filters.TEXT & ~filters.COMMAND, step_valor)],
            S_VENCIMENTO: [MessageHandler(filters.TEXT & ~filters.COMMAND, step_vencimento)],
            S_RECORRENTE: [MessageHandler(filters.TEXT & ~filters.COMMAND, step_recorrente)],
            S_FREQUENCIA: [MessageHandler(filters.TEXT & ~filters.COMMAND, step_frequencia)],
            S_OBS:        [MessageHandler(filters.TEXT, step_obs)],
            S_CONFIRMA:   [MessageHandler(filters.TEXT & ~filters.COMMAND, step_confirma)],
        },
        fallbacks=[CommandHandler("cancelar", cmd_cancelar)],
    )

    app.add_handler(conv)
    app.add_handler(CommandHandler("start",          cmd_start))
    app.add_handler(CommandHandler("ajuda",          cmd_ajuda))
    app.add_handler(CommandHandler("cp",             cmd_cp))
    app.add_handler(CommandHandler("pagar",          cmd_pagar))
    app.add_handler(CommandHandler("listar",         cmd_listar))
    app.add_handler(CommandHandler("resumo",         cmd_resumo))
    app.add_handler(CommandHandler("resumo_dia",     cmd_resumo_dia))
    app.add_handler(CommandHandler("resumo_semana",  cmd_resumo_semana))
    app.add_handler(CommandHandler("resumo_mes",     cmd_resumo_mes))
    app.add_handler(MessageHandler(filters.Document.FileExtension("xlsx"), handle_xlsx))

    # Jobs agendados
    job_queue = app.job_queue

    # Alertas a cada hora (verifica se é dia de alertar)
    job_queue.run_repeating(job_alertas, interval=3600, first=10)

    # Resumo diário às 07h horário de Brasília
    job_queue.run_daily(
        job_resumo_diario,
        time=datetime.strptime("07:00", "%H:%M").replace(tzinfo=TZ).timetz()
    )

    log.info("Bot NK Contas a Pagar iniciado.")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()

import os
import json
import logging
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
from zoneinfo import ZoneInfo

from flask import Flask, request, jsonify
from flask_cors import CORS
import gspread
from google.oauth2.service_account import Credentials

# ── Config ────────────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app, origins=[
    "https://leobdiasg4-maker.github.io",
    "http://localhost:*",
    "http://127.0.0.1:*"
])

SHEET_ID = os.environ["SHEET_ID"]
TZ = ZoneInfo("America/Sao_Paulo")

# Defaults caso Config ainda não exista
DEFAULT_CONFIG = {
    "Empresa":  ["NK Soluções", "NK Pré-Moldados"],
    "Categoria":["DAS","INSS","FGTS","Folha","Fornecedor","Aluguel","Empréstimo","Honorários","Outros"],
    "Conta":    ["Nubank PJ","Sicoob","C6","PagBank","Santander NK Soluções","Santander NK Pré-Moldados","Cora","Dinheiro"],
    "Status":   ["Pendente","Pago","Atrasado","Parcial"],
}

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
    return gc.open_by_key(SHEET_ID)

def ws(name):
    return get_sheet().worksheet(name)

def ensure_config(sh):
    """Garante aba Config no formato vertical Tipo|Valor."""
    existing = [w.title for w in sh.worksheets()]
    if "Config" not in existing:
        w = sh.add_worksheet("Config", rows=200, cols=2)
        rows = [["Tipo", "Valor"]]
        for tipo, valores in DEFAULT_CONFIG.items():
            for v in valores:
                rows.append([tipo, v])
        w.update("A1", rows)
        return w
    # Migra formato antigo (colunas lado a lado) para vertical se necessário
    w = sh.worksheet("Config")
    vals = w.get_all_values()
    if vals and vals[0] == ["Tipo", "Valor"]:
        return w  # já está no formato correto
    # Migração
    rows = [["Tipo", "Valor"]]
    for tipo, valores in DEFAULT_CONFIG.items():
        for v in valores:
            rows.append([tipo, v])
    w.clear()
    w.update("A1", rows)
    return w

def get_config_values(tipo: str) -> list:
    try:
        sh = get_sheet()
        ensure_config(sh)
        w = sh.worksheet("Config")
        rows = w.get_all_values()[1:]  # pula cabeçalho
        return [r[1] for r in rows if len(r) >= 2 and r[0] == tipo and r[1]]
    except Exception as e:
        log.error("get_config_values %s: %s", tipo, e)
        return DEFAULT_CONFIG.get(tipo, [])

def next_cp_id(w) -> str:
    ids = [int(v.replace("CP","")) for v in w.col_values(1)[1:] if v.startswith("CP")]
    return f"CP{(max(ids)+1 if ids else 1):04d}"

def parse_date_br(s: str):
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s.strip(), fmt).date()
        except ValueError:
            continue
    return None

def next_vencimento(venc_str: str, freq: str) -> str:
    d = parse_date_br(venc_str)
    if not d:
        return ""
    if freq == "Mensal":
        d = d + relativedelta(months=1)
    elif freq == "Semanal":
        d = d + relativedelta(weeks=1)
    elif freq == "Anual":
        d = d + relativedelta(years=1)
    return d.strftime("%d/%m/%Y")

# ── Health ────────────────────────────────────────────────────────────────────
@app.route("/")
def health():
    return jsonify({"status": "ok", "service": "NK Contas API"})

@app.route("/config-publica", methods=["GET"])
def config_publica():
    """Retorna configurações públicas para o frontend."""
    return jsonify({
        "sheetId": os.environ.get("SHEET_ID", ""),
        "apiKey":  os.environ.get("GOOGLE_API_KEY", ""),
    })

# ── Config ────────────────────────────────────────────────────────────────────
@app.route("/config", methods=["GET"])
def get_config():
    """Retorna todas as listas de configuração."""
    try:
        sh = get_sheet()
        ensure_config(sh)
        w = sh.worksheet("Config")
        rows = w.get_all_values()[1:]
        result = {}
        for r in rows:
            if len(r) >= 2 and r[0] and r[1]:
                result.setdefault(r[0], []).append(r[1])
        return jsonify(result)
    except Exception as e:
        log.error("get_config: %s", e)
        return jsonify(DEFAULT_CONFIG)

@app.route("/config", methods=["POST"])
def add_config():
    """Adiciona um valor a uma lista."""
    try:
        body  = request.json
        tipo  = body["tipo"]
        valor = body["valor"].strip()
        if not tipo or not valor:
            return jsonify({"error": "tipo e valor são obrigatórios"}), 400
        sh = get_sheet()
        ensure_config(sh)
        w = sh.worksheet("Config")
        # Verifica duplicata
        rows = w.get_all_values()[1:]
        if any(r[0] == tipo and r[1].lower() == valor.lower() for r in rows if len(r) >= 2):
            return jsonify({"error": "Valor já existe"}), 409
        w.append_row([tipo, valor])
        return jsonify({"ok": True})
    except Exception as e:
        log.error("add_config: %s", e)
        return jsonify({"error": str(e)}), 500

@app.route("/config", methods=["DELETE"])
def del_config():
    """Remove um valor de uma lista."""
    try:
        body  = request.json
        tipo  = body["tipo"]
        valor = body["valor"]
        sh = get_sheet()
        w = sh.worksheet("Config")
        rows = w.get_all_values()
        for i, r in enumerate(rows[1:], start=2):
            if len(r) >= 2 and r[0] == tipo and r[1] == valor:
                w.delete_rows(i)
                return jsonify({"ok": True})
        return jsonify({"error": "Não encontrado"}), 404
    except Exception as e:
        log.error("del_config: %s", e)
        return jsonify({"error": str(e)}), 500

@app.route("/config", methods=["PUT"])
def edit_config():
    """Edita um valor existente."""
    try:
        body      = request.json
        tipo      = body["tipo"]
        valor_old = body["valorAntigo"]
        valor_new = body["valorNovo"].strip()
        sh = get_sheet()
        w = sh.worksheet("Config")
        rows = w.get_all_values()
        for i, r in enumerate(rows[1:], start=2):
            if len(r) >= 2 and r[0] == tipo and r[1] == valor_old:
                w.update_cell(i, 2, valor_new)
                return jsonify({"ok": True})
        return jsonify({"error": "Não encontrado"}), 404
    except Exception as e:
        log.error("edit_config: %s", e)
        return jsonify({"error": str(e)}), 500

# ── Usuarios ──────────────────────────────────────────────────────────────────
@app.route("/usuarios", methods=["GET"])
def get_usuarios():
    try:
        w = ws("Usuarios")
        rows = w.get_all_values()
        if len(rows) <= 1:
            return jsonify([])
        data = []
        for i, r in enumerate(rows[1:], start=2):
            data.append({
                "login":    r[0] if len(r) > 0 else "",
                "senha":    r[1] if len(r) > 1 else "",
                "nome":     r[2] if len(r) > 2 else "",
                "rowIndex": i
            })
        return jsonify(data)
    except Exception as e:
        log.error("get_usuarios: %s", e)
        return jsonify({"error": str(e)}), 500

@app.route("/usuarios/senha", methods=["PUT"])
def update_senha():
    try:
        body      = request.json
        row_index = int(body["rowIndex"])
        nova      = body["senha"]
        w = ws("Usuarios")
        w.update_cell(row_index, 2, nova)
        return jsonify({"ok": True})
    except Exception as e:
        log.error("update_senha: %s", e)
        return jsonify({"error": str(e)}), 500

# ── Contas ────────────────────────────────────────────────────────────────────
@app.route("/contas", methods=["POST"])
def criar_conta():
    try:
        body = request.json
        w    = ws("Contas")
        cp_id = next_cp_id(w)
        now  = datetime.now(TZ).strftime("%d/%m/%Y %H:%M")

        row = [
            cp_id,
            body.get("empresa", ""),
            body.get("categoria", ""),
            body.get("descricao", ""),
            body.get("credor", ""),
            float(body.get("valor", 0)),
            body.get("vencimento", ""),
            "Pendente",
            "", "", "",
            body.get("recorrente", "Não"),
            body.get("frequencia", "Única"),
            body.get("obs", ""),
            body.get("lancadoPor", "Web"),
            now
        ]
        w.append_row(row)
        return jsonify({"ok": True, "id": cp_id})
    except Exception as e:
        log.error("criar_conta: %s", e)
        return jsonify({"error": str(e)}), 500

@app.route("/contas/pagar", methods=["PUT"])
def pagar_conta():
    try:
        body        = request.json
        row_index   = int(body["rowIndex"])
        data_pag    = body.get("dataPagamento", "")
        valor_pago  = body.get("valorPago", "")
        conta_usada = body.get("contaBancaria", "")

        w = ws("Contas")
        # Marca como pago
        w.update(f"H{row_index}:K{row_index}", [["Pago", data_pag, valor_pago, conta_usada]])

        # Recorrente — cria próxima parcela
        row_data = w.row_values(row_index)
        recorrente = row_data[11] if len(row_data) > 11 else "Não"
        freq       = row_data[12] if len(row_data) > 12 else "Única"
        vencimento = row_data[6]  if len(row_data) > 6  else ""

        proximo_id = None
        if recorrente == "Sim" and freq != "Única" and vencimento:
            prox_venc = next_vencimento(vencimento, freq)
            if prox_venc:
                novo_id = next_cp_id(w)
                now = datetime.now(TZ).strftime("%d/%m/%Y %H:%M")
                nova_row = [
                    novo_id,
                    row_data[1] if len(row_data) > 1 else "",   # empresa
                    row_data[2] if len(row_data) > 2 else "",   # categoria
                    row_data[3] if len(row_data) > 3 else "",   # descrição
                    row_data[4] if len(row_data) > 4 else "",   # credor
                    row_data[5] if len(row_data) > 5 else "",   # valor
                    prox_venc,
                    "Pendente",
                    "", "", "",
                    recorrente, freq,
                    row_data[13] if len(row_data) > 13 else "", # obs
                    "Auto-recorrente",
                    now
                ]
                w.append_row(nova_row)
                proximo_id = novo_id

        return jsonify({"ok": True, "proximoId": proximo_id})
    except Exception as e:
        log.error("pagar_conta: %s", e)
        return jsonify({"error": str(e)}), 500

@app.route("/contas/cancelar", methods=["PUT"])
def cancelar_conta():
    try:
        body      = request.json
        row_index = int(body["rowIndex"])
        motivo    = body.get("motivo", "").strip()
        if not motivo:
            return jsonify({"error": "Motivo é obrigatório"}), 400
        w = ws("Contas")
        # Coluna Q (17) = Cancelado, Coluna R (18) = Motivo Cancelamento
        w.update(f"Q{row_index}:R{row_index}", [["Sim", motivo]])
        return jsonify({"ok": True})
    except Exception as e:
        log.error("cancelar_conta: %s", e)
        return jsonify({"error": str(e)}), 500

# ── Main ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)

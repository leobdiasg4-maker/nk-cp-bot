import os
import json
import logging
from datetime import datetime
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

# ── Health ────────────────────────────────────────────────────────────────────
@app.route("/")
def health():
    return jsonify({"status": "ok", "service": "NK Contas API"})

# ── Usuarios ──────────────────────────────────────────────────────────────────
@app.route("/usuarios", methods=["GET"])
def get_usuarios():
    try:
        w = ws("Usuarios")
        rows = w.get_all_values()
        if len(rows) <= 1:
            return jsonify([])
        headers = rows[0]
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

        # Gera próximo ID
        ids  = [int(v.replace("CP","")) for v in w.col_values(1)[1:] if v.startswith("CP")]
        cp_id = f"CP{(max(ids)+1 if ids else 1):04d}"
        now  = datetime.now(TZ).strftime("%d/%m/%Y %H:%M")

        row = [
            cp_id,
            body.get("empresa", ""),
            body.get("categoria", ""),
            body.get("descricao", ""),
            body.get("credor", ""),
            body.get("valor", ""),
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
        body      = request.json
        row_index = int(body["rowIndex"])
        data_pag  = body.get("dataPagamento", "")
        valor_pago = body.get("valorPago", "")
        conta_usada = body.get("contaBancaria", "")

        w = ws("Contas")
        w.update(f"H{row_index}:K{row_index}", [["Pago", data_pag, valor_pago, conta_usada]])
        return jsonify({"ok": True})
    except Exception as e:
        log.error("pagar_conta: %s", e)
        return jsonify({"error": str(e)}), 500

# ── Main ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)

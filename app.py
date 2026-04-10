import os
import json
from datetime import datetime

from flask import Flask, request, abort, jsonify, render_template
from flask_cors import CORS

from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import (
    MessageEvent, TextMessage, TextSendMessage,
    TemplateSendMessage, ButtonsTemplate,
    MessageTemplateAction, CarouselTemplate,
    CarouselColumn, URIAction
)

import gspread
from google.oauth2.service_account import Credentials
from google.auth.transport import requests as google_requests
from google.oauth2 import id_token as google_id_token


# =========================
# Flask
# =========================
app = Flask(__name__)
CORS(app)

# =========================
# ENV
# =========================
LINE_CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET", "")
LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN", "")
GOOGLE_CREDENTIALS_JSON = os.getenv("GOOGLE_CREDENTIALS_JSON", "")
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "")
LIFF_ID = os.getenv("LIFF_ID", "")
PAGE_SIZE = 10

if not GOOGLE_CREDENTIALS_JSON:
    raise ValueError("缺少環境變數 GOOGLE_CREDENTIALS_JSON")
if not GOOGLE_SHEET_ID:
    raise ValueError("缺少環境變數 GOOGLE_SHEET_ID")

# =========================
# LINE BOT
# =========================
line_bot_api = None
handler = None
if LINE_CHANNEL_SECRET and LINE_CHANNEL_ACCESS_TOKEN:
    line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
    handler = WebhookHandler(LINE_CHANNEL_SECRET)

# =========================
# Google Sheet
# =========================
scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

creds_info = json.loads(GOOGLE_CREDENTIALS_JSON)
credentials = Credentials.from_service_account_info(creds_info, scopes=scope)
gc = gspread.authorize(credentials)
spreadsheet = gc.open_by_key(GOOGLE_SHEET_ID)
sheet = spreadsheet.sheet1

user_states = {}
user_temp_data = {}


@app.route("/ping")
def ping():
    return jsonify({
        "ok": True,
        "message": "pong"
    }), 200


@app.route("/health")
def health():
    try:
        headers = sheet.row_values(1)
        return jsonify({
            "ok": True,
            "message": "OK",
            "headers": headers
        })
    except Exception as e:
        return jsonify({"ok": False, "message": str(e)}), 500


@app.route("/")
def home():
    return jsonify({
        "ok": True,
        "message": "LINE BOT + LIFF Inventory Running"
    })


@app.route("/liff")
def liff_page():
    return render_template("liff_inventory_mobile_full.html")


def to_int(v):
    try:
        return int(float(str(v)))
    except:
        return 0


def find_rows(keyword):
    data = sheet.get_all_records()
    result = []
    keyword = str(keyword).lower()

    for idx, row in enumerate(data, start=2):
        name = str(row.get("品名", ""))
        size = str(row.get("尺寸", ""))

        if keyword in name.lower() or keyword in size.lower():
            result.append({
                "row_number": idx,
                "品名": name,
                "尺寸": size,
                "數量": to_int(row.get("數量", 0)),
                "位置": row.get("位置", "")
            })
    return result


@app.get("/api/search")
def api_search():
    keyword = request.args.get("keyword", "")
    items = find_rows(keyword)
    return jsonify({"ok": True, "rows": items})


@app.get("/api/stock")
def api_stock():
    data = sheet.get_all_records()
    rows = []
    for idx, row in enumerate(data, start=2):
        rows.append({
            "row_number": idx,
            "品名": str(row.get("品名", "")),
            "尺寸": str(row.get("尺寸", "")),
            "數量": to_int(row.get("數量", 0)),
            "位置": str(row.get("位置", ""))
        })
    return jsonify({"ok": True, "rows": rows})


@app.post("/api/in")
def api_in():
    body = request.get_json(silent=True) or {}
    row = int(body.get("row_number", 0))
    qty = int(body.get("qty", 0))

    if row < 2 or qty <= 0:
        return jsonify({"ok": False, "message": "參數錯誤"}), 400

    current = to_int(sheet.cell(row, 3).value)
    new_qty = current + qty
    sheet.update_cell(row, 3, new_qty)

    return jsonify({"ok": True, "new_qty": new_qty})


@app.post("/api/out")
def api_out():
    body = request.get_json(silent=True) or {}
    row = int(body.get("row_number", 0))
    qty = int(body.get("qty", 0))

    if row < 2 or qty <= 0:
        return jsonify({"ok": False, "message": "參數錯誤"}), 400

    current = to_int(sheet.cell(row, 3).value)
    if qty > current:
        return jsonify({"ok": False, "message": f"出庫失敗，目前庫存只有 {current}"}), 400

    new_qty = current - qty
    sheet.update_cell(row, 3, new_qty)

    return jsonify({"ok": True, "new_qty": new_qty})


@app.post("/api/manual-in")
def api_manual_in():
    body = request.get_json(silent=True) or {}

    name = str(body.get("name", "")).strip()
    size = str(body.get("size", "")).strip()
    qty = int(body.get("qty", 0))
    loc = str(body.get("loc", "")).strip()

    if not name or qty <= 0:
        return jsonify({"ok": False, "message": "品名與數量必填"}), 400

    sheet.append_row([name, size, qty, loc])
    return jsonify({"ok": True})


@app.route("/callback", methods=["POST"])
def callback():
    if not handler:
        return "OK"

    signature = request.headers.get("X-Line-Signature", "")
    body = request.get_data(as_text=True)

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)

    return "OK"


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

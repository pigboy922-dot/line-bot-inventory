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

# =========================
# ⭐⭐⭐ 新增：PING（防 502 核心）⭐⭐⭐
# =========================
@app.route("/ping")
def ping():
    return jsonify({
        "ok": True,
        "message": "pong"
    }), 200


# =========================
# health（保留）
# =========================
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


# =========================
# 首頁
# =========================
@app.route("/")
def home():
    return jsonify({
        "ok": True,
        "message": "LINE BOT + LIFF Inventory Running"
    })


# =========================
# LIFF
# =========================
@app.route("/liff")
def liff_page():
    return render_template("liff_inventory_mobile_full.html")


# =========================
# 工具
# =========================
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


def get_row(row_number):
    data = sheet.get_all_records()
    for idx, row in enumerate(data, start=2):
        if idx == row_number:
            return row
    return None


# =========================
# API（LIFF 用）
# =========================

@app.get("/api/search")
def api_search():
    keyword = request.args.get("keyword", "")
    items = find_rows(keyword)
    return jsonify({"ok": True, "rows": items})


@app.get("/api/stock")
def api_stock():
    data = sheet.get_all_records()
    return jsonify({"ok": True, "rows": data})


@app.post("/api/in")
def api_in():
    body = request.json
    row = int(body.get("row_number"))
    qty = int(body.get("qty"))

    current = to_int(sheet.cell(row, 3).value)
    new_qty = current + qty

    sheet.update_cell(row, 3, new_qty)

    return jsonify({"ok": True})


@app.post("/api/out")
def api_out():
    body = request.json
    row = int(body.get("row_number"))
    qty = int(body.get("qty"))

    current = to_int(sheet.cell(row, 3).value)
    new_qty = max(0, current - qty)

    sheet.update_cell(row, 3, new_qty)

    return jsonify({"ok": True})


@app.post("/api/manual-in")
def api_manual_in():
    body = request.json

    name = body.get("name")
    size = body.get("size")
    qty = int(body.get("qty"))
    loc = body.get("loc")

    sheet.append_row([name, size, qty, loc])

    return jsonify({"ok": True})


# =========================
# LINE BOT callback（保留）
# =========================
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

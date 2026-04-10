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
# LINE
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
# Log Sheet
# =========================
def ensure_log_worksheet():
    headers = [
        "時間","聊天室類型","群組名稱","群組ID","room_id",
        "user_key","動作","品名","尺寸","原數量",
        "異動數量","新數量","位置","備註"
    ]
    try:
        ws = spreadsheet.worksheet("出入庫紀錄")
        first_row = ws.row_values(1)
        if not first_row:
            ws.update("A1:N1", [headers])
        return ws
    except Exception:
        ws = spreadsheet.add_worksheet(title="出入庫紀錄", rows=2000, cols=14)
        ws.update("A1:N1", [headers])
        return ws

log_sheet = ensure_log_worksheet()

# =========================
# 工具
# =========================
def to_int(value):
    try:
        if value is None or value == "":
            return 0
        return int(float(str(value).strip()))
    except:
        return 0

def get_headers():
    headers = sheet.row_values(1)
    if not headers:
        raise Exception("Google Sheet 第一列沒有表頭")
    return headers

def get_col_index(name):
    headers = get_headers()
    for i, h in enumerate(headers, start=1):
        if str(h).strip() == name:
            return i
    raise Exception(f"找不到欄位：{name}")

def required_columns_ok():
    headers = [str(h).strip() for h in get_headers()]
    required = ["品名","尺寸","數量","位置"]
    return [c for c in required if c not in headers]

# =========================
# API
# =========================
@app.route("/")
def home():
    return jsonify({
        "ok": True,
        "message": "LINE BOT + LIFF Inventory Running",
        "line_bot_enabled": bool(line_bot_api and handler),
        "liff_enabled": True
    })

# ⭐⭐⭐ 重點只在這裡 ⭐⭐⭐
@app.route("/health")
@app.route("/ping")
def ping():
    try:
        missing = required_columns_ok()
        return jsonify({
            "ok": True,
            "message": "OK",
            "sheet_id": GOOGLE_SHEET_ID,
            "missing_columns": missing,
            "line_bot_enabled": bool(line_bot_api and handler),
            "liff_id_configured": bool(LIFF_ID)
        })
    except Exception as e:
        return jsonify({"ok": False, "message": str(e)}), 500

# =========================
# LIFF
# =========================
@app.route("/liff")
def liff_page():
    return render_template("liff_inventory_mobile_full.html")

# =========================
# LINE CALLBACK
# =========================
@app.route("/callback", methods=["POST"])
def callback():
    if not handler:
        return jsonify({"ok": False, "message": "LINE BOT 未設定完成"}), 500

    signature = request.headers.get("X-Line-Signature", "")
    body = request.get_data(as_text=True)

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)

    return "OK"

# =========================
# 啟動
# =========================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

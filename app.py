import os
import json
from datetime import datetime
from flask import Flask, request, abort, jsonify, render_template
from flask_cors import CORS

from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import (
    MessageEvent, TextMessage, TextSendMessage,
    TemplateSendMessage, ButtonsTemplate, MessageTemplateAction,
    CarouselTemplate, CarouselColumn, URIAction
)

import gspread
from google.oauth2.service_account import Credentials
from google.auth.transport import requests as google_requests
from google.oauth2 import id_token as google_id_token

app = Flask(__name__)
CORS(app)

# =========================
# 環境變數設定
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
# LINE BOT 初始化
# =========================
line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# =========================
# Google Sheet 初始化
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
# 建立出入庫紀錄工作表
# =========================
def ensure_log_worksheet():
    headers = [
        "時間", "聊天室類型", "群組名稱", "群組ID", "room_id", "user_key",
        "動作", "品名", "尺寸", "原數量", "異動數量", "新數量", "位置", "備註"
    ]
    try:
        ws = spreadsheet.worksheet("出入庫紀錄")
        if not ws.row_values(1):
            ws.update("A1:N1", [headers])
        return ws
    except Exception:
        ws = spreadsheet.add_worksheet(title="出入庫紀錄", rows=2000, cols=14)
        ws.update("A1:N1", [headers])
        return ws

log_sheet = ensure_log_worksheet()

# =========================
# 工具函式
# =========================
def to_int(value):
    try:
        if value is None or value == "":
            return 0
        return int(float(str(value).strip()))
    except Exception:
        return 0

def get_headers():
    headers = sheet.row_values(1)
    if not headers:
        raise Exception("Google Sheet 第一列沒有表頭")
    return headers

def get_col_index(header_name):
    headers = get_headers()
    for i, h in enumerate(headers, start=1):
        if str(h).strip() == header_name:
            return i
    raise Exception(f"找不到欄位：{header_name}")

def required_columns_ok():
    headers = [str(h).strip() for h in get_headers()]
    required = ["品名", "尺寸", "數量", "位置"]
    return [c for c in required if c not in headers]

def find_matching_rows(keyword):
    data = sheet.get_all_records()
    keyword = str(keyword).strip().lower()
    result = []
    for idx, row in enumerate(data, start=2):
        name = str(row.get("品名", "")).strip()
        size = str(row.get("尺寸", "")).strip()
        qty = to_int(row.get("數量", 0))
        loc = str(row.get("位置", "")).strip()
        if keyword in name.lower() or keyword in size.lower():
            result.append({
                "row_number": idx,
                "品名": name,
                "尺寸": size,
                "數量": qty,
                "位置": loc
            })
    return result

def get_item_by_row(row_number):
    data = sheet.get_all_records()
    for idx, row in enumerate(data, start=2):
        if idx == row_number:
            return {
                "row_number": idx,
                "品名": str(row.get("品名", "")).strip(),
                "尺寸": str(row.get("尺寸", "")).strip(),
                "數量": to_int(row.get("數量", 0)),
                "位置": str(row.get("位置", "")).strip()
            }
    return None

# =========================
# 首頁與 PING
# =========================
@app.route("/")
def home():
    return jsonify({
        "ok": True,
        "message": "LINE BOT + LIFF Inventory Running"
    })

@app.route("/ping", methods=["GET"])
def ping():
    try:
        missing = required_columns_ok()
        return jsonify({
            "ok": True,
            "message": "pong",
            "missing_columns": missing,
            "line_bot_enabled": True,
            "liff_id_configured": bool(LIFF_ID),
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }), 200
    except Exception as e:
        return jsonify({"ok": False, "message": str(e)}), 500

# =========================
# LIFF 頁面
# =========================
@app.route("/liff")
def liff_page():
    return render_template("liff_inventory_mobile_full.html")

# =========================
# LINE Webhook
# =========================
@app.route("/callback", methods=["POST"])
def callback():
    signature = request.headers.get("X-Line-Signature")
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return "OK"

# =========================
# 主選單
# =========================
def build_main_menu():
    liff_url = f"https://liff.line.me/{LIFF_ID}"
    return TemplateSendMessage(
        alt_text="塊材管理選單",
        template=ButtonsTemplate(
            title="塊材管理",
            text="請選擇功能",
            actions=[
                MessageTemplateAction(label="查詢庫存", text="查詢庫存"),
                MessageTemplateAction(label="手動入庫", text="手動入庫"),
                URIAction(label="塊材查詢", uri=liff_url),
            ]
        )
    )

# =========================
# LINE 訊息處理
# =========================
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    text = event.message.text.strip()

    if text in ["喚醒", "ping", "Ping", "PING"]:
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="機器人已喚醒並正常運作！")
        )
        return

    if text in ["塊材管理", "menu", "選單"]:
        line_bot_api.reply_message(event.reply_token, build_main_menu())
        return

    if text == "查詢庫存":
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="請輸入品名或尺寸關鍵字")
        )
        return

    # 關鍵字搜尋
    items = find_matching_rows(text)
    if items:
        lines = [
            f"品名:{i['品名']}｜尺寸:{i['尺寸']}｜數量:{i['數量']}｜位置:{i['位置']}"
            for i in items[:10]
        ]
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="\n".join(lines))
        )
    else:
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="找不到符合的庫存資料")
        )

# =========================
# API：搜尋
# =========================
@app.get("/api/search")
def api_search():
    keyword = request.args.get("q", "").strip()
    items = find_matching_rows(keyword) if keyword else []
    return jsonify({"ok": True, "items": items})

# =========================
# 主程式啟動
# =========================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

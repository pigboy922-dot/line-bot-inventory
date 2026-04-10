import os
import json
from datetime import datetime

from flask import Flask, request, jsonify, render_template
from flask_cors import CORS

from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import (
    MessageEvent, TextMessage, TextSendMessage,
    TemplateSendMessage, ButtonsTemplate,
    MessageTemplateAction, URIAction
)

import gspread
from google.oauth2.service_account import Credentials

# =========================
# Flask 初始化
# =========================
app = Flask(__name__)
CORS(app)

# =========================
# 環境變數
# =========================
LINE_CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET")
LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN")
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
GOOGLE_CREDENTIALS_JSON = os.getenv("GOOGLE_CREDENTIALS_JSON")
LIFF_ID = os.getenv("LIFF_ID")

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

credentials_dict = json.loads(GOOGLE_CREDENTIALS_JSON)
credentials = Credentials.from_service_account_info(credentials_dict, scopes=scope)
gc = gspread.authorize(credentials)
sheet = gc.open_by_key(GOOGLE_SHEET_ID).sheet1

# =========================
# 工具函式
# =========================
def to_int(value):
    try:
        return int(value)
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
    missing = [c for c in required if c not in headers]
    return missing


def find_matching_rows(keyword):
    data = sheet.get_all_records()
    keyword = str(keyword).strip().lower()
    result = []
    for idx, row in enumerate(data, start=2):
        if keyword in str(row.get("品名", "")).lower() or \
           keyword in str(row.get("尺寸", "")).lower():
            result.append({
                "row_number": idx,
                "品名": row.get("品名", ""),
                "尺寸": row.get("尺寸", ""),
                "數量": to_int(row.get("數量", 0)),
                "位置": row.get("位置", "")
            })
    return result


def get_item_by_row(row_number):
    data = sheet.get_all_records()
    for idx, row in enumerate(data, start=2):
        if idx == row_number:
            return {
                "row_number": idx,
                "品名": row.get("品名", ""),
                "尺寸": row.get("尺寸", ""),
                "數量": to_int(row.get("數量", 0)),
                "位置": row.get("位置", "")
            }
    return None


# =========================
# LINE 選單
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
                MessageTemplateAction(label="全部庫存", text="全部庫存"),
                URIAction(label="塊材查詢", uri=liff_url),
                MessageTemplateAction(label="手動入庫", text="手動入庫"),
            ]
        )
    )


# =========================
# 首頁
# =========================
@app.route("/")
def home():
    return jsonify({
        "ok": True,
        "message": "LINE Bot Inventory Running",
        "line_bot_enabled": True,
        "liff_enabled": bool(LIFF_ID)
    })


# =========================
# 健康檢查
# =========================
@app.route("/ping", methods=["GET"])
def ping():
    return jsonify({
        "ok": True,
        "message": "pong",
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }), 200


@app.route("/health", methods=["GET"])
def health():
    try:
        missing = required_columns_ok()
        return jsonify({
            "ok": True,
            "message": "OK",
            "missing_columns": missing,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }), 200
    except Exception as e:
        return jsonify({
            "ok": False,
            "message": str(e)
        }), 500


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
        return "Invalid signature", 400
    except Exception as e:
        print("Webhook error:", str(e))
        return "Error", 500

    return "OK"


# =========================
# LINE 訊息處理
# =========================
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    text = event.message.text.strip()

    # 在群組或個人聊天室輸入「塊材查詢」或「塊材管理」時顯示選單
    if text in ["塊材查詢", "塊材管理"]:
        line_bot_api.reply_message(
            event.reply_token,
            build_main_menu()
        )
        return

    if text == "查詢庫存":
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="請輸入品名或尺寸關鍵字")
        )
        return

    if text == "全部庫存":
        data = sheet.get_all_records()
        if not data:
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="目前沒有庫存資料")
            )
            return

        lines = []
        for idx, row in enumerate(data, start=2):
            lines.append(
                f"{idx}. {row.get('品名')} | {row.get('尺寸')} | 數量:{row.get('數量')} | 位置:{row.get('位置')}"
            )

        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="\n".join(lines[:50]))
        )
        return


# =========================
# API：搜尋庫存
# =========================
@app.get("/api/search")
def api_search():
    keyword = request.args.get("keyword", "")
    items = find_matching_rows(keyword)
    return jsonify({"ok": True, "items": items})


# =========================
# 主程式啟動
# =========================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

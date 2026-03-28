import os
import json
from datetime import datetime
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.models import (
    MessageEvent, TextMessage, TextSendMessage,
    TemplateSendMessage, ButtonsTemplate, MessageTemplateAction,
    CarouselTemplate, CarouselColumn
)
from linebot.exceptions import InvalidSignatureError
import gspread
from google.oauth2.service_account import Credentials

app = Flask(__name__)

LINE_CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET")
LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN")
GOOGLE_CREDENTIALS_JSON = os.getenv("GOOGLE_CREDENTIALS_JSON")
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")

if not LINE_CHANNEL_SECRET:
    raise ValueError("缺少環境變數 LINE_CHANNEL_SECRET")
if not LINE_CHANNEL_ACCESS_TOKEN:
    raise ValueError("缺少環境變數 LINE_CHANNEL_ACCESS_TOKEN")
if not GOOGLE_CREDENTIALS_JSON:
    raise ValueError("缺少環境變數 GOOGLE_CREDENTIALS_JSON")
if not GOOGLE_SHEET_ID:
    raise ValueError("缺少環境變數 GOOGLE_SHEET_ID")

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

creds_info = json.loads(GOOGLE_CREDENTIALS_JSON)
credentials = Credentials.from_service_account_info(creds_info, scopes=scope)
gc = gspread.authorize(credentials)

spreadsheet = gc.open_by_key(GOOGLE_SHEET_ID)
sheet = spreadsheet.sheet1

user_state = {}
user_data = {}
PAGE_SIZE = 10


def ensure_log_worksheet(title, headers):
    try:
        ws = spreadsheet.worksheet(title)
        first_row = ws.row_values(1)
        if not first_row:
            ws.update("A1:N1", [headers])
        return ws
    except Exception:
        ws = spreadsheet.add_worksheet(title=title, rows=2000, cols=max(len(headers), 14))
        ws.update("A1:N1", [headers])
        return ws


log_sheet = ensure_log_worksheet(
    "出入庫紀錄",
    ["時間", "聊天室類型", "群組名稱", "群組ID", "room_id", "user_key", "動作", "品名", "尺寸", "原數量", "異動數量", "新數量", "位置", "備註"]
)


# =========================
# 🔥 基本路由
# =========================

@app.route("/", methods=["GET"])
def home():
    return "LINE BOT Running", 200


# ✅ 關鍵：給 cron 用（超輕量）
@app.route("/health")
def health():
    return "OK"


@app.route("/callback", methods=["POST"])
def callback():
    signature = request.headers.get("X-Line-Signature", "")
    body = request.get_data(as_text=True)

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    except Exception as e:
        print("callback error:", e)
        abort(400)

    return "OK"


# =========================
# 工具
# =========================

def get_user_key(event):
    source = event.source
    group_id = getattr(source, "group_id", None)
    room_id = getattr(source, "room_id", None)
    user_id = getattr(source, "user_id", None)

    if group_id and user_id:
        return f"group_{group_id}_user_{user_id}"
    if room_id and user_id:
        return f"room_{room_id}_user_{user_id}"
    if user_id:
        return f"user_{user_id}"
    return "unknown_user"


def reply_text(reply_token, text):
    line_bot_api.reply_message(reply_token, TextSendMessage(text=str(text)[:5000]))


# =========================
# 主訊息處理（簡化保留）
# =========================

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    text = event.message.text.strip()

    if text == "測試":
        reply_text(event.reply_token, "OK 測試成功")
        return

    if text == "全部庫存":
        data = sheet.get_all_records()
        if not data:
            reply_text(event.reply_token, "沒有資料")
            return

        msg = ""
        for i, row in enumerate(data[:10], start=1):
            msg += f"{i}. {row.get('品名','')} / {row.get('數量','')}\n"

        reply_text(event.reply_token, msg)


# =========================
# 啟動
# =========================

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

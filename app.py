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
# Flask
# =========================
app = Flask(__name__)
CORS(app)

# =========================
# ENV
# =========================
LINE_CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET", "")
LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN", "")
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "")
GOOGLE_CREDENTIALS_JSON = os.getenv("GOOGLE_CREDENTIALS_JSON", "")
LIFF_ID = os.getenv("LIFF_ID", "")

if not GOOGLE_SHEET_ID:
    raise ValueError("缺少 GOOGLE_SHEET_ID")

# =========================
# LINE INIT
# =========================
line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# =========================
# GOOGLE SHEET INIT
# =========================
scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

creds_dict = json.loads(GOOGLE_CREDENTIALS_JSON)
credentials = Credentials.from_service_account_info(creds_dict, scopes=scope)
gc = gspread.authorize(credentials)
sheet = gc.open_by_key(GOOGLE_SHEET_ID).sheet1

# =========================
# TOOL
# =========================
def to_int(v):
    try:
        return int(float(v))
    except:
        return 0


def get_headers():
    h = sheet.row_values(1)
    if not h:
        raise Exception("Sheet沒有表頭")
    return h


def get_col(name):
    for i, h in enumerate(get_headers(), start=1):
        if str(h).strip() == name:
            return i
    raise Exception(f"找不到欄位 {name}")


def find(keyword):
    data = sheet.get_all_records()
    keyword = keyword.lower().strip()

    res = []
    for i, r in enumerate(data, start=2):
        if keyword in str(r.get("品名", "")).lower() or keyword in str(r.get("尺寸", "")).lower():
            res.append({
                "row": i,
                "品名": r.get("品名", ""),
                "尺寸": r.get("尺寸", ""),
                "數量": to_int(r.get("數量", 0)),
                "位置": r.get("位置", "")
            })
    return res


# =========================
# LIFF 卡片（你要的重點）
# =========================
def liff_card():
    url = f"https://liff.line.me/{LIFF_ID}"
    return TemplateSendMessage(
        alt_text="LIFF庫存系統",
        template=ButtonsTemplate(
            title="LIFF庫存系統",
            text="點擊開啟塊材查詢",
            actions=[
                URIAction(label="打開塊材查詢", uri=url)
            ]
        )
    )


# =========================
# HOME
# =========================
@app.route("/")
def home():
    return jsonify({"ok": True})


# =========================
# PING（修正版）
# =========================
@app.route("/ping")
def ping():
    return jsonify({
        "ok": True,
        "message": "pong",
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    })


@app.route("/health")
def health():
    return jsonify({
        "ok": True,
        "sheet_ok": True
    })


# =========================
# LIFF PAGE
# =========================
@app.route("/liff")
def liff():
    return render_template("liff_inventory_mobile_full.html")


# =========================
# CALLBACK
# =========================
@app.route("/callback", methods=["POST"])
def callback():
    signature = request.headers.get("X-Line-Signature", "")
    body = request.get_data(as_text=True)

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        return "bad signature", 400

    return "OK"


# =========================
# MESSAGE HANDLER（核心修正）
# =========================
@handler.add(MessageEvent, message=TextMessage)
def handle(event):
    text = event.message.text.strip()

    # ✅ 你要的：喚醒一定進 LIFF
    if text in ["塊材查詢", "塊材管理", "庫存", "查詢"]:
        line_bot_api.reply_message(
            event.reply_token,
            liff_card()
        )
        return

    # fallback
    if text == "ping":
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="pong")
        )
        return


# =========================
# API SEARCH
# =========================
@app.route("/api/search")
def api_search():
    q = request.args.get("q", "")
    return jsonify({"ok": True, "items": find(q)})


# =========================
# RUN
# =========================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

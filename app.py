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
    URIAction
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

if not GOOGLE_CREDENTIALS_JSON:
    raise ValueError("缺少 GOOGLE_CREDENTIALS_JSON")

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

try:
    creds_dict = json.loads(GOOGLE_CREDENTIALS_JSON)
except Exception as e:
    raise ValueError(f"GOOGLE_CREDENTIALS_JSON 格式錯誤: {e}")

credentials = Credentials.from_service_account_info(creds_dict, scopes=scope)
gc = gspread.authorize(credentials)
sheet = gc.open_by_key(GOOGLE_SHEET_ID).sheet1

# =========================
# TOOL
# =========================
def to_int(v):
    try:
        return int(float(v))
    except Exception:
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


def normalize_item(row_index, r):
    return {
        "row": row_index,
        "品名": r.get("品名", ""),
        "尺寸": r.get("尺寸", ""),
        "數量": to_int(r.get("數量", 0)),
        "位置": r.get("位置", ""),
        "備註": r.get("備註", "")
    }


def get_all_items():
    data = sheet.get_all_records()
    res = []
    for i, r in enumerate(data, start=2):
        res.append(normalize_item(i, r))
    return res


def find(keyword):
    data = sheet.get_all_records()
    keyword = str(keyword).lower().strip()

    res = []
    for i, r in enumerate(data, start=2):
        name = str(r.get("品名", "")).lower()
        size = str(r.get("尺寸", "")).lower()

        # keyword 空字串時，回全部
        if keyword == "" or keyword in name or keyword in size:
            res.append(normalize_item(i, r))
    return res

# =========================
# LIFF 卡片
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
# ERROR HANDLER
# =========================
@app.errorhandler(Exception)
def handle_error(e):
    return jsonify({
        "ok": False,
        "error": str(e)
    }), 500

# =========================
# HOME
# =========================
@app.route("/")
def home():
    return jsonify({"ok": True})

# =========================
# PING
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
    try:
        headers = get_headers()
        return jsonify({
            "ok": True,
            "sheet_ok": True,
            "headers": headers
        })
    except Exception as e:
        return jsonify({
            "ok": False,
            "sheet_ok": False,
            "error": str(e)
        }), 500

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
# MESSAGE HANDLER
# =========================
@handler.add(MessageEvent, message=TextMessage)
def handle(event):
    text = event.message.text.strip()

    if text in ["塊材查詢", "塊材管理", "庫存", "查詢"]:
        line_bot_api.reply_message(
            event.reply_token,
            liff_card()
        )
        return

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
# API ALL
# =========================
@app.route("/api/all")
def api_all():
    return jsonify({
        "ok": True,
        "items": get_all_items()
    })

# =========================
# API IN
# =========================
@app.route("/api/in", methods=["POST"])
def api_in():
    data = request.get_json(silent=True) or {}

    name = str(data.get("品名", "")).strip()
    qty = to_int(data.get("數量", 0))

    if not name:
        return jsonify({"ok": False, "msg": "缺少品名"}), 400

    if qty <= 0:
        return jsonify({"ok": False, "msg": "數量需大於0"}), 400

    items = find(name)
    if not items:
        return jsonify({"ok": False, "msg": "找不到品項"}), 404

    row = items[0]["row"]
    col_qty = get_col("數量")

    current = to_int(sheet.cell(row, col_qty).value)
    new_qty = current + qty
    sheet.update_cell(row, col_qty, new_qty)

    return jsonify({
        "ok": True,
        "msg": "入庫成功",
        "row": row,
        "品名": name,
        "new_qty": new_qty
    })

# =========================
# API OUT
# =========================
@app.route("/api/out", methods=["POST"])
def api_out():
    data = request.get_json(silent=True) or {}

    name = str(data.get("品名", "")).strip()
    qty = to_int(data.get("數量", 0))

    if not name:
        return jsonify({"ok": False, "msg": "缺少品名"}), 400

    if qty <= 0:
        return jsonify({"ok": False, "msg": "數量需大於0"}), 400

    items = find(name)
    if not items:
        return jsonify({"ok": False, "msg": "找不到品項"}), 404

    row = items[0]["row"]
    col_qty = get_col("數量")

    current = to_int(sheet.cell(row, col_qty).value)
    if current < qty:
        return jsonify({"ok": False, "msg": "庫存不足"}), 400

    new_qty = current - qty
    sheet.update_cell(row, col_qty, new_qty)

    return jsonify({
        "ok": True,
        "msg": "出庫成功",
        "row": row,
        "品名": name,
        "new_qty": new_qty
    })

# =========================
# RUN
# =========================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

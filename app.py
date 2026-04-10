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


app = Flask(__name__)
CORS(app)

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

line_bot_api = None
handler = None
if LINE_CHANNEL_SECRET and LINE_CHANNEL_ACCESS_TOKEN:
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

user_states = {}
user_temp_data = {}


@app.route("/ping")
def ping():
    return jsonify({"ok": True, "message": "pong"}), 200


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
        name = str(row.get("品名", "")).strip()
        size = str(row.get("尺寸", "")).strip()
        qty = row.get("數量", 0)
        loc = str(row.get("位置", "")).strip()
        if keyword in name.lower() or keyword in size.lower():
            result.append({
                "row_number": idx,
                "品名": name,
                "尺寸": size,
                "數量": to_int(qty),
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


@app.route("/")
def home():
    return jsonify({
        "ok": True,
        "message": "LINE BOT + LIFF Inventory Running",
        "line_bot_enabled": bool(line_bot_api and handler),
        "liff_enabled": True
    })


@app.route("/health")
def health():
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


@app.route("/liff")
def liff_page():
    return render_template("liff_inventory_mobile_full.html")


@app.get("/api/search")
def api_search():
    try:
        missing = required_columns_ok()
        if missing:
            return jsonify({
                "ok": False,
                "message": f"Google Sheet 缺少欄位：{', '.join(missing)}"
            }), 500

        q = request.args.get("q", "").strip()
        if not q:
            return jsonify({"ok": True, "items": []})

        items = find_matching_rows(q)
        return jsonify({"ok": True, "items": items})
    except Exception as e:
        return jsonify({"ok": False, "message": str(e)}), 500


@app.get("/api/stock")
def api_stock():
    try:
        missing = required_columns_ok()
        if missing:
            return jsonify({
                "ok": False,
                "message": f"Google Sheet 缺少欄位：{', '.join(missing)}"
            }), 500

        page = max(1, int(request.args.get("page", 1)))
        page_size = max(1, min(50, int(request.args.get("page_size", PAGE_SIZE))))

        data = sheet.get_all_records()
        total_count = len(data)
        total_pages = max(1, (total_count + page_size - 1) // page_size)

        if page > total_pages:
            page = total_pages

        start_idx = (page - 1) * page_size
        end_idx = start_idx + page_size
        rows = data[start_idx:end_idx]

        items = []
        for idx, row in enumerate(rows, start=start_idx + 2):
            items.append({
                "row_number": idx,
                "品名": str(row.get("品名", "")).strip(),
                "尺寸": str(row.get("尺寸", "")).strip(),
                "數量": to_int(row.get("數量", 0)),
                "位置": str(row.get("位置", "")).strip()
            })

        return jsonify({
            "ok": True,
            "page": page,
            "page_size": page_size,
            "total_pages": total_pages,
            "total_count": total_count,
            "items": items
        })
    except Exception as e:
        return jsonify({"ok": False, "message": str(e)}), 500


@app.post("/api/in")
def api_in():
    try:
        missing = required_columns_ok()
        if missing:
            return jsonify({
                "ok": False,
                "message": f"Google Sheet 缺少欄位：{', '.join(missing)}"
            }), 500

        body = request.get_json(silent=True) or {}
        row_number = int(body.get("row_number", 0))
        qty = int(body.get("qty", 0))

        if row_number < 2 or qty <= 0:
            return jsonify({"ok": False, "message": "參數錯誤"}), 400

        item = get_item_by_row(row_number)
        if not item:
            return jsonify({"ok": False, "message": "找不到該筆資料"}), 404

        current = item["數量"]
        new_qty = current + qty
        sheet.update_cell(row_number, 3, new_qty)

        return jsonify({
            "ok": True,
            "message": "入庫成功",
            "item": {
                **item,
                "數量": new_qty
            }
        })
    except Exception as e:
        return jsonify({"ok": False, "message": str(e)}), 500


@app.post("/api/out")
def api_out():
    try:
        missing = required_columns_ok()
        if missing:
            return jsonify({
                "ok": False,
                "message": f"Google Sheet 缺少欄位：{', '.join(missing)}"
            }), 500

        body = request.get_json(silent=True) or {}
        row_number = int(body.get("row_number", 0))
        qty = int(body.get("qty", 0))

        if row_number < 2 or qty <= 0:
            return jsonify({"ok": False, "message": "參數錯誤"}), 400

        item = get_item_by_row(row_number)
        if not item:
            return jsonify({"ok": False, "message": "找不到該筆資料"}), 404

        current = item["數量"]
        if qty > current:
            return jsonify({
                "ok": False,
                "message": f"出庫失敗，目前庫存只有 {current}"
            }), 400

        new_qty = current - qty
        sheet.update_cell(row_number, 3, new_qty)

        return jsonify({
            "ok": True,
            "message": "出庫成功",
            "item": {
                **item,
                "數量": new_qty
            }
        })
    except Exception as e:
        return jsonify({"ok": False, "message": str(e)}), 500


@app.post("/api/manual-in")
def api_manual_in():
    try:
        missing = required_columns_ok()
        if missing:
            return jsonify({
                "ok": False,
                "message": f"Google Sheet 缺少欄位：{', '.join(missing)}"
            }), 500

        body = request.get_json(silent=True) or {}
        name = str(body.get("name", "")).strip()
        size = str(body.get("size", "")).strip()
        qty = int(body.get("qty", 0))
        location = str(body.get("location", "")).strip()

        if not name or qty <= 0:
            return jsonify({"ok": False, "message": "品名與數量必填"}), 400

        sheet.append_row([name, size, qty, location])

        return jsonify({
            "ok": True,
            "message": "手動入庫成功"
        })
    except Exception as e:
        return jsonify({"ok": False, "message": str(e)}), 500


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


@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    text = event.message.text.strip()
    if text == "塊材管理":
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="塊材管理功能正常")
        )
        return


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

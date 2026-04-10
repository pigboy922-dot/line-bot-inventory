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

def ensure_log_worksheet():
    headers = [
        "時間", "聊天室類型", "群組名稱", "群組ID", "room_id",
        "user_key", "動作", "品名", "尺寸", "原數量",
        "異動數量", "新數量", "位置", "備註"
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

def verify_liff_id_token(raw_token):
    if not raw_token:
        return {"ok": False, "message": "缺少 LIFF ID Token"}
    if not LIFF_ID:
        return {"ok": False, "message": "伺服器未設定 LIFF_ID"}
    try:
        payload = google_id_token.verify_oauth2_token(
            raw_token,
            google_requests.Request(),
            audience=LIFF_ID
        )
        return {
            "ok": True,
            "sub": payload.get("sub", ""),
            "name": payload.get("name", ""),
            "picture": payload.get("picture", "")
        }
    except Exception as e:
        return {"ok": False, "message": f"LIFF Token 驗證失敗：{str(e)}"}

def get_actor_info():
    token = request.headers.get("X-LIFF-ID-Token", "").strip()
    verify = verify_liff_id_token(token) if token else {"ok": False}
    body = request.get_json(silent=True) or {}
    return {
        "line_user_id": body.get("line_user_id", "") or verify.get("sub", ""),
        "line_name": body.get("line_name", "") or verify.get("name", ""),
        "verify": verify
    }

def get_user_key(event):
    source = event.source
    if source.type == "user":
        return f"user:{source.user_id}"
    if source.type == "group":
        return f"group:{source.group_id}:user:{source.user_id}"
    if source.type == "room":
        return f"room:{source.room_id}:user:{source.user_id}"
    return "unknown"

# ...（後面邏輯完全照你原本內容，包含 callback / handle_message / API / main）

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

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
LINE_CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET", "").strip()
LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN", "").strip()
GOOGLE_CREDENTIALS_JSON = os.getenv("GOOGLE_CREDENTIALS_JSON", "").strip()
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "").strip()
LIFF_ID = os.getenv("LIFF_ID", "").strip()
LIFF_CHANNEL_ID = os.getenv("LIFF_CHANNEL_ID", "").strip()
APP_BASE_URL = os.getenv("APP_BASE_URL", "").strip().rstrip("/")

PAGE_SIZE = 10


# =========================
# LINE SDK
# =========================
line_bot_api = None
handler = None

if LINE_CHANNEL_SECRET and LINE_CHANNEL_ACCESS_TOKEN:
    try:
        line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
        handler = WebhookHandler(LINE_CHANNEL_SECRET)
    except Exception as e:
        print(f"[LINE INIT ERROR] {e}")
        line_bot_api = None
        handler = None


# =========================
# Google Sheet Lazy Init
# =========================
_scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

_gc = None
_spreadsheet = None
_sheet = None
_log_sheet = None


def get_liff_url():
    if LIFF_ID:
        return f"https://liff.line.me/{LIFF_ID}"
    if APP_BASE_URL:
        return f"{APP_BASE_URL}/liff"
    return "/liff"


def init_gspread():
    global _gc, _spreadsheet, _sheet

    if _gc and _spreadsheet and _sheet:
        return _gc, _spreadsheet, _sheet

    if not GOOGLE_CREDENTIALS_JSON:
        raise ValueError("缺少環境變數 GOOGLE_CREDENTIALS_JSON")

    if not GOOGLE_SHEET_ID:
        raise ValueError("缺少環境變數 GOOGLE_SHEET_ID")

    try:
        creds_info = json.loads(GOOGLE_CREDENTIALS_JSON)
    except Exception as e:
        raise ValueError(f"GOOGLE_CREDENTIALS_JSON 格式錯誤：{e}")

    credentials = Credentials.from_service_account_info(creds_info, scopes=_scope)
    gc = gspread.authorize(credentials)
    spreadsheet = gc.open_by_key(GOOGLE_SHEET_ID)
    sheet = spreadsheet.sheet1

    _gc = gc
    _spreadsheet = spreadsheet
    _sheet = sheet
    return _gc, _spreadsheet, _sheet


def get_spreadsheet():
    _, spreadsheet, _ = init_gspread()
    return spreadsheet


def get_sheet():
    _, _, sheet = init_gspread()
    return sheet


def ensure_log_worksheet():
    global _log_sheet
    if _log_sheet is not None:
        return _log_sheet

    spreadsheet = get_spreadsheet()
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
    except Exception:
        ws = spreadsheet.add_worksheet(title="出入庫紀錄", rows=2000, cols=14)
        ws.update("A1:N1", [headers])

    _log_sheet = ws
    return _log_sheet


# =========================
# Helpers
# =========================
user_states = {}
user_temp_data = {}


def to_int(value):
    try:
        if value is None or value == "":
            return 0
        return int(float(str(value).strip()))
    except Exception:
        return 0


def get_headers():
    sheet = get_sheet()
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
    sheet = get_sheet()
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
    sheet = get_sheet()
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

    audiences = []

    if LIFF_CHANNEL_ID:
        audiences.append(LIFF_CHANNEL_ID.strip())

    if LIFF_ID:
        liff_id = LIFF_ID.strip()
        if liff_id not in audiences:
            audiences.append(liff_id)

    if not audiences:
        return {"ok": False, "message": "未設定 LIFF_CHANNEL_ID 或 LIFF_ID"}

    last_error = None

    for aud in audiences:
        try:
            payload = google_id_token.verify_oauth2_token(
                raw_token,
                google_requests.Request(),
                audience=aud
            )
            return {
                "ok": True,
                "sub": payload.get("sub", ""),
                "name": payload.get("name", ""),
                "picture": payload.get("picture", ""),
                "aud": payload.get("aud", ""),
                "used_audience": aud
            }
        except Exception as e:
            last_error = str(e)

    return {
        "ok": False,
        "message": f"LIFF Token 驗證失敗：{last_error}"
    }


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


def get_chatroom_info(event):
    source = event.source
    info = {
        "聊天室類型": source.type,
        "群組名稱": "",
        "群組ID": "",
        "room_id": "",
        "user_key": get_user_key(event)
    }

    if source.type == "group":
        info["群組ID"] = getattr(source, "group_id", "")
        if line_bot_api:
            try:
                summary = line_bot_api.get_group_summary(source.group_id)
                info["群組名稱"] = getattr(summary, "group_name", "") or ""
            except Exception:
                info["群組名稱"] = ""
    elif source.type == "room":
        info["room_id"] = getattr(source, "room_id", "")

    return info


def log_inventory_action_line(event, action, item=None, old_qty="", change_qty="", new_qty="", note=""):
    try:
        log_sheet = ensure_log_worksheet()
        chat = get_chatroom_info(event)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        name = ""
        size = ""
        loc = ""
        if item:
            name = str(item.get("品名", "")).strip()
            size = str(item.get("尺寸", "")).strip()
            loc = str(item.get("位置", "")).strip()

        log_sheet.append_row([
            now, chat["聊天室類型"], chat["群組名稱"], chat["群組ID"],
            chat["room_id"], chat["user_key"], action, name, size,
            old_qty, change_qty, new_qty, loc, note
        ])
    except Exception as e:
        print(f"[LOG LINE ERROR] {e}")


def log_inventory_action_liff(action, item=None, old_qty="", change_qty="", new_qty="", note="", actor=None):
    try:
        log_sheet = ensure_log_worksheet()
        actor = actor or {}
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        name = ""
        size = ""
        loc = ""
        if item:
            name = str(item.get("品名", "")).strip()
            size = str(item.get("尺寸", "")).strip()
            loc = str(item.get("位置", "")).strip()

        log_sheet.append_row([
            now, "liff", "", "", "",
            actor.get("line_user_id", ""), action, name, size,
            old_qty, change_qty, new_qty, loc,
            note or actor.get("line_name", "")
        ])
    except Exception as e:
        print(f"[LOG LIFF ERROR] {e}")


def reset_user_state(user_key):
    user_states.pop(user_key, None)
    user_temp_data.pop(user_key, None)


def build_main_menu():
    liff_url = get_liff_url()
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


def build_liff_open_card():
    liff_url = get_liff_url()
    return TemplateSendMessage(
        alt_text="打開塊材查詢",
        template=ButtonsTemplate(
            title="塊材查詢",
            text="點下面按鈕開啟 LIFF 庫存系統",
            actions=[
                URIAction(label="打開塊材查詢", uri=liff_url)
            ]
        )
    )


def build_search_results_carousel(items, mode="out"):
    columns = []
    for item in items[:10]:
        title = f"{item['品名'][:20]}" or "查詢結果"
        text = f"尺寸:{item['尺寸'][:20]}\n數量:{item['數量']}\n位置:{item['位置'][:20]}"
        actions = []

        if mode == "out":
            actions.append(MessageTemplateAction(
                label="直接出庫",
                text=f"直接出庫::{item['row_number']}"
            ))
        elif mode == "in":
            actions.append(MessageTemplateAction(
                label="直接入庫",
                text=f"直接入庫::{item['row_number']}"
            ))

        actions.append(MessageTemplateAction(label="返回選單", text="返回選單"))

        columns.append(CarouselColumn(
            title=title[:40],
            text=text[:60],
            actions=actions[:3]
        ))

    return TemplateSendMessage(
        alt_text="查詢結果",
        template=CarouselTemplate(columns=columns)
    )


def build_all_stock_carousel(page=1):
    sheet = get_sheet()
    data = sheet.get_all_records()
    total_count = len(data)
    total_pages = max(1, (total_count + PAGE_SIZE - 1) // PAGE_SIZE)

    page = max(1, min(page, total_pages))
    start_idx = (page - 1) * PAGE_SIZE
    end_idx = start_idx + PAGE_SIZE
    rows = data[start_idx:end_idx]

    columns = []
    for idx, row in enumerate(rows, start=start_idx + 2):
        name = str(row.get("品名", "")).strip()
        size = str(row.get("尺寸", "")).strip()
        qty = to_int(row.get("數量", 0))
        loc = str(row.get("位置", "")).strip()

        columns.append(CarouselColumn(
            title=(name[:20] or "未命名"),
            text=f"尺寸:{size[:20]}\n數量:{qty}\n位置:{loc[:20]}",
            actions=[
                MessageTemplateAction(label="直接出庫", text=f"直接出庫::{idx}"),
                MessageTemplateAction(label="直接入庫", text=f"直接入庫::{idx}"),
                MessageTemplateAction(label="返回選單", text="返回選單"),
            ]
        ))

    messages = [
        TextSendMessage(text=f"全部庫存　第 {page}/{total_pages} 頁，共 {total_count} 筆"),
        TemplateSendMessage(
            alt_text="全部庫存",
            template=CarouselTemplate(columns=columns)
        )
    ]

    nav_actions = []
    if page > 1:
        nav_actions.append(MessageTemplateAction(label="上一頁", text=f"全部庫存::{page-1}"))
    if page < total_pages:
        nav_actions.append(MessageTemplateAction(label="下一頁", text=f"全部庫存::{page+1}"))
    nav_actions.append(MessageTemplateAction(label="返回選單", text="返回選單"))

    messages.append(
        TemplateSendMessage(
            alt_text="分頁操作",
            template=ButtonsTemplate(
                title="分頁操作",
                text="請選擇",
                actions=nav_actions[:4]
            )
        )
    )
    return messages


# =========================
# Routes
# =========================
@app.route("/")
def home():
    return jsonify({
        "ok": True,
        "message": "LINE BOT + LIFF Inventory Running",
        "line_bot_enabled": bool(line_bot_api and handler),
        "liff_enabled": True
    })


@app.route("/ping")
def ping():
    return "ok", 200


@app.route("/health")
def health():
    return jsonify({
        "ok": True,
        "message": "server alive",
        "line_bot_enabled": bool(line_bot_api and handler),
        "has_google_credentials": bool(GOOGLE_CREDENTIALS_JSON),
        "has_google_sheet_id": bool(GOOGLE_SHEET_ID),
        "has_liff_id": bool(LIFF_ID),
        "has_liff_channel_id": bool(LIFF_CHANNEL_ID),
    })


@app.route("/healthz-sheet")
def healthz_sheet():
    try:
        missing = required_columns_ok()
        return jsonify({
            "ok": True,
            "message": "sheet connected",
            "sheet_id": GOOGLE_SHEET_ID,
            "missing_columns": missing
        })
    except Exception as e:
        return jsonify({
            "ok": False,
            "message": str(e)
        }), 500


@app.route("/liff")
def liff_page():
    return render_template("liff_inventory_mobile_full.html")


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
    except Exception as e:
        print(f"[CALLBACK ERROR] {e}")
        return jsonify({"ok": False, "message": str(e)}), 500

    return "OK"


# =========================
# LINE Message Handler
# =========================
def handle_message(event):
    user_key = get_user_key(event)
    text = event.message.text.strip()
    state = user_states.get(user_key, "")

    try:
        missing = required_columns_ok()
    except Exception as e:
        if line_bot_api:
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text=f"Google Sheet 連線失敗：{e}")
            )
        return

    if text == "塊材查詢":
        reset_user_state(user_key)
        line_bot_api.reply_message(event.reply_token, build_liff_open_card())
        return

    if missing:
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=f"Google Sheet 缺少欄位：{', '.join(missing)}")
        )
        return

    if text == "取消":
        reset_user_state(user_key)
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text="已取消目前操作"))
        return

    if text == "返回選單":
        reset_user_state(user_key)
        line_bot_api.reply_message(event.reply_token, build_main_menu())
        return

    if text == "塊材管理":
        reset_user_state(user_key)
        line_bot_api.reply_message(event.reply_token, build_main_menu())
        return

    if text == "查詢庫存":
        user_states[user_key] = "waiting_search_keyword"
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="請輸入品名或尺寸關鍵字，例如 KF-0030N 509 BDP-1，只要輸入 509 即可查詢")
        )
        return

    if text == "入庫":
        user_states[user_key] = "waiting_in_keyword"
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="請輸入要入庫的品名或尺寸關鍵字")
        )
        return

    if text == "出庫":
        user_states[user_key] = "waiting_out_keyword"
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="請輸入要出庫的品名或尺寸關鍵字")
        )
        return

    if text == "手動入庫":
        user_states[user_key] = "manual_in_name"
        user_temp_data[user_key] = {}
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text="請輸入品名"))
        return

    if text == "全部庫存":
        reset_user_state(user_key)
        line_bot_api.reply_message(event.reply_token, build_all_stock_carousel(page=1))
        return

    if text.startswith("全部庫存::"):
        try:
            page = int(text.split("::", 1)[1])
        except Exception:
            page = 1
        reset_user_state(user_key)
        line_bot_api.reply_message(event.reply_token, build_all_stock_carousel(page=page))
        return

    if text.startswith("直接出庫::"):
        try:
            row_num = int(text.split("::", 1)[1])
        except Exception:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text="列號格式錯誤"))
            return

        item = get_item_by_row(row_num)
        if not item:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text="找不到該筆資料"))
            return

        user_states[user_key] = "waiting_out_qty"
        user_temp_data[user_key] = {"row_number": row_num, "item": item}
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(
                text=(
                    f"請輸入出庫數量：\n"
                    f"品名：{item['品名']}\n"
                    f"尺寸：{item['尺寸']}\n"
                    f"目前數量：{item['數量']}"
                )
            )
        )
        return

    if text.startswith("直接入庫::"):
        try:
            row_num = int(text.split("::", 1)[1])
        except Exception:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text="列號格式錯誤"))
            return

        item = get_item_by_row(row_num)
        if not item:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text="找不到該筆資料"))
            return

        user_states[user_key] = "waiting_in_qty"
        user_temp_data[user_key] = {"row_number": row_num, "item": item}
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(
                text=(
                    f"請輸入入庫數量：\n"
                    f"品名：{item['品名']}\n"
                    f"尺寸：{item['尺寸']}\n"
                    f"目前數量：{item['數量']}"
                )
            )
        )
        return

    if state == "waiting_search_keyword":
        items = find_matching_rows(text)
        reset_user_state(user_key)

        if not items:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text="找不到符合的庫存資料"))
            return

        lines = []
        for item in items[:20]:
            lines.append(
                f"列號:{item['row_number']}｜品名:{item['品名']}｜尺寸:{item['尺寸']}｜數量:{item['數量']}｜位置:{item['位置']}"
            )

        line_bot_api.reply_message(event.reply_token, TextSendMessage(text="\n".join(lines)))
        return

    if state == "waiting_in_keyword":
        items = find_matching_rows(text)
        if not items:
            reset_user_state(user_key)
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text="找不到符合資料，請重新輸入"))
            return

        user_states[user_key] = "waiting_in_select"
        user_temp_data[user_key] = {"matched_items": items}
        line_bot_api.reply_message(event.reply_token, build_search_results_carousel(items, mode="in"))
        return

    if state == "waiting_out_keyword":
        items = find_matching_rows(text)
        if not items:
            reset_user_state(user_key)
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text="找不到符合資料，請重新輸入"))
            return

        user_states[user_key] = "waiting_out_select"
        user_temp_data[user_key] = {"matched_items": items}
        line_bot_api.reply_message(event.reply_token, build_search_results_carousel(items, mode="out"))
        return

    if state == "waiting_in_qty":
        try:
            add_qty = int(text)
        except Exception:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text="請輸入正整數"))
            return

        if add_qty <= 0:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text="入庫數量必須大於 0"))
            return

        temp = user_temp_data.get(user_key, {})
        row_num = temp.get("row_number")
        item = get_item_by_row(row_num)

        if not item:
            reset_user_state(user_key)
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text="找不到該筆資料"))
            return

        sheet = get_sheet()
        qty_col = get_col_index("數量")
        old_qty = to_int(sheet.cell(row_num, qty_col).value)
        new_qty = old_qty + add_qty
        sheet.update_cell(row_num, qty_col, new_qty)
        item["數量"] = new_qty

        log_inventory_action_line(
            event, "入庫",
            item=item,
            old_qty=old_qty,
            change_qty=add_qty,
            new_qty=new_qty,
            note="LINE BOT 入庫"
        )

        reset_user_state(user_key)
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(
                text=(
                    f"入庫完成\n"
                    f"品名：{item['品名']}\n"
                    f"尺寸：{item['尺寸']}\n"
                    f"原數量：{old_qty}\n"
                    f"入庫：{add_qty}\n"
                    f"新數量：{new_qty}"
                )
            )
        )
        return

    if state == "waiting_out_qty":
        try:
            out_qty = int(text)
        except Exception:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text="請輸入正整數"))
            return

        if out_qty <= 0:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text="出庫數量必須大於 0"))
            return

        temp = user_temp_data.get(user_key, {})
        row_num = temp.get("row_number")
        item = get_item_by_row(row_num)

        if not item:
            reset_user_state(user_key)
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text="找不到該筆資料"))
            return

        sheet = get_sheet()
        qty_col = get_col_index("數量")
        old_qty = to_int(sheet.cell(row_num, qty_col).value)

        if out_qty > old_qty:
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text=f"目前庫存只有 {old_qty}，不能出庫 {out_qty}")
            )
            return

        new_qty = old_qty - out_qty
        sheet.update_cell(row_num, qty_col, new_qty)
        item["數量"] = new_qty

        log_inventory_action_line(
            event, "出庫",
            item=item,
            old_qty=old_qty,
            change_qty=out_qty,
            new_qty=new_qty,
            note="LINE BOT 出庫"
        )

        reset_user_state(user_key)
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(
                text=(
                    f"出庫完成\n"
                    f"品名：{item['品名']}\n"
                    f"尺寸：{item['尺寸']}\n"
                    f"原數量：{old_qty}\n"
                    f"出庫：{out_qty}\n"
                    f"新數量：{new_qty}"
                )
            )
        )
        return

    if state == "manual_in_name":
        user_temp_data[user_key]["品名"] = text
        user_states[user_key] = "manual_in_size"
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text="請輸入尺寸"))
        return

    if state == "manual_in_size":
        user_temp_data[user_key]["尺寸"] = text
        user_states[user_key] = "manual_in_qty"
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text="請輸入數量"))
        return

    if state == "manual_in_qty":
        try:
            qty = int(text)
        except Exception:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text="請輸入正整數"))
            return

        if qty <= 0:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text="數量必須大於 0"))
            return

        user_temp_data[user_key]["數量"] = qty
        user_states[user_key] = "manual_in_loc"
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text="請輸入位置"))
        return

    if state == "manual_in_loc":
        temp = user_temp_data.get(user_key, {})
        name = temp.get("品名", "").strip()
        size = temp.get("尺寸", "").strip()
        qty = temp.get("數量", 0)
        loc = text.strip()

        headers = get_headers()
        new_row = [""] * len(headers)
        new_row[get_col_index("品名") - 1] = name
        new_row[get_col_index("尺寸") - 1] = size
        new_row[get_col_index("數量") - 1] = qty
        new_row[get_col_index("位置") - 1] = loc

        sheet = get_sheet()
        sheet.append_row(new_row)

        item = {"品名": name, "尺寸": size, "數量": qty, "位置": loc}

        log_inventory_action_line(
            event, "手動入庫",
            item=item,
            old_qty=0,
            change_qty=qty,
            new_qty=qty,
            note="LINE BOT 手動入庫"
        )

        reset_user_state(user_key)
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(
                text=(
                    f"手動入庫完成\n"
                    f"品名：{name}\n"
                    f"尺寸：{size}\n"
                    f"數量：{qty}\n"
                    f"位置：{loc}"
                )
            )
        )
        return

    allowed_idle_commands = {
        "塊材管理", "塊材查詢", "查詢庫存", "入庫", "出庫",
        "手動入庫", "全部庫存", "取消", "返回選單"
    }
    if text not in allowed_idle_commands:
        return


if handler is not None:
    handler.add(MessageEvent, message=TextMessage)(handle_message)


# =========================
# LIFF / API
# =========================
@app.get("/api/search")
def api_search():
    try:
        missing = required_columns_ok()
        if missing:
            return jsonify({"ok": False, "message": f"Sheet 缺少欄位：{', '.join(missing)}"}), 400

        keyword = request.args.get("q", "").strip()
        if not keyword:
            return jsonify({"ok": True, "items": []})

        items = find_matching_rows(keyword)
        return jsonify({"ok": True, "items": items[:50]})
    except Exception as e:
        return jsonify({"ok": False, "message": str(e)}), 500


@app.get("/api/stock")
def api_stock():
    try:
        missing = required_columns_ok()
        if missing:
            return jsonify({"ok": False, "message": f"Sheet 缺少欄位：{', '.join(missing)}"}), 400

        try:
            page = max(1, int(request.args.get("page", 1)))
            page_size = max(1, int(request.args.get("page_size", PAGE_SIZE)))
        except Exception:
            return jsonify({"ok": False, "message": "page 或 page_size 格式錯誤"}), 400

        sheet = get_sheet()
        data = sheet.get_all_records()
        total_count = len(data)
        total_pages = max(1, (total_count + page_size - 1) // page_size)
        page = min(page, total_pages)

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
                "message": f"Sheet 缺少欄位：{', '.join(missing)}"
            }), 400

        data = request.get_json(silent=True) or {}
        actor = get_actor_info()

        raw_token = request.headers.get("X-LIFF-ID-Token", "").strip()
        if raw_token and not actor.get("verify", {}).get("ok", False):
            return jsonify({
                "ok": False,
                "message": actor.get("verify", {}).get("message", "LIFF Token 驗證失敗"),
                "actor_verify": actor.get("verify", {})
            }), 401

        try:
            row_num = int(data.get("row_number", 0))
            add_qty = int(data.get("qty", 0))
        except Exception:
            return jsonify({
                "ok": False,
                "message": "row_number 或 qty 格式錯誤"
            }), 400

        if row_num < 2:
            return jsonify({
                "ok": False,
                "message": "row_number 錯誤"
            }), 400

        if add_qty <= 0:
            return jsonify({
                "ok": False,
                "message": "入庫數量必須大於 0"
            }), 400

        item = get_item_by_row(row_num)
        if not item:
            return jsonify({
                "ok": False,
                "message": "找不到該筆資料"
            }), 404

        sheet = get_sheet()
        qty_col = get_col_index("數量")
        old_qty = to_int(sheet.cell(row_num, qty_col).value)
        new_qty = old_qty + add_qty
        sheet.update_cell(row_num, qty_col, new_qty)
        item["數量"] = new_qty

        log_inventory_action_liff(
            "入庫",
            item=item,
            old_qty=old_qty,
            change_qty=add_qty,
            new_qty=new_qty,
            note=f"LIFF入庫 / {actor.get('line_name', '')}",
            actor=actor
        )

        return jsonify({
            "ok": True,
            "message": f"入庫完成，最新數量：{new_qty}",
            "old_qty": old_qty,
            "new_qty": new_qty,
            "item": item,
            "actor_verify": actor.get("verify", {})
        })

    except Exception as e:
        return jsonify({
            "ok": False,
            "message": str(e)
        }), 500


@app.post("/api/out")
def api_out():
    try:
        missing = required_columns_ok()
        if missing:
            return jsonify({
                "ok": False,
                "message": f"Sheet 缺少欄位：{', '.join(missing)}"
            }), 400

        data = request.get_json(silent=True) or {}
        actor = get_actor_info()

        raw_token = request.headers.get("X-LIFF-ID-Token", "").strip()
        if raw_token and not actor.get("verify", {}).get("ok", False):
            return jsonify({
                "ok": False,
                "message": actor.get("verify", {}).get("message", "LIFF Token 驗證失敗"),
                "actor_verify": actor.get("verify", {})
            }), 401

        try:
            row_num = int(data.get("row_number", 0))
            out_qty = int(data.get("qty", 0))
        except Exception:
            return jsonify({
                "ok": False,
                "message": "row_number 或 qty 格式錯誤"
            }), 400

        if row_num < 2:
            return jsonify({
                "ok": False,
                "message": "row_number 錯誤"
            }), 400

        if out_qty <= 0:
            return jsonify({
                "ok": False,
                "message": "出庫數量必須大於 0"
            }), 400

        item = get_item_by_row(row_num)
        if not item:
            return jsonify({
                "ok": False,
                "message": "找不到該筆資料"
            }), 404

        sheet = get_sheet()
        qty_col = get_col_index("數量")
        old_qty = to_int(sheet.cell(row_num, qty_col).value)

        if out_qty > old_qty:
            return jsonify({
                "ok": False,
                "message": f"目前庫存只有 {old_qty}，不能出庫 {out_qty}"
            }), 400

        new_qty = old_qty - out_qty
        sheet.update_cell(row_num, qty_col, new_qty)
        item["數量"] = new_qty

        log_inventory_action_liff(
            "出庫",
            item=item,
            old_qty=old_qty,
            change_qty=out_qty,
            new_qty=new_qty,
            note=f"LIFF出庫 / {actor.get('line_name', '')}",
            actor=actor
        )

        return jsonify({
            "ok": True,
            "message": f"出庫完成，最新數量：{new_qty}",
            "old_qty": old_qty,
            "new_qty": new_qty,
            "item": item,
            "actor_verify": actor.get("verify", {})
        })

    except Exception as e:
        return jsonify({
            "ok": False,
            "message": str(e)
        }), 500


@app.post("/api/manual-in")
def api_manual_in():
    try:
        missing = required_columns_ok()
        if missing:
            return jsonify({"ok": False, "message": f"Sheet 缺少欄位：{', '.join(missing)}"}), 400

        data = request.get_json(silent=True) or {}
        actor = get_actor_info()

        raw_token = request.headers.get("X-LIFF-ID-Token", "").strip()
        if raw_token and not actor.get("verify", {}).get("ok", False):
            return jsonify({
                "ok": False,
                "message": actor.get("verify", {}).get("message", "LIFF Token 驗證失敗"),
                "actor_verify": actor.get("verify", {})
            }), 401

        name = str(data.get("name", "")).strip()
        size = str(data.get("size", "")).strip()
        location = str(data.get("location", "")).strip()

        try:
            qty = int(data.get("qty", 0))
        except Exception:
            return jsonify({"ok": False, "message": "qty 格式錯誤"}), 400

        if not name:
            return jsonify({"ok": False, "message": "請輸入品名"}), 400
        if not size:
            return jsonify({"ok": False, "message": "請輸入尺寸"}), 400
        if qty <= 0:
            return jsonify({"ok": False, "message": "數量必須大於 0"}), 400
        if not location:
            return jsonify({"ok": False, "message": "請輸入位置"}), 400

        headers = get_headers()
        new_row = [""] * len(headers)
        new_row[get_col_index("品名") - 1] = name
        new_row[get_col_index("尺寸") - 1] = size
        new_row[get_col_index("數量") - 1] = qty
        new_row[get_col_index("位置") - 1] = location

        sheet = get_sheet()
        sheet.append_row(new_row)

        item = {"品名": name, "尺寸": size, "數量": qty, "位置": location}

        log_inventory_action_liff(
            "手動入庫",
            item=item,
            old_qty=0,
            change_qty=qty,
            new_qty=qty,
            note=f"LIFF手動入庫 / {actor.get('line_name', '')}",
            actor=actor
        )

        return jsonify({
            "ok": True,
            "message": "手動入庫完成",
            "item": item,
            "actor_verify": actor.get("verify", {})
        })
    except Exception as e:
        return jsonify({"ok": False, "message": str(e)}), 500


# =========================
# Main
# =========================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

import json
import os
from datetime import datetime
from typing import Dict, List, Optional

import gspread
from flask import Flask, abort, request
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage
from oauth2client.service_account import ServiceAccountCredentials

app = Flask(__name__)

CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET", "")
CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN", "")
GOOGLE_SERVICE_ACCOUNT_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "")
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "")
TIMEZONE = os.getenv("APP_TIMEZONE", "Asia/Taipei")

if not CHANNEL_SECRET or not CHANNEL_ACCESS_TOKEN:
    raise ValueError("請先設定 LINE_CHANNEL_SECRET 與 LINE_CHANNEL_ACCESS_TOKEN")
if not GOOGLE_SERVICE_ACCOUNT_JSON or not GOOGLE_SHEET_ID:
    raise ValueError("請先設定 GOOGLE_SERVICE_ACCOUNT_JSON 與 GOOGLE_SHEET_ID")

line_bot_api = LineBotApi(CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(CHANNEL_SECRET)

# 簡易對話狀態，適合單一雲端實例使用
SESSIONS: Dict[str, dict] = {}

INVENTORY_HEADERS = ["品名", "尺寸", "庫存", "位置", "備註", "更新時間"]
LOG_HEADERS = ["時間", "類型", "品名", "尺寸", "數量", "位置", "備註", "使用者"]


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def get_gc():
    info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(info, scope)
    return gspread.authorize(creds)


def get_sheet(name: str, headers: List[str]):
    gc = get_gc()
    sh = gc.open_by_key(GOOGLE_SHEET_ID)
    try:
        ws = sh.worksheet(name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=name, rows=1000, cols=max(10, len(headers) + 2))
        ws.append_row(headers)
    current_headers = ws.row_values(1)
    if current_headers != headers:
        ws.resize(rows=max(ws.row_count, 2), cols=max(ws.col_count, len(headers)))
        ws.update("A1:{}1".format(chr(64 + len(headers))), [headers])
    return ws


def inventory_ws():
    return get_sheet("inventory", INVENTORY_HEADERS)


def log_ws():
    return get_sheet("transactions", LOG_HEADERS)


def get_records() -> List[dict]:
    ws = inventory_ws()
    rows = ws.get_all_records(expected_headers=INVENTORY_HEADERS)
    cleaned = []
    for idx, row in enumerate(rows, start=2):
        row["_row"] = idx
        row["品名"] = str(row.get("品名", "")).strip()
        row["尺寸"] = str(row.get("尺寸", "")).strip()
        row["位置"] = str(row.get("位置", "")).strip()
        row["備註"] = str(row.get("備註", "")).strip()
        try:
            row["庫存"] = int(float(row.get("庫存", 0) or 0))
        except Exception:
            row["庫存"] = 0
        cleaned.append(row)
    return cleaned


def session_key(event) -> str:
    source = event.source
    return getattr(source, "user_id", None) or getattr(source, "group_id", None) or getattr(source, "room_id", None) or "default"


def actor_name(event) -> str:
    source = event.source
    return getattr(source, "user_id", None) or getattr(source, "group_id", None) or "unknown"


def normalize(s: str) -> str:
    return (s or "").strip().lower()


def search_inventory(keyword: str) -> List[dict]:
    keyword_n = normalize(keyword)
    result = []
    for row in get_records():
        text = f"{row['品名']} {row['尺寸']} {row['位置']} {row['備註']}".lower()
        if keyword_n in text:
            result.append(row)
    return result


def distinct_items(rows: List[dict]) -> List[dict]:
    seen = set()
    items = []
    for row in rows:
        key = (row["品名"], row["尺寸"])
        if key not in seen:
            seen.add(key)
            items.append(row)
    return items


def format_search_result(rows: List[dict], keyword: str) -> str:
    if not rows:
        return f"找不到包含【{keyword}】的資料。"
    lines = [f"找到 {len(rows)} 筆【{keyword}】相關資料："]
    for i, row in enumerate(rows[:15], start=1):
        lines.append(
            f"{i}. {row['品名']} / {row['尺寸']} / 庫存:{row['庫存']} / 位置:{row['位置'] or '-'}"
        )
    if len(rows) > 15:
        lines.append(f"…其餘 {len(rows)-15} 筆未顯示")
    return "\n".join(lines)


def set_session(key: str, data: dict):
    SESSIONS[key] = data


def clear_session(key: str):
    SESSIONS.pop(key, None)


def get_session(key: str) -> Optional[dict]:
    return SESSIONS.get(key)


def create_or_update_inventory(item_name: str, size: str, qty_delta: int, location: str, note: str):
    ws = inventory_ws()
    records = get_records()
    match = None
    for row in records:
        if normalize(row["品名"]) == normalize(item_name) and normalize(row["尺寸"]) == normalize(size):
            match = row
            break

    if match:
        new_qty = match["庫存"] + qty_delta
        if new_qty < 0:
            raise ValueError(f"庫存不足，目前庫存 {match['庫存']}，無法出庫 {abs(qty_delta)}")
        ws.update(
            f"A{match['_row']}:F{match['_row']}",
            [[item_name, size, new_qty, location or match["位置"], note or match["備註"], now_str()]],
        )
    else:
        if qty_delta < 0:
            raise ValueError("找不到此品項，無法直接出庫。")
        ws.append_row([item_name, size, qty_delta, location, note, now_str()])


def append_log(action: str, item_name: str, size: str, qty: int, location: str, note: str, actor: str):
    ws = log_ws()
    ws.append_row([now_str(), action, item_name, size, qty, location, note, actor])


@app.route("/", methods=["GET"])
def home():
    return "LINE Google Sheet inventory bot is running."


@app.route("/callback", methods=["POST"])
def callback():
    signature = request.headers.get("X-Line-Signature", "")
    body = request.get_data(as_text=True)
    app.logger.info("Request body: %s", body)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return "OK"


@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    text = event.message.text.strip()
    key = session_key(event)
    sess = get_session(key)

    if text in ["取消", "結束", "exit"]:
        clear_session(key)
        reply(event, "已取消目前流程。")
        return

    if sess:
        handle_session_reply(event, sess, text)
        return

    if text.startswith("查詢"):
        keyword = text.replace("查詢", "", 1).strip()
        if not keyword:
            reply(event, "請輸入：查詢 509")
            return
        rows = search_inventory(keyword)
        reply(event, format_search_result(rows, keyword))
        return

    if text.startswith("入庫") or text.startswith("出庫"):
        action = "入庫" if text.startswith("入庫") else "出庫"
        keyword = text.replace(action, "", 1).strip()
        if not keyword:
            reply(event, f"請輸入：{action} 509")
            return
        matches = distinct_items(search_inventory(keyword))
        if not matches:
            set_session(key, {
                "step": "new_item_name",
                "action": action,
                "keyword": keyword,
                "item_name": keyword,
                "size": "",
                "location": "",
                "note": "",
                "qty": None,
            })
            reply(event, f"找不到【{keyword}】現有資料。\n請直接輸入品名，或直接送出同樣品名建立新資料。")
            return
        if len(matches) == 1:
            row = matches[0]
            set_session(key, {
                "step": "size",
                "action": action,
                "keyword": keyword,
                "item_name": row["品名"],
                "size": row["尺寸"],
                "location": row["位置"],
                "note": row["備註"],
                "qty": None,
            })
            reply(event, f"品名：{row['品名']}\n目前尺寸：{row['尺寸'] or '-'}\n請輸入尺寸（直接送出原尺寸也可以）")
            return

        set_session(key, {
            "step": "choose_item",
            "action": action,
            "keyword": keyword,
            "choices": matches,
        })
        lines = [f"找到 {len(matches)} 個符合【{keyword}】的品項，請回覆編號："]
        for i, row in enumerate(matches[:15], start=1):
            lines.append(f"{i}. {row['品名']} / {row['尺寸']} / 庫存:{row['庫存']}")
        lines.append("例如輸入：1")
        reply(event, "\n".join(lines))
        return

    help_text = (
        "可用指令：\n"
        "1. 查詢 509\n"
        "2. 入庫 509\n"
        "3. 出庫 509\n"
        "4. 輸入 取消 可結束流程"
    )
    reply(event, help_text)



def handle_session_reply(event, sess: dict, text: str):
    key = session_key(event)
    step = sess.get("step")

    if step == "choose_item":
        try:
            idx = int(text) - 1
            row = sess["choices"][idx]
        except Exception:
            reply(event, "請輸入正確編號，例如 1")
            return
        sess.update({
            "step": "size",
            "item_name": row["品名"],
            "size": row["尺寸"],
            "location": row["位置"],
            "note": row["備註"],
            "qty": None,
        })
        set_session(key, sess)
        reply(event, f"已選擇：{row['品名']}\n目前尺寸：{row['尺寸'] or '-'}\n請輸入尺寸")
        return

    if step == "new_item_name":
        sess["item_name"] = text.strip()
        sess["step"] = "size"
        set_session(key, sess)
        reply(event, "請輸入尺寸")
        return

    if step == "size":
        sess["size"] = text.strip()
        sess["step"] = "location"
        set_session(key, sess)
        reply(event, "請輸入位置（不知道可輸入 - ）")
        return

    if step == "location":
        sess["location"] = text.strip()
        sess["step"] = "note"
        set_session(key, sess)
        reply(event, "有無任何需要備註？沒有請輸入 -")
        return

    if step == "note":
        sess["note"] = text.strip()
        sess["step"] = "qty"
        set_session(key, sess)
        reply(event, f"請輸入{sess['action']}數量")
        return

    if step == "qty":
        try:
            qty = int(text)
            if qty <= 0:
                raise ValueError
        except Exception:
            reply(event, "數量請輸入正整數，例如 5")
            return
        sess["qty"] = qty
        sess["step"] = "confirm"
        set_session(key, sess)
        summary = (
            f"請確認是否{sess['action']}：\n"
            f"品名：{sess['item_name']}\n"
            f"尺寸：{sess['size']}\n"
            f"位置：{sess['location']}\n"
            f"備註：{sess['note']}\n"
            f"數量：{sess['qty']}\n\n"
            f"確認請輸入：確認\n取消請輸入：取消"
        )
        reply(event, summary)
        return

    if step == "confirm":
        if text != "確認":
            reply(event, "請輸入【確認】完成，或輸入【取消】結束。")
            return
        action = sess["action"]
        qty_delta = sess["qty"] if action == "入庫" else -sess["qty"]
        try:
            create_or_update_inventory(
                sess["item_name"],
                sess["size"],
                qty_delta,
                sess["location"],
                sess["note"],
            )
            append_log(
                action,
                sess["item_name"],
                sess["size"],
                sess["qty"],
                sess["location"],
                sess["note"],
                actor_name(event),
            )
        except Exception as e:
            reply(event, f"處理失敗：{e}")
            clear_session(key)
            return

        rows = search_inventory(sess["item_name"])
        matched = None
        for row in rows:
            if normalize(row["品名"]) == normalize(sess["item_name"]) and normalize(row["尺寸"]) == normalize(sess["size"]):
                matched = row
                break
        clear_session(key)
        remain = matched["庫存"] if matched else "-"
        reply(event, f"已完成{action}\n品名：{sess['item_name']}\n尺寸：{sess['size']}\n本次數量：{sess['qty']}\n目前庫存：{remain}")
        return


def reply(event, text: str):
    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=text[:4900]))


if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

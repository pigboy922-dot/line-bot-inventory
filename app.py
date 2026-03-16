import os
import json
from datetime import datetime
from typing import Dict, List, Optional, Tuple

from flask import Flask, request, abort

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

import gspread
from google.oauth2.service_account import Credentials
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import (
    MessageEvent,
    TextMessage,
    TextSendMessage,
    QuickReply,
    QuickReplyButton,
    MessageAction,
)

app = Flask(__name__)

LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN", "")
LINE_CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET", "")
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "")
GOOGLE_JSON_RAW = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON") or os.getenv("GOOGLE_CREDENTIALS_JSON") or ""
INVENTORY_SHEET_NAME = os.getenv("INVENTORY_SHEET_NAME", "Sheet1")
LOG_SHEET_NAME = os.getenv("LOG_SHEET_NAME", "紀錄")
LOW_STOCK_THRESHOLD = int(os.getenv("LOW_STOCK_THRESHOLD", "10"))
APP_TIMEZONE = os.getenv("APP_TIMEZONE", "Asia/Taipei")
PORT = int(os.getenv("PORT", "10000"))

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

SESSIONS: Dict[str, dict] = {}


def now_str() -> str:
    try:
        from zoneinfo import ZoneInfo
        return datetime.now(ZoneInfo(APP_TIMEZONE)).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def session_key(event: MessageEvent) -> str:
    src = event.source
    for attr in ("user_id", "group_id", "room_id"):
        value = getattr(src, attr, None)
        if value:
            return value
    return "default"


def reset_session(key: str) -> None:
    SESSIONS[key] = {"state": None}


def quick_reply_texts(items: List[str]) -> QuickReply:
    return QuickReply(items=[QuickReplyButton(action=MessageAction(label=t[:20], text=t)) for t in items[:13]])


def reply(token: str, text: str, buttons: Optional[List[str]] = None) -> None:
    kwargs = {}
    if buttons:
        kwargs["quick_reply"] = quick_reply_texts(buttons)
    line_bot_api.reply_message(token, TextSendMessage(text=text, **kwargs))


def normalize_header(value: str) -> str:
    return (value or "").strip().lower().replace(" ", "")


HEADER_ALIASES = {
    "name": ["品名", "名稱", "型號", "品號", "料號"],
    "size": ["尺寸", "尺寸/mm", "尺寸mm", "規格"],
    "qty": ["數量", "庫存", "qty"],
    "location": ["位置", "儲位", "庫位"],
    "note": ["備註", "說明"],
}


def parse_google_json() -> dict:
    if not GOOGLE_JSON_RAW:
        raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON 未設定")
    try:
        return json.loads(GOOGLE_JSON_RAW)
    except json.JSONDecodeError:
        # fallback for accidentally escaped string
        try:
            return json.loads(GOOGLE_JSON_RAW.strip().strip('"'))
        except Exception as e:
            raise RuntimeError(f"GOOGLE JSON 格式錯誤: {e}")


def open_sheets() -> Tuple[gspread.Worksheet, gspread.Worksheet]:
    info = parse_google_json()
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(GOOGLE_SHEET_ID)
    ws = sh.worksheet(INVENTORY_SHEET_NAME)
    try:
        log_ws = sh.worksheet(LOG_SHEET_NAME)
    except gspread.WorksheetNotFound:
        log_ws = sh.add_worksheet(title=LOG_SHEET_NAME, rows=1000, cols=10)
        log_ws.append_row(["時間", "動作", "品名", "尺寸", "數量", "位置", "備註", "來源"])
    return ws, log_ws


def get_header_map(headers: List[str]) -> Dict[str, int]:
    normalized = {normalize_header(h): i for i, h in enumerate(headers)}
    result: Dict[str, int] = {}
    for key, aliases in HEADER_ALIASES.items():
        for alias in aliases:
            idx = normalized.get(normalize_header(alias))
            if idx is not None:
                result[key] = idx
                break
    return result


def inventory_records() -> Tuple[gspread.Worksheet, List[dict], Dict[str, int]]:
    ws, _ = open_sheets()
    values = ws.get_all_values()
    if not values:
        raise RuntimeError("試算表沒有資料")
    headers = values[0]
    header_map = get_header_map(headers)
    if "name" not in header_map:
        raise RuntimeError("找不到 品名 欄位")

    records: List[dict] = []
    for row_no, row in enumerate(values[1:], start=2):
        # pad row
        if len(row) < len(headers):
            row = row + [""] * (len(headers) - len(row))
        rec = {
            "row": row_no,
            "name": row[header_map.get("name", -1)] if header_map.get("name") is not None else "",
            "size": row[header_map.get("size", -1)] if header_map.get("size") is not None else "",
            "qty": row[header_map.get("qty", -1)] if header_map.get("qty") is not None else "0",
            "location": row[header_map.get("location", -1)] if header_map.get("location") is not None else "",
            "note": row[header_map.get("note", -1)] if header_map.get("note") is not None else "",
        }
        if any(str(v).strip() for k, v in rec.items() if k != "row"):
            records.append(rec)
    return ws, records, header_map


def search_records(keyword: str) -> List[dict]:
    _, records, _ = inventory_records()
    kw = keyword.strip().lower()
    results = []
    for rec in records:
        hay = f"{rec['name']} {rec['size']} {rec['location']} {rec['note']}".lower()
        if kw in hay:
            results.append(rec)
    return results




def format_all_inventory(records: List[dict], limit: int = 40) -> str:
    lines = [f"全部塊材庫存（顯示前 {min(len(records), limit)} 筆 / 共 {len(records)} 筆）："]
    for i, rec in enumerate(records[:limit], start=1):
        lines.append(f"{i}. {rec['name']} / {rec['size']} / 數量 {rec['qty']} / 位置 {rec['location']}")
    if len(records) > limit:
        lines.append("資料較多，請改用關鍵字查詢更精準。")
    return "\n".join(lines)
def format_result_list(results: List[dict]) -> str:
    lines = ["找到以下結果，請輸入編號："]
    for i, rec in enumerate(results[:10], start=1):
        lines.append(f"{i}. {rec['name']} / {rec['size']} / 庫存 {rec['qty']} / 位置 {rec['location']}")
    return "\n".join(lines)


def show_main_menu(token: str) -> None:
    reply(
        token,
        "塊材查詢按鈕版\n請選擇功能：",
        ["查詢庫存", "全部庫存", "入庫", "出庫", "低庫存", "取消"],
    )


def append_log(action: str, rec: dict, qty: int, note: str, source: str = "LINE") -> None:
    _, log_ws = open_sheets()
    log_ws.append_row([
        now_str(),
        action,
        rec.get("name", ""),
        rec.get("size", ""),
        str(qty),
        rec.get("location", ""),
        note,
        source,
    ])


def update_inventory_qty(rec: dict, new_qty: int, note: Optional[str] = None) -> None:
    ws, _, header_map = inventory_records()
    row = rec["row"]
    qty_col = header_map.get("qty")
    if qty_col is None:
        raise RuntimeError("找不到 數量 欄位")
    ws.update_cell(row, qty_col + 1, str(new_qty))
    if note is not None and header_map.get("note") is not None:
        ws.update_cell(row, header_map["note"] + 1, note)


def add_inventory_row(data: dict) -> None:
    ws, _, header_map = inventory_records()
    headers = ws.row_values(1)
    row = [""] * len(headers)
    mapping = {
        "name": data.get("name", ""),
        "size": data.get("size", ""),
        "qty": str(data.get("qty", "0")),
        "location": data.get("location", ""),
        "note": data.get("note", ""),
    }
    for key, idx in header_map.items():
        row[idx] = mapping.get(key, "")
    ws.append_row(row)
    append_log("新增", data, int(data.get("qty", 0)), data.get("note", ""))


@app.route("/")
def index():
    return "LINE Bot running", 200


@app.route("/callback", methods=["POST"])
def callback():
    signature = request.headers.get("X-Line-Signature", "")
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return "OK", 200


@handler.add(MessageEvent, message=TextMessage)
def handle_message(event: MessageEvent):
    text = (event.message.text or "").strip()
    key = session_key(event)
    session = SESSIONS.setdefault(key, {"state": None})

    try:
        if text in ["取消", "4"]:
            reset_session(key)
            reply(event.reply_token, "已取消流程。", ["塊材查詢"])
            return

        if text in ["塊材查詢", "主選單"]:
            reset_session(key)
            show_main_menu(event.reply_token)
            return

        if text in ["全部庫存", "查詢所有塊材庫存", "所有庫存"]:
            _, records, _ = inventory_records()
            if not records:
                reply(event.reply_token, "目前沒有任何塊材庫存資料。", ["塊材查詢"])
                return
            reply(event.reply_token, format_all_inventory(records), ["塊材查詢", "查詢庫存", "低庫存"])
            return

        if text == "低庫存":
            _, records, _ = inventory_records()
            lows = []
            for rec in records:
                try:
                    qty = int(float(rec.get("qty") or 0))
                except Exception:
                    qty = 0
                if qty <= LOW_STOCK_THRESHOLD:
                    lows.append(rec)
            if not lows:
                reply(event.reply_token, f"目前沒有低於 {LOW_STOCK_THRESHOLD} 的庫存。", ["塊材查詢"])
                return
            lines = [f"低庫存（≦{LOW_STOCK_THRESHOLD}）："]
            for rec in lows[:20]:
                lines.append(f"- {rec['name']} / {rec['size']} / {rec['qty']} / {rec['location']}")
            reply(event.reply_token, "\n".join(lines), ["塊材查詢"])
            return

        if text in ["查詢庫存", "1"] and session.get("state") is None:
            session.update({"state": "await_keyword", "mode": "query"})
            reply(event.reply_token, "請輸入關鍵字", ["取消"])
            return

        if text in ["入庫", "2"] and session.get("state") is None:
            session.update({"state": "await_keyword", "mode": "in"})
            reply(event.reply_token, "請輸入關鍵字", ["取消"])
            return

        if text in ["出庫", "3"] and session.get("state") is None:
            session.update({"state": "await_keyword", "mode": "out"})
            reply(event.reply_token, "請輸入關鍵字", ["取消"])
            return

        if session.get("state") == "await_keyword":
            mode = session.get("mode")
            results = search_records(text)
            if not results:
                if mode == "in":
                    session.update({"state": "ask_manual_add", "keyword": text})
                    reply(
                        event.reply_token,
                        f"找不到【{text}】現有品項。\n可選擇手動新增。",
                        ["手動新增", "重新搜尋", "取消"],
                    )
                    return
                reply(event.reply_token, f"找不到【{text}】相關資料。", ["塊材查詢", "重新搜尋", "取消"])
                session.update({"state": None})
                return
            session.update({"state": "await_select", "results": results[:10]})
            buttons = [str(i) for i in range(1, min(len(results), 10) + 1)] + ["取消"]
            reply(event.reply_token, format_result_list(results), buttons)
            return

        if text == "重新搜尋":
            session.update({"state": "await_keyword"})
            reply(event.reply_token, "請重新輸入關鍵字", ["取消"])
            return

        if session.get("state") == "ask_manual_add":
            if text == "手動新增":
                session.update({"state": "add_name", "new_item": {"name": session.get("keyword", "")}})
                reply(event.reply_token, "請輸入完整品名", ["取消"])
                return
            if text == "重新搜尋":
                session.update({"state": "await_keyword"})
                reply(event.reply_token, "請重新輸入關鍵字", ["取消"])
                return

        if session.get("state") == "add_name":
            session["new_item"]["name"] = text
            session["state"] = "add_size"
            reply(event.reply_token, "請輸入尺寸", ["取消"])
            return

        if session.get("state") == "add_size":
            session["new_item"]["size"] = text
            session["state"] = "add_location"
            reply(event.reply_token, "請輸入位置", ["取消"])
            return

        if session.get("state") == "add_location":
            session["new_item"]["location"] = text
            session["state"] = "add_note"
            reply(event.reply_token, "請輸入備註，沒有請輸入 無", ["取消"])
            return

        if session.get("state") == "add_note":
            session["new_item"]["note"] = "" if text == "無" else text
            session["state"] = "add_qty"
            reply(event.reply_token, "請輸入入庫數量", ["取消"])
            return

        if session.get("state") == "add_qty":
            qty = int(text)
            session["new_item"]["qty"] = qty
            session["state"] = "add_confirm"
            ni = session["new_item"]
            msg = (
                "確認新增品項並入庫？\n"
                f"品名：{ni['name']}\n"
                f"尺寸：{ni['size']}\n"
                f"位置：{ni['location']}\n"
                f"備註：{ni.get('note','')}\n"
                f"數量：{ni['qty']}"
            )
            reply(event.reply_token, msg, ["確認新增", "取消"])
            return

        if session.get("state") == "add_confirm" and text == "確認新增":
            add_inventory_row(session["new_item"])
            reset_session(key)
            reply(event.reply_token, "新增成功。", ["塊材查詢"])
            return

        if session.get("state") == "await_select":
            if not text.isdigit():
                reply(event.reply_token, "請輸入編號。", [str(i) for i in range(1, len(session.get("results", [])) + 1)] + ["取消"])
                return
            idx = int(text) - 1
            results = session.get("results", [])
            if idx < 0 or idx >= len(results):
                reply(event.reply_token, "編號超出範圍。", [str(i) for i in range(1, len(results) + 1)] + ["取消"])
                return
            rec = results[idx]
            mode = session.get("mode")
            session["selected"] = rec
            if mode == "query":
                reset_session(key)
                reply(
                    event.reply_token,
                    f"品名：{rec['name']}\n尺寸：{rec['size']}\n庫存：{rec['qty']}\n位置：{rec['location']}\n備註：{rec['note']}",
                    ["塊材查詢", "入庫", "出庫"],
                )
                return
            session["state"] = "await_qty"
            action = "入庫" if mode == "in" else "出庫"
            reply(event.reply_token, f"已選擇：{rec['name']} / {rec['size']}\n請輸入{action}數量", ["取消"])
            return

        if session.get("state") == "await_qty":
            qty = int(text)
            if qty <= 0:
                raise ValueError("數量需大於 0")
            session["qty"] = qty
            session["state"] = "await_note"
            reply(event.reply_token, "請輸入備註，沒有請輸入 無", ["取消"])
            return

        if session.get("state") == "await_note":
            session["note"] = "" if text == "無" else text
            session["state"] = "await_confirm"
            rec = session["selected"]
            action = "入庫" if session.get("mode") == "in" else "出庫"
            reply(
                event.reply_token,
                f"確認{action}？\n品名：{rec['name']}\n尺寸：{rec['size']}\n數量：{session['qty']}\n位置：{rec['location']}\n備註：{session['note']}",
                ["確認", "取消"],
            )
            return

        if session.get("state") == "await_confirm" and text == "確認":
            rec = session["selected"]
            qty = int(session["qty"])
            note = session.get("note", "")
            current = int(float(rec.get("qty") or 0))
            if session.get("mode") == "in":
                new_qty = current + qty
                update_inventory_qty(rec, new_qty, note or rec.get("note", ""))
                append_log("入庫", rec, qty, note)
                reply(event.reply_token, f"入庫成功。\n{rec['name']} / {rec['size']}\n最新庫存：{new_qty}", ["塊材查詢"])
            else:
                if qty > current:
                    reply(event.reply_token, f"出庫失敗，庫存不足。現有庫存：{current}", ["塊材查詢", "取消"])
                    reset_session(key)
                    return
                new_qty = current - qty
                update_inventory_qty(rec, new_qty, note or rec.get("note", ""))
                append_log("出庫", rec, qty, note)
                reply(event.reply_token, f"出庫成功。\n{rec['name']} / {rec['size']}\n最新庫存：{new_qty}", ["塊材查詢"])
            reset_session(key)
            return

        show_main_menu(event.reply_token)

    except Exception as e:
        reset_session(key)
        reply(event.reply_token, f"系統處理失敗：{e}", ["塊材查詢"])


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=PORT)

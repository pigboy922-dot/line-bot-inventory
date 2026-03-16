import os
import json
import traceback
from datetime import datetime
from typing import Dict, Any, List, Optional

from flask import Flask, request, abort
from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage

load_dotenv()

app = Flask(__name__)

LINE_CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET", "")
LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN", "")
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "")
GOOGLE_SERVICE_ACCOUNT_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "")
APP_TIMEZONE = os.getenv("APP_TIMEZONE", "Asia/Taipei")
WORKSHEET_NAME = os.getenv("GOOGLE_WORKSHEET_NAME", "Sheet1")

if not LINE_CHANNEL_SECRET or not LINE_CHANNEL_ACCESS_TOKEN:
    print("[WARN] LINE_CHANNEL_SECRET / LINE_CHANNEL_ACCESS_TOKEN 未設定")

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# 簡易記憶體狀態。Render 重啟後會清空，但一般互動流程足夠。
USER_STATES: Dict[str, Dict[str, Any]] = {}

HEADER_ALIASES = {
    "name": ["品名", "名稱", "型號", "品號", "料號", "item", "name"],
    "size": ["尺寸", "尺寸/mm", "尺寸mm", "規格", "size"],
    "qty": ["數量", "庫存", "qty", "quantity"],
    "location": ["位置", "儲位", "location", "loc"],
    "remark": ["備註", "註記", "remark", "note", "notes"],
    "updated_at": ["更新時間", "最後更新", "updated_at", "time", "時間"],
}


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def parse_service_account_info() -> Dict[str, Any]:
    raw = GOOGLE_SERVICE_ACCOUNT_JSON.strip()
    if not raw:
        raise ValueError("GOOGLE_SERVICE_ACCOUNT_JSON 未設定")
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        # 允許貼上被跳脫的 json
        return json.loads(raw.encode("utf-8").decode("unicode_escape"))


def get_worksheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(parse_service_account_info(), scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(GOOGLE_SHEET_ID)
    try:
        return sh.worksheet(WORKSHEET_NAME)
    except gspread.WorksheetNotFound:
        return sh.sheet1


def normalize_text(s: Any) -> str:
    return str(s).strip() if s is not None else ""


def canonicalize_headers(headers: List[str]) -> Dict[str, int]:
    mapping: Dict[str, int] = {}
    norm_headers = [normalize_text(h) for h in headers]
    for canon, aliases in HEADER_ALIASES.items():
        for idx, h in enumerate(norm_headers):
            h_low = h.lower().replace(" ", "")
            for alias in aliases:
                if h_low == alias.lower().replace(" ", ""):
                    mapping[canon] = idx
                    break
            if canon in mapping:
                break
    return mapping


def ensure_headers(ws):
    headers = ws.row_values(1)
    if not headers:
        ws.update("A1:E1", [["品名", "尺寸", "數量", "位置", "備註"]])
        headers = ws.row_values(1)
    mapping = canonicalize_headers(headers)

    desired = ["品名", "尺寸", "數量", "位置", "備註", "更新時間"]
    changed = False
    current = headers[:]
    if len(current) < len(desired):
        current.extend([""] * (len(desired) - len(current)))
    canonical_current = canonicalize_headers(current)

    # 依據現有可辨識欄位填補固定位置，避免抓不到「尺寸/mm」之類欄位
    for i, key in enumerate(["name", "size", "qty", "location", "remark", "updated_at"]):
        label = desired[i]
        if i >= len(headers) or not normalize_text(headers[i]):
            current[i] = label
            changed = True
        elif key not in canonical_current:
            current[i] = label
            changed = True

    if changed:
        ws.update(f"A1:F1", [current[:6]])
        headers = ws.row_values(1)
        mapping = canonicalize_headers(headers)

    return headers, mapping


def all_records(ws) -> List[Dict[str, Any]]:
    headers, mapping = ensure_headers(ws)
    values = ws.get_all_values()
    rows = values[1:] if len(values) > 1 else []
    records = []
    for idx, row in enumerate(rows, start=2):
        row = row + [""] * (len(headers) - len(row))
        record = {
            "row_number": idx,
            "name": row[mapping.get("name", 0)] if mapping.get("name") is not None else "",
            "size": row[mapping.get("size", 1)] if mapping.get("size") is not None else "",
            "qty": row[mapping.get("qty", 2)] if mapping.get("qty") is not None else "",
            "location": row[mapping.get("location", 3)] if mapping.get("location") is not None else "",
            "remark": row[mapping.get("remark", 4)] if mapping.get("remark") is not None else "",
            "updated_at": row[mapping.get("updated_at", 5)] if mapping.get("updated_at") is not None else "",
        }
        if any(normalize_text(v) for k, v in record.items() if k != "row_number"):
            records.append(record)
    return records


def search_records(keyword: str) -> List[Dict[str, Any]]:
    ws = get_worksheet()
    records = all_records(ws)
    kw = normalize_text(keyword).lower().replace(" ", "")
    if not kw:
        return []
    results = []
    for r in records:
        hay = " ".join([normalize_text(r["name"]), normalize_text(r["size"]), normalize_text(r["location"])]).lower().replace(" ", "")
        if kw in hay:
            results.append(r)
    return results


def source_key(event) -> str:
    src = event.source
    t = getattr(src, "type", "user")
    if t == "group":
        return f"group:{src.group_id}"
    if t == "room":
        return f"room:{src.room_id}"
    return f"user:{src.user_id}"


def reset_state(key: str):
    USER_STATES.pop(key, None)


def set_state(key: str, **kwargs):
    USER_STATES[key] = kwargs


def get_state(key: str) -> Dict[str, Any]:
    return USER_STATES.get(key, {})


def menu_text() -> str:
    return "塊材查詢選單：\n1. 查詢\n2. 入庫\n3. 出庫\n4. 取消\n\n請輸入 1、2、3、4"


def format_search_results(results: List[Dict[str, Any]]) -> str:
    if not results:
        return "查無資料，請重新輸入關鍵字，或輸入 4 取消。"
    lines = ["找到以下結果，請輸入編號："]
    for i, r in enumerate(results[:20], start=1):
        lines.append(f"{i}. {r['name']} / 尺寸 {r['size'] or '-'} / 庫存 {r['qty'] or '0'} / 位置 {r['location'] or '-'}")
    if len(results) > 20:
        lines.append(f"\n共 {len(results)} 筆，先顯示前 20 筆。")
    return "\n".join(lines)


def to_int(value: str) -> Optional[int]:
    try:
        return int(str(value).strip())
    except Exception:
        return None


def update_row_qty(row_number: int, new_qty: int, remark: str = ""):
    ws = get_worksheet()
    headers, mapping = ensure_headers(ws)
    qty_col = mapping.get("qty", 2) + 1
    remark_col = mapping.get("remark", 4) + 1
    updated_at_col = mapping.get("updated_at", 5) + 1
    ws.update_cell(row_number, qty_col, str(new_qty))
    if remark:
        ws.update_cell(row_number, remark_col, remark)
    ws.update_cell(row_number, updated_at_col, now_str())


def append_row(name: str, size: str, qty: int, location: str, remark: str = ""):
    ws = get_worksheet()
    ensure_headers(ws)
    ws.append_row([name, size, str(qty), location, remark, now_str()])


def reply_safe(reply_token: str, text: str):
    line_bot_api.reply_message(reply_token, TextSendMessage(text=text))


@app.route("/", methods=["GET", "HEAD"])
def home():
    return "LINE Bot running", 200


@app.route("/callback", methods=["POST"])
def callback():
    signature = request.headers.get("X-Line-Signature", "")
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    except Exception:
        print(traceback.format_exc())
        return "OK", 200
    return "OK", 200


@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    user_text = normalize_text(event.message.text)
    key = source_key(event)
    state = get_state(key)

    try:
        if user_text in ["取消", "4"] and (state or user_text == "4"):
            reset_state(key)
            reply_safe(event.reply_token, "已取消流程。\n\n輸入「塊材查詢」可重新開始。")
            return

        if user_text == "塊材查詢":
            set_state(key, step="menu")
            reply_safe(event.reply_token, menu_text())
            return

        if not state:
            return

        step = state.get("step")

        if step == "menu":
            if user_text == "1":
                set_state(key, step="query_keyword", action="query")
                reply_safe(event.reply_token, "請輸入關鍵字")
                return
            if user_text == "2":
                set_state(key, step="stockin_keyword", action="stockin")
                reply_safe(event.reply_token, "請輸入關鍵字")
                return
            if user_text == "3":
                set_state(key, step="stockout_keyword", action="stockout")
                reply_safe(event.reply_token, "請輸入關鍵字")
                return
            reply_safe(event.reply_token, "請輸入 1、2、3、4")
            return

        if step in ["query_keyword", "stockin_keyword", "stockout_keyword"]:
            results = search_records(user_text)
            action = state.get("action")

            if action == "query":
                if not results:
                    reply_safe(event.reply_token, f"找不到【{user_text}】相關資料。\n請重新輸入關鍵字，或輸入 4 取消。")
                    return
                reset_state(key)
                reply_safe(event.reply_token, format_search_results(results))
                return

            if action == "stockin":
                if not results:
                    set_state(key, step="stockin_not_found", action="stockin", keyword=user_text)
                    reply_safe(event.reply_token, f"找不到【{user_text}】現有編號。\n1. 手動新增\n2. 重新輸入關鍵字\n\n請輸入 1 或 2")
                    return
                set_state(key, step="stockin_pick", action="stockin", keyword=user_text, results=results)
                reply_safe(event.reply_token, format_search_results(results))
                return

            if action == "stockout":
                if not results:
                    reply_safe(event.reply_token, f"找不到【{user_text}】相關資料。\n請重新輸入關鍵字，或輸入 4 取消。")
                    return
                set_state(key, step="stockout_pick", action="stockout", keyword=user_text, results=results)
                reply_safe(event.reply_token, format_search_results(results))
                return

        if step == "stockin_not_found":
            if user_text == "1":
                set_state(key, step="manual_name")
                reply_safe(event.reply_token, "請輸入完整品名")
                return
            if user_text == "2":
                set_state(key, step="stockin_keyword", action="stockin")
                reply_safe(event.reply_token, "請重新輸入關鍵字")
                return
            reply_safe(event.reply_token, "請輸入 1 或 2")
            return

        if step == "manual_name":
            state["new_name"] = user_text
            state["step"] = "manual_size"
            USER_STATES[key] = state
            reply_safe(event.reply_token, "請輸入尺寸")
            return

        if step == "manual_size":
            state["new_size"] = user_text
            state["step"] = "manual_location"
            USER_STATES[key] = state
            reply_safe(event.reply_token, "請輸入位置")
            return

        if step == "manual_location":
            state["new_location"] = user_text
            state["step"] = "manual_remark"
            USER_STATES[key] = state
            reply_safe(event.reply_token, "請輸入備註，沒有請輸入 無")
            return

        if step == "manual_remark":
            state["new_remark"] = "" if user_text == "無" else user_text
            state["step"] = "manual_qty"
            USER_STATES[key] = state
            reply_safe(event.reply_token, "請輸入入庫數量")
            return

        if step == "manual_qty":
            qty = to_int(user_text)
            if qty is None or qty < 0:
                reply_safe(event.reply_token, "數量格式錯誤，請輸入整數")
                return
            state["new_qty"] = qty
            state["step"] = "manual_confirm"
            USER_STATES[key] = state
            preview = (
                f"請確認新增入庫：\n"
                f"品名：{state['new_name']}\n"
                f"尺寸：{state['new_size']}\n"
                f"位置：{state['new_location']}\n"
                f"備註：{state.get('new_remark', '') or '無'}\n"
                f"數量：{qty}\n\n"
                f"確認請輸入：確認"
            )
            reply_safe(event.reply_token, preview)
            return

        if step == "manual_confirm":
            if user_text != "確認":
                reply_safe(event.reply_token, "請輸入「確認」完成新增，或輸入 4 取消。")
                return
            append_row(
                state["new_name"],
                state["new_size"],
                state["new_qty"],
                state["new_location"],
                state.get("new_remark", ""),
            )
            reset_state(key)
            reply_safe(event.reply_token, "新增入庫完成。")
            return

        if step == "stockin_pick":
            results = state.get("results", [])
            idx = to_int(user_text)
            if idx is None or idx < 1 or idx > len(results):
                reply_safe(event.reply_token, "請輸入正確編號")
                return
            picked = results[idx - 1]
            state["picked"] = picked
            state["step"] = "stockin_qty"
            USER_STATES[key] = state
            reply_safe(event.reply_token, f"已選擇：{picked['name']} / 尺寸 {picked['size']} / 現有庫存 {picked['qty']}\n請輸入入庫數量")
            return

        if step == "stockin_qty":
            qty = to_int(user_text)
            if qty is None or qty < 0:
                reply_safe(event.reply_token, "數量格式錯誤，請輸入整數")
                return
            picked = state["picked"]
            current = to_int(picked.get("qty", "0")) or 0
            state["qty_change"] = qty
            state["new_total"] = current + qty
            state["step"] = "stockin_remark"
            USER_STATES[key] = state
            reply_safe(event.reply_token, "請輸入備註，沒有請輸入 無")
            return

        if step == "stockin_remark":
            remark = "" if user_text == "無" else user_text
            state["remark"] = remark
            state["step"] = "stockin_confirm"
            USER_STATES[key] = state
            picked = state["picked"]
            preview = (
                f"請確認入庫：\n"
                f"品名：{picked['name']}\n"
                f"尺寸：{picked['size']}\n"
                f"位置：{picked['location']}\n"
                f"原庫存：{picked['qty']}\n"
                f"入庫：{state['qty_change']}\n"
                f"新庫存：{state['new_total']}\n"
                f"備註：{remark or '無'}\n\n"
                f"確認請輸入：確認"
            )
            reply_safe(event.reply_token, preview)
            return

        if step == "stockin_confirm":
            if user_text != "確認":
                reply_safe(event.reply_token, "請輸入「確認」完成入庫，或輸入 4 取消。")
                return
            picked = state["picked"]
            update_row_qty(picked["row_number"], state["new_total"], state.get("remark", ""))
            reset_state(key)
            reply_safe(event.reply_token, "入庫完成。")
            return

        if step == "stockout_pick":
            results = state.get("results", [])
            idx = to_int(user_text)
            if idx is None or idx < 1 or idx > len(results):
                reply_safe(event.reply_token, "請輸入正確編號")
                return
            picked = results[idx - 1]
            state["picked"] = picked
            state["step"] = "stockout_qty"
            USER_STATES[key] = state
            reply_safe(event.reply_token, f"已選擇：{picked['name']} / 尺寸 {picked['size']} / 現有庫存 {picked['qty']}\n請輸入出庫數量")
            return

        if step == "stockout_qty":
            qty = to_int(user_text)
            picked = state["picked"]
            current = to_int(picked.get("qty", "0")) or 0
            if qty is None or qty < 0:
                reply_safe(event.reply_token, "數量格式錯誤，請輸入整數")
                return
            if qty > current:
                reply_safe(event.reply_token, f"出庫數量不可大於現有庫存 {current}")
                return
            state["qty_change"] = qty
            state["new_total"] = current - qty
            state["step"] = "stockout_remark"
            USER_STATES[key] = state
            reply_safe(event.reply_token, "請輸入備註，沒有請輸入 無")
            return

        if step == "stockout_remark":
            remark = "" if user_text == "無" else user_text
            state["remark"] = remark
            state["step"] = "stockout_confirm"
            USER_STATES[key] = state
            picked = state["picked"]
            preview = (
                f"請確認出庫：\n"
                f"品名：{picked['name']}\n"
                f"尺寸：{picked['size']}\n"
                f"位置：{picked['location']}\n"
                f"原庫存：{picked['qty']}\n"
                f"出庫：{state['qty_change']}\n"
                f"新庫存：{state['new_total']}\n"
                f"備註：{remark or '無'}\n\n"
                f"確認請輸入：確認"
            )
            reply_safe(event.reply_token, preview)
            return

        if step == "stockout_confirm":
            if user_text != "確認":
                reply_safe(event.reply_token, "請輸入「確認」完成出庫，或輸入 4 取消。")
                return
            picked = state["picked"]
            update_row_qty(picked["row_number"], state["new_total"], state.get("remark", ""))
            reset_state(key)
            reply_safe(event.reply_token, "出庫完成。")
            return

        reply_safe(event.reply_token, "流程未識別，請輸入「塊材查詢」重新開始。")
    except Exception as e:
        print("[ERROR] handle_message failed:")
        print(traceback.format_exc())
        try:
            reply_safe(event.reply_token, f"系統處理失敗：{str(e)}")
        except Exception:
            pass


if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

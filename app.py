import os
import json
from datetime import datetime
from flask import Flask, request, jsonify
from flask_cors import CORS
import gspread
from google.oauth2.service_account import Credentials
from google.auth.transport import requests as google_requests
from google.oauth2 import id_token as google_id_token

app = Flask(__name__)
CORS(app)

GOOGLE_CREDENTIALS_JSON = os.getenv("GOOGLE_CREDENTIALS_JSON", "")
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "")
LIFF_ID = os.getenv("LIFF_ID", "")
PAGE_SIZE_DEFAULT = 10

if not GOOGLE_CREDENTIALS_JSON:
    raise ValueError("缺少環境變數 GOOGLE_CREDENTIALS_JSON")
if not GOOGLE_SHEET_ID:
    raise ValueError("缺少環境變數 GOOGLE_SHEET_ID")

scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
creds_info = json.loads(GOOGLE_CREDENTIALS_JSON)
credentials = Credentials.from_service_account_info(creds_info, scopes=scope)
gc = gspread.authorize(credentials)
spreadsheet = gc.open_by_key(GOOGLE_SHEET_ID)
sheet = spreadsheet.sheet1


def ensure_log_worksheet():
    headers = [
        "時間", "聊天室類型", "群組名稱", "群組ID", "room_id", "user_key",
        "動作", "品名", "尺寸", "原數量", "異動數量", "新數量", "位置", "備註"
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


def log_inventory_action(action, item=None, old_qty="", change_qty="", new_qty="", note="", actor=None):
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
        now,
        "liff",
        "",
        "",
        "",
        actor.get("line_user_id", ""),
        action,
        name,
        size,
        old_qty,
        change_qty,
        new_qty,
        loc,
        note or actor.get("line_name", "")
    ])


@app.get("/")
def home():
    return jsonify({"ok": True, "message": "LIFF Inventory API Running"})


@app.get("/health")
def health():
    try:
        missing = required_columns_ok()
        return jsonify({
            "ok": True,
            "message": "OK",
            "sheet_id": GOOGLE_SHEET_ID,
            "missing_columns": missing
        })
    except Exception as e:
        return jsonify({"ok": False, "message": str(e)}), 500


@app.get("/api/search")
def api_search():
    missing = required_columns_ok()
    if missing:
        return jsonify({"ok": False, "message": f"Sheet 缺少欄位：{', '.join(missing)}"}), 400

    keyword = request.args.get("q", "").strip()
    if not keyword:
        return jsonify({"ok": True, "items": []})
    items = find_matching_rows(keyword)
    return jsonify({"ok": True, "items": items[:50]})


@app.get("/api/stock")
def api_stock():
    missing = required_columns_ok()
    if missing:
        return jsonify({"ok": False, "message": f"Sheet 缺少欄位：{', '.join(missing)}"}), 400

    try:
        page = max(1, int(request.args.get("page", 1)))
        page_size = max(1, int(request.args.get("page_size", PAGE_SIZE_DEFAULT)))
    except Exception:
        return jsonify({"ok": False, "message": "page 或 page_size 格式錯誤"}), 400

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


@app.post("/api/in")
def api_in():
    missing = required_columns_ok()
    if missing:
        return jsonify({"ok": False, "message": f"Sheet 缺少欄位：{', '.join(missing)}"}), 400

    data = request.get_json(silent=True) or {}
    actor = get_actor_info()

    try:
        row_num = int(data.get("row_number", 0))
        add_qty = int(data.get("qty", 0))
    except Exception:
        return jsonify({"ok": False, "message": "row_number 或 qty 格式錯誤"}), 400

    if row_num < 2:
        return jsonify({"ok": False, "message": "row_number 錯誤"}), 400
    if add_qty <= 0:
        return jsonify({"ok": False, "message": "入庫數量必須大於 0"}), 400

    item = get_item_by_row(row_num)
    if not item:
        return jsonify({"ok": False, "message": "找不到該筆資料"}), 404

    qty_col = get_col_index("數量")
    old_qty = to_int(sheet.cell(row_num, qty_col).value)
    new_qty = old_qty + add_qty
    sheet.update_cell(row_num, qty_col, new_qty)

    item["數量"] = new_qty
    log_inventory_action(
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
        "item": item
    })


@app.post("/api/out")
def api_out():
    missing = required_columns_ok()
    if missing:
        return jsonify({"ok": False, "message": f"Sheet 缺少欄位：{', '.join(missing)}"}), 400

    data = request.get_json(silent=True) or {}
    actor = get_actor_info()

    try:
        row_num = int(data.get("row_number", 0))
        out_qty = int(data.get("qty", 0))
    except Exception:
        return jsonify({"ok": False, "message": "row_number 或 qty 格式錯誤"}), 400

    if row_num < 2:
        return jsonify({"ok": False, "message": "row_number 錯誤"}), 400
    if out_qty <= 0:
        return jsonify({"ok": False, "message": "出庫數量必須大於 0"}), 400

    item = get_item_by_row(row_num)
    if not item:
        return jsonify({"ok": False, "message": "找不到該筆資料"}), 404

    qty_col = get_col_index("數量")
    old_qty = to_int(sheet.cell(row_num, qty_col).value)
    if out_qty > old_qty:
        return jsonify({"ok": False, "message": f"目前庫存只有 {old_qty}，不能出庫 {out_qty}"}), 400

    new_qty = old_qty - out_qty
    sheet.update_cell(row_num, qty_col, new_qty)

    item["數量"] = new_qty
    log_inventory_action(
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
        "item": item
    })


@app.post("/api/manual-in")
def api_manual_in():
    missing = required_columns_ok()
    if missing:
        return jsonify({"ok": False, "message": f"Sheet 缺少欄位：{', '.join(missing)}"}), 400

    data = request.get_json(silent=True) or {}
    actor = get_actor_info()

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
    sheet.append_row(new_row)

    item = {"品名": name, "尺寸": size, "數量": qty, "位置": location}
    log_inventory_action(
        "手動入庫",
        item=item,
        old_qty=0,
        change_qty=qty,
        new_qty=qty,
        note=f"LIFF手動入庫 / {actor.get('line_name', '')}",
        actor=actor
    )

    return jsonify({"ok": True, "message": "手動入庫完成", "item": item})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

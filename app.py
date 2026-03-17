import os
import json
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.models import *
from linebot.exceptions import InvalidSignatureError
import gspread
from google.oauth2.service_account import Credentials

app = Flask(__name__)

LINE_CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET")
LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN")
GOOGLE_CREDENTIALS_JSON = os.getenv("GOOGLE_CREDENTIALS_JSON")
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

creds_info = json.loads(GOOGLE_CREDENTIALS_JSON)
credentials = Credentials.from_service_account_info(creds_info, scopes=scope)
gc = gspread.authorize(credentials)
sheet = gc.open_by_key(GOOGLE_SHEET_ID).sheet1

user_state = {}
user_data = {}

PAGE_SIZE = 10


@app.route("/", methods=["GET"])
def home():
    return "LINE BOT Running", 200


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


def get_user_key(event):
    source = event.source
    if hasattr(source, "user_id") and source.user_id:
        return source.user_id
    if hasattr(source, "group_id") and source.group_id:
        return f"group_{source.group_id}"
    return "unknown"


@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    user_text = str(event.message.text).strip()
    user_id = get_user_key(event)

    print("訊息:", user_text)
    print("狀態:", user_state.get(user_id))

    # ===== 取消 =====
    if user_text == "取消":
        clear_user_session(user_id)
        reply_text(event.reply_token, "已取消")
        return

    # ===== 主選單 =====
    if user_text == "塊材查詢":
        clear_user_session(user_id)
        send_menu(event.reply_token)
        return

    # ===== 功能選單 =====
    if user_text == "查詢庫存":
        user_state[user_id] = "search"
        reply_text(event.reply_token, "請輸入關鍵字")
        return

    if user_text == "全部庫存":
        show_all_stock(event.reply_token, 1)
        return

    if user_text == "入庫":
        user_state[user_id] = "in_keyword"
        reply_text(event.reply_token, "請輸入關鍵字")
        return

    if user_text == "出庫":
        user_state[user_id] = "out_keyword"
        reply_text(event.reply_token, "請輸入關鍵字")
        return

    if user_text == "手動入庫":
        user_state[user_id] = "manual_name"
        user_data[user_id] = {}
        reply_text(event.reply_token, "請輸入品名")
        return

    # ===== 查詢 =====
    if user_state.get(user_id) == "search":
        search_stock(event.reply_token, user_text)
        clear_user_session(user_id)
        return

    # ===== 入庫 =====
    if user_state.get(user_id) == "in_keyword":
        search_stock_for_in(event.reply_token, user_id, user_text)
        return

    if user_state.get(user_id) == "in_qty":
        process_in_qty(event.reply_token, user_id, user_text)
        return

    # ===== 出庫 =====
    if user_state.get(user_id) == "out_keyword":
        search_stock_for_out(event.reply_token, user_id, user_text)
        return

    if user_state.get(user_id) == "out_qty":
        process_out_qty(event.reply_token, user_id, user_text)
        return

    # ===== 手動入庫 =====
    if user_state.get(user_id) == "manual_name":
        user_data[user_id]["品名"] = user_text
        user_state[user_id] = "manual_size"
        reply_text(event.reply_token, "請輸入尺寸")
        return

    if user_state.get(user_id) == "manual_size":
        user_data[user_id]["尺寸"] = user_text
        user_state[user_id] = "manual_qty"
        reply_text(event.reply_token, "請輸入數量")
        return

    if user_state.get(user_id) == "manual_qty":
        if not user_text.isdigit():
            reply_text(event.reply_token, "請輸入數字")
            return
        user_data[user_id]["數量"] = int(user_text)
        user_state[user_id] = "manual_loc"
        reply_text(event.reply_token, "請輸入位置")
        return

    if user_state.get(user_id) == "manual_loc":
        user_data[user_id]["位置"] = user_text
        save_manual_stock(event.reply_token, user_id)
        return

    # ❌ 其它全部不回應
    return


# ===== 功能區 =====

def send_menu(token):
    line_bot_api.reply_message(token,
        TemplateSendMessage(
            alt_text="選單",
            template=ButtonsTemplate(
                title="塊材管理",
                text="請選擇功能",
                actions=[
                    MessageTemplateAction(label="查詢庫存", text="查詢庫存"),
                    MessageTemplateAction(label="入庫", text="入庫"),
                    MessageTemplateAction(label="出庫", text="出庫"),
                    MessageTemplateAction(label="手動入庫", text="手動入庫")
                ]
            )
        )
    )


def show_all_stock(token, page):
    data = sheet.get_all_records()
    start = (page - 1) * PAGE_SIZE
    end = start + PAGE_SIZE
    rows = data[start:end]

    msg = f"庫存（第{page}頁）\n\n"
    for i, r in enumerate(rows, start=start + 1):
        msg += f"{i}. {r['品名']} / {r['尺寸']} / {r['數量']}\n"

    reply_text(token, msg)


def search_stock(token, keyword):
    data = sheet.get_all_records()
    msg = "搜尋結果：\n"

    for i, r in enumerate(data, start=2):
        if keyword in str(r['品名']) or keyword in str(r['尺寸']):
            msg += f"{i}. {r['品名']} / {r['尺寸']} / {r['數量']}\n"

    reply_text(token, msg)


def search_stock_for_in(token, user_id, keyword):
    data = sheet.get_all_records()

    for i, r in enumerate(data, start=2):
        if keyword in str(r['品名']) or keyword in str(r['尺寸']):
            user_data[user_id] = {"row": i}
            user_state[user_id] = "in_qty"
            reply_text(token, f"{r['品名']} 目前:{r['數量']}，輸入入庫數量")
            return

    reply_text(token, "找不到資料，可使用『手動入庫』")


def process_in_qty(token, user_id, qty):
    if not qty.isdigit():
        reply_text(token, "請輸入數字")
        return

    row = user_data[user_id]["row"]
    col = get_col("數量")

    old = int(sheet.cell(row, col).value)
    new = old + int(qty)

    sheet.update_cell(row, col, new)

    reply_text(token, f"入庫完成 {old} → {new}")
    clear_user_session(user_id)


def search_stock_for_out(token, user_id, keyword):
    data = sheet.get_all_records()

    for i, r in enumerate(data, start=2):
        if keyword in str(r['品名']) or keyword in str(r['尺寸']):
            user_data[user_id] = {"row": i}
            user_state[user_id] = "out_qty"
            reply_text(token, f"{r['品名']} 目前:{r['數量']}，輸入出庫數量")
            return

    reply_text(token, "找不到資料")


def process_out_qty(token, user_id, qty):
    if not qty.isdigit():
        reply_text(token, "請輸入數字")
        return

    row = user_data[user_id]["row"]
    col = get_col("數量")

    old = int(sheet.cell(row, col).value)
    out = int(qty)

    if out > old:
        reply_text(token, "庫存不足")
        return

    new = old - out
    sheet.update_cell(row, col, new)

    reply_text(token, f"出庫完成 {old} → {new}")
    clear_user_session(user_id)


def save_manual_stock(token, user_id):
    item = user_data[user_id]

    headers = sheet.row_values(1)
    new_row = [""] * len(headers)

    new_row[get_col("品名") - 1] = item["品名"]
    new_row[get_col("尺寸") - 1] = item["尺寸"]
    new_row[get_col("數量") - 1] = item["數量"]
    new_row[get_col("位置") - 1] = item["位置"]

    sheet.append_row(new_row)

    reply_text(token, "✅ 手動入庫完成")
    clear_user_session(user_id)


def get_col(name):
    headers = sheet.row_values(1)
    return headers.index(name) + 1


def clear_user_session(uid):
    user_state.pop(uid, None)
    user_data.pop(uid, None)


def reply_text(token, text):
    line_bot_api.reply_message(token, TextSendMessage(text=str(text)[:5000]))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)

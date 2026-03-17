import os
import json
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.models import (
    MessageEvent, TextMessage, TextSendMessage,
    TemplateSendMessage, ButtonsTemplate, MessageTemplateAction,
    CarouselTemplate, CarouselColumn
)
import gspread
from oauth2client.service_account import ServiceAccountCredentials

app = Flask(__name__)

LINE_CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET")
LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN")

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

creds_json = json.loads(os.getenv("GOOGLE_CREDENTIALS_JSON"))
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_json, scope)
gc = gspread.authorize(creds)

sheet = gc.open_by_key(os.getenv("GOOGLE_SHEET_ID")).sheet1

# 使用者狀態
user_state = {}

# 暫存使用者資料
user_data = {}


# Render / UptimeRobot / 外部監控用首頁
@app.route("/", methods=["GET"])
def home():
    return "LINE BOT Running", 200


@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers.get('X-Line-Signature')
    body = request.get_data(as_text=True)

    print("Request body:", body)

    try:
        handler.handle(body, signature)
    except Exception as e:
        print("callback error:", e)
        abort(400)

    return 'OK'


@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    user_text = str(event.message.text).strip()
    user_id = event.source.user_id

    print("收到訊息：", user_text)
    print("目前狀態：", user_state.get(user_id))
    print("暫存資料：", user_data.get(user_id))

    # 取消
    if user_text == "取消":
        clear_user_session(user_id)
        reply_text(event.reply_token, "已取消操作，請輸入「塊材查詢」開啟選單")
        return

    # 直接出庫格式：直接出庫::列號
    if user_text.startswith("直接出庫::"):
        try:
            row_number = int(user_text.split("::")[1])
            start_direct_out(event.reply_token, user_id, row_number)
        except Exception as e:
            print("direct out parse error:", e)
            reply_text(event.reply_token, "出庫資料錯誤，請重新查詢")
        return

    # 主選單
    if user_text == "塊材查詢":
        clear_user_session(user_id)
        send_menu(event.reply_token)
        return

    elif user_text == "查詢庫存":
        clear_user_session(user_id)
        user_state[user_id] = "waiting_search_keyword"
        reply_text(event.reply_token, "請輸入關鍵字，例如 503")
        return

    elif user_text == "全部庫存":
        clear_user_session(user_id)
        show_all_stock(event.reply_token)
        return

    elif user_text == "入庫":
        clear_user_session(user_id)
        user_state[user_id] = "waiting_in_keyword"
        reply_text(event.reply_token, "請輸入要入庫的品名或尺寸關鍵字")
        return

    elif user_text == "出庫":
        clear_user_session(user_id)
        user_state[user_id] = "waiting_out_keyword"
        reply_text(event.reply_token, "請輸入要出庫的品名或尺寸關鍵字")
        return

    elif user_text == "返回選單":
        clear_user_session(user_id)
        send_menu(event.reply_token)
        return

    elif user_text == "手動入庫":
        clear_user_session(user_id)
        user_data[user_id] = {}
        user_state[user_id] = "manual_in_name"
        reply_text(event.reply_token, "請輸入品名")
        return

    # 查詢流程
    if user_state.get(user_id) == "waiting_search_keyword":
        search_stock(event.reply_token, user_id, user_text)
        return

    # 入庫流程
    if user_state.get(user_id) == "waiting_in_keyword":
        search_stock_for_in(event.reply_token, user_id, user_text)
        return

    elif user_state.get(user_id) == "waiting_in_select":
        process_in_select(event.reply_token, user_id, user_text)
        return

    elif user_state.get(user_id) == "waiting_in_qty":
        process_in_qty(event.reply_token, user_id, user_text)
        return

    # 出庫流程
    if user_state.get(user_id) == "waiting_out_keyword":
        search_stock_for_out(event.reply_token, user_id, user_text)
        return

    elif user_state.get(user_id) == "waiting_out_select":
        process_out_select(event.reply_token, user_id, user_text)
        return

    elif user_state.get(user_id) == "waiting_out_qty":
        process_out_qty(event.reply_token, user_id, user_text)
        return

    # 手動入庫流程
    if user_state.get(user_id) == "manual_in_name":
        user_data[user_id]["品名"] = user_text
        user_state[user_id] = "manual_in_size"
        reply_text(event.reply_token, "請輸入尺寸")
        return

    elif user_state.get(user_id) == "manual_in_size":
        user_data[user_id]["尺寸"] = user_text
        user_state[user_id] = "manual_in_qty"
        reply_text(event.reply_token, "請輸入數量")
        return

    elif user_state.get(user_id) == "manual_in_qty":
        if not is_valid_int(user_text):
            reply_text(event.reply_token, "數量請輸入整數")
            return
        qty = int(user_text)
        if qty <= 0:
            reply_text(event.reply_token, "數量必須大於 0")
            return
        user_data[user_id]["數量"] = qty
        user_state[user_id] = "manual_in_loc"
        reply_text(event.reply_token, "請輸入位置")
        return

    elif user_state.get(user_id) == "manual_in_loc":
        user_data[user_id]["位置"] = user_text
        save_manual_stock(event.reply_token, user_id)
        return

    reply_text(event.reply_token, "請輸入「塊材查詢」開啟選單")


def clear_user_session(user_id):
    user_state.pop(user_id, None)
    user_data.pop(user_id, None)


def send_menu(reply_token):
    buttons = TemplateSendMessage(
        alt_text='塊材選單',
        template=ButtonsTemplate(
            title='塊材管理',
            text='請選擇功能',
            actions=[
                MessageTemplateAction(label='查詢庫存', text='查詢庫存'),
                MessageTemplateAction(label='入庫', text='入庫'),
                MessageTemplateAction(label='出庫', text='出庫'),
                MessageTemplateAction(label='全部庫存', text='全部庫存')
            ]
        )
    )
    line_bot_api.reply_message(reply_token, buttons)


def show_all_stock(reply_token):
    try:
        data = sheet.get_all_records()
        if not data:
            reply_text(reply_token, "目前沒有資料")
            return

        msg = "全部塊材庫存：\n"
        for row in data[:40]:
            msg += f"{row.get('品名', '')} / {row.get('尺寸', '')} / 數量:{row.get('數量', '')} / 位置:{row.get('位置', '')}\n"

        if len(msg) > 4500:
            msg = msg[:4500]

        reply_text(reply_token, msg)

    except Exception as e:
        print("show_all_stock error:", e)
        reply_text(reply_token, f"讀取失敗：{str(e)}")


def search_stock(reply_token, user_id, keyword):
    try:
        matches = find_matching_rows(keyword)

        if not matches:
            clear_user_session(user_id)
            reply_text(reply_token, "找不到相關資料")
            return

        lines = []
        for item in matches[:10]:
            lines.append(
                f"{item['row_number']}. {item['品名']} / {item['尺寸']} / 數量:{item['數量']} / 位置:{item['位置']}"
            )

        msg = "搜尋結果：\n" + "\n".join(lines)
        if len(msg) > 4500:
            msg = msg[:4500]

        clear_user_session(user_id)

        messages = [TextSendMessage(text=msg)]

        columns = []
        for item in matches[:10]:
            title = str(item['品名'])[:40] if item['品名'] else "庫存資料"
            text = f"{item['尺寸']}\n數量:{item['數量']} / 位置:{item['位置']}"
            text = text[:60]

            columns.append(
                CarouselColumn(
                    title=title,
                    text=text,
                    actions=[
                        MessageTemplateAction(
                            label='出庫',
                            text=f"直接出庫::{item['row_number']}"
                        )
                    ]
                )
            )

        messages.append(
            TemplateSendMessage(
                alt_text='搜尋結果出庫選單',
                template=CarouselTemplate(columns=columns)
            )
        )

        line_bot_api.reply_message(reply_token, messages)

    except Exception as e:
        print("search_stock error:", e)
        clear_user_session(user_id)
        reply_text(reply_token, f"查詢失敗：{str(e)}")


def start_direct_out(reply_token, user_id, row_number):
    try:
        data = sheet.get_all_records()

        target = None
        for idx, row in enumerate(data, start=2):
            if idx == row_number:
                target = {
                    "row_number": idx,
                    "品名": str(row.get("品名", "")).strip(),
                    "尺寸": str(row.get("尺寸", "")).strip(),
                    "數量": to_int(row.get("數量", 0)),
                    "位置": str(row.get("位置", "")).strip()
                }
                break

        if not target:
            clear_user_session(user_id)
            reply_text(reply_token, "找不到該筆資料，請重新查詢")
            return

        user_data[user_id] = {
            "selected_item": target
        }
        user_state[user_id] = "waiting_out_qty"

        reply_text(
            reply_token,
            f"已選擇出庫：\n"
            f"{target['品名']} / {target['尺寸']} / 目前數量:{target['數量']} / 位置:{target['位置']}\n\n"
            f"請直接輸入要扣除的數量"
        )

    except Exception as e:
        print("start_direct_out error:", e)
        clear_user_session(user_id)
        reply_text(reply_token, f"出庫操作失敗：{str(e)}")


def search_stock_for_in(reply_token, user_id, keyword):
    try:
        matches = find_matching_rows(keyword)

        if not matches:
            clear_user_session(user_id)
            line_bot_api.reply_message(reply_token, [
                TextSendMessage(text="找不到相關品名"),
                TemplateSendMessage(
                    alt_text='手動入庫選單',
                    template=ButtonsTemplate(
                        title='找不到品名',
                        text='是否要手動入庫？',
                        actions=[
                            MessageTemplateAction(label='手動入庫', text='手動入庫'),
                            MessageTemplateAction(label='返回選單', text='返回選單')
                        ]
                    )
                )
            ])
            return

        user_data[user_id] = {
            "matches": matches
        }
        user_state[user_id] = "waiting_in_select"

        msg = "以下為可入庫品項：\n"
        for i, item in enumerate(matches[:15], start=1):
            msg += f"{i}. {item['品名']} / {item['尺寸']} / 目前數量:{item['數量']} / 位置:{item['位置']}\n"
        msg += "\n請輸入要入庫的編號"
        msg += "\n若都不是，請輸入：手動入庫"
        msg += "\n取消請輸入：取消"

        if len(msg) > 4500:
            msg = msg[:4500]

        reply_text(reply_token, msg)

    except Exception as e:
        print("search_stock_for_in error:", e)
        clear_user_session(user_id)
        reply_text(reply_token, f"入庫查詢失敗：{str(e)}")


def process_in_select(reply_token, user_id, user_text):
    try:
        if user_text == "手動入庫":
            user_data[user_id] = {}
            user_state[user_id] = "manual_in_name"
            reply_text(reply_token, "請輸入品名")
            return

        if not is_valid_int(user_text):
            reply_text(reply_token, "請輸入正確編號")
            return

        selected_index = int(user_text) - 1
        matches = user_data.get(user_id, {}).get("matches", [])

        if selected_index < 0 or selected_index >= len(matches):
            reply_text(reply_token, "編號超出範圍，請重新輸入")
            return

        selected_item = matches[selected_index]
        user_data[user_id]["selected_item"] = selected_item
        user_state[user_id] = "waiting_in_qty"

        reply_text(
            reply_token,
            f"你選擇的是：\n"
            f"{selected_item['品名']} / {selected_item['尺寸']} / 目前數量:{selected_item['數量']} / 位置:{selected_item['位置']}\n\n"
            f"請輸入入庫數量"
        )

    except Exception as e:
        print("process_in_select error:", e)
        clear_user_session(user_id)
        reply_text(reply_token, f"入庫選擇失敗：{str(e)}")


def process_in_qty(reply_token, user_id, user_text):
    try:
        if not is_valid_int(user_text):
            reply_text(reply_token, "入庫數量請輸入整數")
            return

        add_qty = int(user_text)
        if add_qty <= 0:
            reply_text(reply_token, "入庫數量必須大於 0")
            return

        selected_item = user_data.get(user_id, {}).get("selected_item")
        if not selected_item:
            clear_user_session(user_id)
            reply_text(reply_token, "資料遺失，請重新操作")
            return

        row_num = selected_item["row_number"]
        qty_col = get_col_index("數量")

        old_qty = to_int(sheet.cell(row_num, qty_col).value)
        new_qty = old_qty + add_qty

        sheet.update_cell(row_num, qty_col, new_qty)

        reply_text(
            reply_token,
            f"入庫完成：\n"
            f"品名：{selected_item['品名']}\n"
            f"尺寸：{selected_item['尺寸']}\n"
            f"原數量：{old_qty}\n"
            f"入庫數量：{add_qty}\n"
            f"最新數量：{new_qty}"
        )

        clear_user_session(user_id)

    except Exception as e:
        print("process_in_qty error:", e)
        clear_user_session(user_id)
        reply_text(reply_token, f"入庫更新失敗：{str(e)}")


def search_stock_for_out(reply_token, user_id, keyword):
    try:
        matches = find_matching_rows(keyword)

        if not matches:
            clear_user_session(user_id)
            reply_text(reply_token, "找不到可出庫的品項")
            return

        user_data[user_id] = {
            "matches": matches
        }
        user_state[user_id] = "waiting_out_select"

        msg = "以下為可出庫品項：\n"
        for i, item in enumerate(matches[:15], start=1):
            msg += f"{i}. {item['品名']} / {item['尺寸']} / 目前數量:{item['數量']} / 位置:{item['位置']}\n"
        msg += "\n請輸入要出庫的編號"
        msg += "\n取消請輸入：取消"

        if len(msg) > 4500:
            msg = msg[:4500]

        reply_text(reply_token, msg)

    except Exception as e:
        print("search_stock_for_out error:", e)
        clear_user_session(user_id)
        reply_text(reply_token, f"出庫查詢失敗：{str(e)}")


def process_out_select(reply_token, user_id, user_text):
    try:
        if not is_valid_int(user_text):
            reply_text(reply_token, "請輸入正確編號")
            return

        selected_index = int(user_text) - 1
        matches = user_data.get(user_id, {}).get("matches", [])

        if selected_index < 0 or selected_index >= len(matches):
            reply_text(reply_token, "編號超出範圍，請重新輸入")
            return

        selected_item = matches[selected_index]
        user_data[user_id]["selected_item"] = selected_item
        user_state[user_id] = "waiting_out_qty"

        reply_text(
            reply_token,
            f"你選擇的是：\n"
            f"{selected_item['品名']} / {selected_item['尺寸']} / 目前數量:{selected_item['數量']} / 位置:{selected_item['位置']}\n\n"
            f"請輸入出庫數量"
        )

    except Exception as e:
        print("process_out_select error:", e)
        clear_user_session(user_id)
        reply_text(reply_token, f"出庫選擇失敗：{str(e)}")


def process_out_qty(reply_token, user_id, user_text):
    try:
        if not is_valid_int(user_text):
            reply_text(reply_token, "出庫數量請輸入整數")
            return

        out_qty = int(user_text)
        if out_qty <= 0:
            reply_text(reply_token, "出庫數量必須大於 0")
            return

        selected_item = user_data.get(user_id, {}).get("selected_item")
        if not selected_item:
            clear_user_session(user_id)
            reply_text(reply_token, "資料遺失，請重新操作")
            return

        row_num = selected_item["row_number"]
        qty_col = get_col_index("數量")

        old_qty = to_int(sheet.cell(row_num, qty_col).value)

        if out_qty > old_qty:
            reply_text(reply_token, f"出庫失敗：目前庫存只有 {old_qty}，不能出庫 {out_qty}")
            return

        new_qty = old_qty - out_qty
        sheet.update_cell(row_num, qty_col, new_qty)

        reply_text(
            reply_token,
            f"出庫完成：\n"
            f"品名：{selected_item['品名']}\n"
            f"尺寸：{selected_item['尺寸']}\n"
            f"原數量：{old_qty}\n"
            f"出庫數量：{out_qty}\n"
            f"最新數量：{new_qty}"
        )

        clear_user_session(user_id)

    except Exception as e:
        print("process_out_qty error:", e)
        clear_user_session(user_id)
        reply_text(reply_token, f"出庫更新失敗：{str(e)}")


def save_manual_stock(reply_token, user_id):
    try:
        item = user_data.get(user_id, {})
        name = str(item.get("品名", "")).strip()
        size = str(item.get("尺寸", "")).strip()
        qty = int(item.get("數量", 0))
        loc = str(item.get("位置", "")).strip()

        headers = sheet.row_values(1)
        new_row = [""] * len(headers)

        col_name = get_col_index("品名")
        col_size = get_col_index("尺寸")
        col_qty = get_col_index("數量")
        col_loc = get_col_index("位置")

        new_row[col_name - 1] = name
        new_row[col_size - 1] = size
        new_row[col_qty - 1] = qty
        new_row[col_loc - 1] = loc

        sheet.append_row(new_row)

        reply_text(
            reply_token,
            f"手動入庫完成：\n"
            f"品名：{name}\n"
            f"尺寸：{size}\n"
            f"數量：{qty}\n"
            f"位置：{loc}"
        )

        clear_user_session(user_id)

    except Exception as e:
        print("save_manual_stock error:", e)
        clear_user_session(user_id)
        reply_text(reply_token, f"手動入庫失敗：{str(e)}")


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


def get_col_index(header_name):
    headers = sheet.row_values(1)
    for i, h in enumerate(headers, start=1):
        if str(h).strip() == header_name:
            return i
    raise Exception(f"找不到欄位：{header_name}")


def to_int(value):
    try:
        if value is None or value == "":
            return 0
        return int(float(str(value).strip()))
    except:
        return 0


def is_valid_int(text):
    try:
        int(text)
        return True
    except:
        return False


def reply_text(reply_token, text):
    line_bot_api.reply_message(reply_token, TextSendMessage(text=text))


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

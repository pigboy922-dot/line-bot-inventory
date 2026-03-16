import os
import json
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.models import MessageEvent, TextMessage, TextSendMessage, TemplateSendMessage, ButtonsTemplate, MessageTemplateAction
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

    if user_text == "塊材查詢":
        send_menu(event.reply_token)
        return

    elif user_text == "查詢庫存":
        reply_text(event.reply_token, "請輸入關鍵字，例如 503")
        return

    elif user_text == "全部庫存":
        show_all_stock(event.reply_token)
        return

    elif user_text == "入庫":
        reply_text(event.reply_token, "請輸入要入庫的品名關鍵字")
        return

    elif user_text == "出庫":
        reply_text(event.reply_token, "請輸入要出庫的品名關鍵字")
        return

    else:
        search_stock(event.reply_token, user_text)
        return


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
        msg = "全部塊材庫存：\n"

        for row in data[:40]:
            msg += f"{row.get('品名','')} / {row.get('尺寸','')} / 數量:{row.get('數量','')} / 位置:{row.get('位置','')}\n"

        line_bot_api.reply_message(reply_token, TextSendMessage(text=msg))
    except Exception as e:
        print("show_all_stock error:", e)
        line_bot_api.reply_message(reply_token, TextSendMessage(text=f"讀取失敗：{str(e)}"))


def search_stock(reply_token, keyword):
    try:
        data = sheet.get_all_records()
        keyword = str(keyword).strip().lower()

        result = []
        for row in data:
            name = str(row.get("品名", "")).strip().lower()
            size = str(row.get("尺寸", "")).strip()
            qty = str(row.get("數量", "")).strip()
            loc = str(row.get("位置", "")).strip()

            if keyword in name:
                result.append(f"{row.get('品名', '')} / {size} / 數量:{qty} / 位置:{loc}")

        if not result:
            line_bot_api.reply_message(reply_token, TextSendMessage(text="找不到相關品名"))
            return

        msg = "搜尋結果：\n" + "\n".join(result[:20])
        line_bot_api.reply_message(reply_token, TextSendMessage(text=msg))

    except Exception as e:
        print("search_stock error:", e)
        line_bot_api.reply_message(reply_token, TextSendMessage(text=f"查詢失敗：{str(e)}"))


def reply_text(reply_token, text):
    line_bot_api.reply_message(reply_token, TextSendMessage(text=text))


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    user_text = str(event.message.text).strip()
    user_id = get_user_key(event)

    print("收到訊息：", user_text)
    print("目前狀態：", user_state.get(user_id))
    print("暫存資料：", user_data.get(user_id))

    # ⭐ 未啟動前完全不回應
    if user_id not in user_state and user_text != "塊材查詢":
        return

    # ===== 取消 =====
    if user_text == "取消":
        clear_user_session(user_id)
        reply_text(event.reply_token, "已取消操作，請輸入「塊材查詢」開啟選單")
        return

    # ===== 主選單 =====
    if user_text == "塊材查詢":
        clear_user_session(user_id)
        send_menu(event.reply_token)
        return

    # ===== 其它原本程式完全不動 =====

    # 直接出庫格式：直接出庫::列號
    if user_text.startswith("直接出庫::"):
        try:
            row_number = int(user_text.split("::")[1])
            start_direct_out(event.reply_token, user_id, row_number)
        except:
            reply_text(event.reply_token, "出庫資料錯誤")
        return

    if user_text.startswith("直接入庫::"):
        try:
            row_number = int(user_text.split("::")[1])
            start_direct_in(event.reply_token, user_id, row_number)
        except:
            reply_text(event.reply_token, "入庫資料錯誤")
        return

    if user_text.startswith("全部庫存::"):
        try:
            page = int(user_text.split("::")[1])
            clear_user_session(user_id)
            show_all_stock(event.reply_token, page=page)
        except:
            reply_text(event.reply_token, "頁碼錯誤")
        return

    elif user_text == "查詢庫存":
        clear_user_session(user_id)
        user_state[user_id] = "waiting_search_keyword"
        reply_text(event.reply_token, "請輸入關鍵字，例如 503")
        return

    elif user_text == "全部庫存":
        clear_user_session(user_id)
        show_all_stock(event.reply_token, page=1)
        return

    elif user_text == "入庫":
        clear_user_session(user_id)
        user_state[user_id] = "waiting_in_keyword"
        reply_text(event.reply_token, "請輸入關鍵字")
        return

    elif user_text == "出庫":
        clear_user_session(user_id)
        user_state[user_id] = "waiting_out_keyword"
        reply_text(event.reply_token, "請輸入關鍵字")
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

    # ===== 查詢 =====
    if user_state.get(user_id) == "waiting_search_keyword":
        search_stock(event.reply_token, user_id, user_text)
        return

    # ===== 入庫 =====
    if user_state.get(user_id) == "waiting_in_keyword":
        search_stock_for_in(event.reply_token, user_id, user_text)
        return

    elif user_state.get(user_id) == "waiting_in_qty":
        process_in_qty(event.reply_token, user_id, user_text)
        return

    # ===== 出庫 =====
    if user_state.get(user_id) == "waiting_out_keyword":
        search_stock_for_out(event.reply_token, user_id, user_text)
        return

    elif user_state.get(user_id) == "waiting_out_select":
        process_out_select(event.reply_token, user_id, user_text)
        return

    elif user_state.get(user_id) == "waiting_out_qty":
        process_out_qty(event.reply_token, user_id, user_text)
        return

    # ===== 手動入庫（你原本的）=====
    if user_state.get(user_id) == "manual_in_name":
        user_data[user_id]["品名"] = user_text
        user_state[user_id] = "manual_in_qty"
        reply_text(event.reply_token, "請輸入數量")
        return

    elif user_state.get(user_id) == "manual_in_qty":
        if not is_valid_int(user_text):
            reply_text(event.reply_token, "數量請輸入整數")
            return

        qty = int(user_text)
        user_data[user_id]["數量"] = qty
        user_state[user_id] = "manual_in_loc"
        reply_text(event.reply_token, "請輸入位置")
        return

    elif user_state.get(user_id) == "manual_in_loc":
        user_data[user_id]["位置"] = user_text
        save_manual_stock(event.reply_token, user_id)
        return

    # ❌ 最後這行已刪除（關鍵）
    return

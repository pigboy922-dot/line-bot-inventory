LINE Google Sheet 雲端版庫存機器人 安裝教學

一、功能
1. 查詢 509
2. 入庫 509
3. 出庫 509
4. 互動流程：選品名 / 輸入尺寸 / 位置 / 備註 / 數量 / 確認
5. 資料自動寫入 Google Sheet
6. 可部署到雲端，電腦不用開著

二、Google Sheet 準備
1. 建立一份新的 Google Sheet
2. 複製 Sheet 網址中的試算表 ID
   例如：https://docs.google.com/spreadsheets/d/這段就是ID/edit
3. 建議建立兩個工作表：
   - inventory
   - transactions
   如果沒建立也沒關係，程式會自動建立

三、Google Service Account 準備
1. 到 Google Cloud 建立專案
2. 啟用 Google Sheets API 與 Google Drive API
3. 建立 Service Account
4. 建立 JSON 金鑰並下載
5. 用記事本打開 JSON 檔，整份內容複製起來
6. 把 service account 的 client_email 加到 Google Sheet 共用名單，權限給編輯者

四、LINE Developers 準備
1. 建立 Messaging API channel
2. 取得以下兩個值
   - Channel secret
   - Channel access token
3. 開啟：Use webhook
4. 開啟：Allow bot to join group chats
5. 關閉：Auto-reply messages

五、Render 雲端部署
1. 把本資料夾上傳到 GitHub
2. 到 Render 建立新的 Web Service
3. 連接你的 GitHub 專案
4. Build Command：
   pip install -r requirements.txt
5. Start Command：
   python app.py
6. 在 Render 的 Environment Variables 新增：
   - LINE_CHANNEL_SECRET
   - LINE_CHANNEL_ACCESS_TOKEN
   - GOOGLE_SHEET_ID
   - GOOGLE_SERVICE_ACCOUNT_JSON
   - APP_TIMEZONE（可填 Asia/Taipei）

六、Webhook 設定
1. Render 部署成功後會得到網址，例如：
   https://你的專案名稱.onrender.com
2. 到 LINE Developers 的 Webhook URL 填入：
   https://你的專案名稱.onrender.com/callback
3. 按 Verify，成功即可

七、指令範例
1. 查詢 509
2. 入庫 509
3. 出庫 509
4. 取消

八、Google Sheet 欄位
inventory：
品名 | 尺寸 | 庫存 | 位置 | 備註 | 更新時間

transactions：
時間 | 類型 | 品名 | 尺寸 | 數量 | 位置 | 備註 | 使用者

九、常見問題
1. LINE 沒回應
   - 檢查 Webhook URL 是否正確
   - 檢查 Use webhook 是否開啟
   - 檢查 Auto-reply messages 是否關閉
2. Google Sheet 寫不進去
   - 檢查 GOOGLE_SERVICE_ACCOUNT_JSON 是否完整
   - 檢查 Sheet 是否已分享給 service account 的 client_email
3. 出庫失敗
   - 可能是庫存不足

十、你之後可再加的功能
1. 多倉位查詢
2. 低庫存警示
3. 管理員權限
4. 日報表 / 月報表
5. 圖文選單

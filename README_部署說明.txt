1. 把 app.py、requirements.txt、render.yaml 上傳到 Render 專案。
2. Render 環境變數新增：
   GOOGLE_SHEET_ID=1q3cbcckBaPOMz78lkt4-ko-gSwxQcvq2uA9M8r7feZI
   LIFF_ID=2009630662-Qeb3Wh1t
   GOOGLE_CREDENTIALS_JSON=整段 service account JSON（單行）
3. 部署完成後，先打開：
   https://你的-render網址/health
   確認 ok=true。
4. 再把 liff_inventory_mobile_full.html 內的：
   API_BASE: 'https://YOUR-RENDER-SERVICE.onrender.com'
   改成你的 Render 網址。
5. 把更新後的 HTML 覆蓋上傳到 GitHub Pages。
6. LINE Developers 的 LIFF Endpoint URL 指向 GitHub Pages 的 HTML 網址。
7. Google Sheet 第一列必須有：品名、尺寸、數量、位置。
8. service account 的 client_email 要共享成該 Google Sheet 的編輯者。

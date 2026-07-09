"""
一次性本機執行：用專門的 Google 帳號（例如 mobanchucun@gmail.com）授權，
取得可以長期使用的 Drive refresh_token，讓 App 改用「真人帳號」的 Drive 額度
上傳模板 Excel，不再透過沒有儲存額度的 service account（會 403）。

使用步驟：
  1. pip install google-auth-oauthlib
  2. 到 Google Cloud Console（跟 service account 同一個專案）：
     API 和服務 → 憑證 → 建立憑證 → OAuth 用戶端 ID
     應用程式類型選「電腦版應用程式」，建立後下載 JSON，
     存成跟這支程式同一層的 client_secret.json
  3. python gen_drive_refresh_token.py
  4. 瀏覽器會開啟登入畫面，登入要拿來存模板的帳號（例如 mobanchucun@gmail.com），
     同意權限即可（畫面顯示「Google 未驗證這個應用程式」是正常的，
     點「進階」→「前往（不安全）」繼續即可，因為這是你自己的 OAuth 用戶端）
  5. 終端機會印出 client_id / client_secret / refresh_token，
     把這三行貼到 secrets.toml 的 [gcp_oauth] 區塊
"""
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ["https://www.googleapis.com/auth/drive"]

flow = InstalledAppFlow.from_client_secrets_file("client_secret.json", SCOPES)
creds = flow.run_local_server(port=0)

print("\n=== 把以下內容加到 secrets.toml 的 [gcp_oauth] 區塊 ===\n")
print("[gcp_oauth]")
print(f'client_id = "{creds.client_id}"')
print(f'client_secret = "{creds.client_secret}"')
print(f'refresh_token = "{creds.refresh_token}"')

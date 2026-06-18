# 出貨自動化工具

## 本地啟動

```bash
pip install -r requirements.txt
streamlit run app.py
```

本地執行前，先把 `.streamlit/secrets.toml` 填入真實的 Google 憑證。

---

## 部署到 Streamlit Cloud

1. 把這個資料夾推到 GitHub（私有 repo）
2. 到 https://share.streamlit.io 連結 repo，main file 選 `app.py`
3. 在 App Settings → **Secrets** 貼入以下內容（換成真實值）：

```toml
[gcp_service_account]
type = "service_account"
project_id = "erptoexcel"
private_key_id = "07b1db8a..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "erp-995@erptoexcel.iam.gserviceaccount.com"
client_id = "117422..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/erp-995%40erptoexcel.iam.gserviceaccount.com"
```

4. Deploy → 拿到網址給同事用

---

## 檔案說明

| 檔案 | 說明 |
|------|------|
| `app.py` | 主介面 |
| `rtf_parser.py` | 解析銷貨單 RTF |
| `sheets_db.py` | Google Sheets 讀寫 |
| `module_b_invoice.py` | 電子發票產生 |
| `module_c_labels.py` | 標籤產生（含欄位自訂） |
| `.streamlit/secrets.toml` | 憑證設定（不推 GitHub）|

---

## 注意事項

- `.streamlit/secrets.toml` 已加入 `.gitignore`，**不會**推到 GitHub
- Google Service Account 的 JSON key 請定期輪換（建議每年）
- Streamlit Cloud 免費版閒置 7 天會睡眠，第一個使用者開網頁時等約 30 秒喚醒

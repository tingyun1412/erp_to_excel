# 出貨自動化工具

## 本地啟動

```bash
pip install -r requirements.txt
streamlit run app.py
```

執行前，請先於 `.streamlit/secrets.toml` 設定 ERP 資料庫連線資訊：

```toml
[mssql]
server   = "ntserver"
database = "24405403"
user     = "xxx"
password = "xxx"
```

註：此資料庫僅限公司內網存取，須於公司內或連接 VPN 後才能連線；「查詢銷貨單」功能僅支援本機執行。

模板、廠商帳號、發票跳過名單、發票開立紀錄、出貨提醒等資料均已改為本機檔案儲存
（詳見 `src/local_db.py`、`src/local_template_store.py`），不再依賴 Google Sheets / Drive，
亦無需設定任何 Google 憑證。

---

## 部署架構

目前採用「正本存放於公司網路磁碟、各終端機執行本機備份」的部署方式：

1. 將完整資料夾（含封裝完成的 `python/` 內嵌執行環境、`ms-playwright/` 瀏覽器元件，
   以及 `app/` 程式碼）置於公司網路磁碟。
2. 各使用者於本機電腦執行網路磁碟中的 `install_local.bat`：
   - 透過 robocopy 將 `app`、`python`、`ms-playwright` 同步至 `%LOCALAPPDATA%\ShippingAutomationTool`。
   - 於桌面建立「ShippingTool」捷徑，以 `pythonw.exe` 執行 `webview_launcher.py`。
   - 網路磁碟版本更新後，重新執行一次 `install_local.bat` 即可同步至最新版本。
3. 啟動捷徑後，`webview_launcher.py` 會於背景啟動本機 Streamlit 伺服器
   （僅監聽 127.0.0.1，設定見 `.streamlit/config.toml`），並以 pywebview 包裝為
   原生視窗介面（無網址列、無終端機視窗）；執行紀錄輸出至 `logs/app.log`。

---

## 目錄結構

```
app.py                  進入點：Streamlit 主介面（streamlit run app.py）
webview_launcher.py     進入點：本機啟動器，將 Streamlit 封裝為原生視窗
install_local.bat       進入點：部署腳本（由網路磁碟根目錄執行）
requirements.txt
packages.txt
.streamlit/
  config.toml
  secrets.toml          憑證設定檔（不納入版本控制）
src/                    核心邏輯模組
  erp_db.py
  erp_downloader.py
  slp_downloader.py
  lscr_parser.py
  einvoice_submitter.py
  local_db.py
  local_template_store.py
  template_engine.py
  module_b_invoice.py
  module_d_report.py
  rtf_parser.py
```

上述三個進入點（`app.py`、`webview_launcher.py`、`install_local.bat`）皆被外部部署流程／桌面捷徑直接指向，故維持於根目錄；其餘業務邏輯模組統一收納於 `src/`。

## 檔案說明

| 檔案 | 說明 |
|------|------|
| `app.py` | 主介面 |
| `webview_launcher.py` | 本機啟動器，將 Streamlit 封裝為原生視窗（pywebview） |
| `install_local.bat` | 部署腳本，將程式由網路磁碟同步至本機並建立桌面捷徑 |
| `src/erp_db.py` | 依單據號碼查詢 ERP 資料庫（批號／單價／金額／客戶資料） |
| `src/erp_downloader.py` | 登入 ERP 網站並自動下載標籤 PDF |
| `src/slp_downloader.py` | 均華（GMM）供應商平台標籤下載，擷取原始網頁畫面組成 PDF |
| `src/lscr_parser.py` | 解析 LSCR 出貨明細確認單 |
| `src/einvoice_submitter.py` | 電子發票逐張自動送出（e-invoice.com.tw） |
| `src/local_db.py` | 廠商帳號／發票跳過名單／發票開立紀錄／出貨提醒之本機檔案儲存 |
| `src/local_template_store.py` | 標籤模板／LSCR 基礎模板之本機檔案儲存 |
| `src/template_engine.py` | 標籤模板引擎，支援多種模板格式 |
| `src/module_b_invoice.py` | 模組 B：電子發票產生 |
| `src/module_d_report.py` | 模組 D：生產日報表彙總 |
| `src/rtf_parser.py` | 銷貨單 RTF 解析；主流程已改查 ERP 資料庫，目前僅供 `erp_db.py` 呼叫其中的中英文前綴切分工具函式 |
| `.streamlit/secrets.toml` | 憑證設定檔（不納入版本控制） |

---

## 注意事項

- `.streamlit/secrets.toml`、`local_data/`、`files/`、`logs/` 均已列入 `.gitignore`，不會推送至 GitHub。
- 應用程式僅監聽 127.0.0.1，同一網路內其他電腦無法連線存取（詳見 `.streamlit/config.toml` 內註解）。

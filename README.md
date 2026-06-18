# 倉庫出貨自動化工具

## 快速啟動

```bash
pip install -r requirements.txt
streamlit run app.py
```

瀏覽器會自動開啟 http://localhost:8501

---

## 檔案結構

```
erp_app/
├── app.py                  # Streamlit 主介面
├── rtf_parser.py           # 解析銷貨單 RTF
├── module_a_calendar.py    # 模組A：產出出貨行事曆
├── module_b_invoice.py     # 模組B：產出電子發票 Excel
├── module_c_labels.py      # 模組C：產出出貨標籤 Excel
└── requirements.txt
```

---

## 新增標籤格式（給開發者）

在 `module_c_labels.py` 找到 `LABEL_TEMPLATES` 字典：

```python
LABEL_TEMPLATES: dict[str, callable] = {
    "標準格式（含完整資訊）":     _label_standard,
    "簡易格式（料號+品名+數量）": _label_simple,
    # ↓ 在這裡加新格式
    "你的新格式名稱":            your_new_function,
}
```

新增步驟：
1. 寫一個 function，簽名：`def your_fn(ws, row_start, item, order) -> int`
   - `ws`：openpyxl Worksheet
   - `row_start`：從第幾行開始畫這張標籤
   - `item`：品項 dict（含 item_no, description, quantity, unit, customer, ship_date, remark）
   - `order`：銷貨單 dict（含 order_no, order_date, customer_order_no, ...）
   - 回傳值：下一張標籤應該從哪行開始
2. 把 function 加到 `LABEL_TEMPLATES`
3. 重啟 `streamlit run app.py`，下拉選單就會出現新格式

---

## 部署選項

### 選項 A：Streamlit Community Cloud（免費）
1. 把 `erp_app/` 資料夾推到 GitHub
2. 到 https://share.streamlit.io 連結 repo
3. 取得公開網址給同事使用

### 選項 B：打包成 .exe
```bash
pip install pyinstaller
pyinstaller --onefile --add-data "*.py;." app.py
```
（需要在有完整 Python 環境的 Windows 機器上執行）

---

## 已知限制

- RTF 解析依賴檔名格式（`YYYYMMDDXXXX-客戶.rtf`）來抓銷貨單號
- 若 RTF 內部欄位值解析不到，可在「銷貨單預覽」tab 手動補充
- 電子發票的單價/金額需手動補（RTF 格式確認後可改進解析邏輯）

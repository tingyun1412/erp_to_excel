"""
Google Sheets 資料庫層
負責所有對 Google Sheets 的讀寫操作
"""
import json
from datetime import datetime
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SPREADSHEET_ID = "15RRGc0Kmxr6w8cithjEOYJGnmPuWFy9eTBXElVBaR0Y"

SHEET_SCHEDULE  = "出貨排程"
SHEET_LABELS    = "廠商標籤設定"
SHEET_ORDERS    = "銷貨單主檔"

SCHEDULE_HEADERS = [
    "銷貨單號", "出貨日期", "客戶名稱", "料號", "品名",
    "數量", "單位", "客戶料號", "客戶訂單號", "狀態", "備註", "匯入時間"
]
LABELS_HEADERS = ["廠商名稱", "欄位順序", "最後更新"]
ORDERS_HEADERS = [
    "銷貨單號", "銷貨日期", "賣方統編", "買方統編",
    "客戶名稱", "客戶訂單號", "聯絡人", "匯入時間"
]


@st.cache_resource
def get_client():
    """建立並快取 gspread client"""
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)


def get_sheet(sheet_name: str):
    client = get_client()
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    try:
        return spreadsheet.worksheet(sheet_name)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=20)
        # 寫入標題列
        headers_map = {
            SHEET_SCHEDULE: SCHEDULE_HEADERS,
            SHEET_LABELS:   LABELS_HEADERS,
            SHEET_ORDERS:   ORDERS_HEADERS,
        }
        if sheet_name in headers_map:
            ws.append_row(headers_map[sheet_name])
        return ws


# ── 出貨排程 ──────────────────────────────────────────────────────

def load_schedule() -> list[dict]:
    ws = get_sheet(SHEET_SCHEDULE)
    records = ws.get_all_records()
    return records


def append_schedule_rows(rows: list[dict]):
    """新增出貨排程資料，跳過已存在的銷貨單號+料號組合"""
    ws = get_sheet(SHEET_SCHEDULE)
    existing = ws.get_all_records()
    existing_keys = {
        (str(r.get("銷貨單號", "")), str(r.get("料號", "")))
        for r in existing
    }
    now = datetime.now().strftime("%Y/%m/%d %H:%M")
    added = 0
    for row in rows:
        key = (str(row.get("銷貨單號", "")), str(row.get("料號", "")))
        if key in existing_keys:
            continue
        ws.append_row([
            row.get("銷貨單號", ""),
            row.get("出貨日期", ""),
            row.get("客戶名稱", ""),
            row.get("料號", ""),
            row.get("品名", ""),
            row.get("數量", ""),
            row.get("單位", "PC"),
            row.get("客戶料號", ""),
            row.get("客戶訂單號", ""),
            row.get("狀態", "待出貨"),
            row.get("備註", ""),
            now,
        ])
        existing_keys.add(key)
        added += 1
    return added


def update_schedule_status(row_index: int, status: str, remark: str = ""):
    """更新指定列的狀態（row_index 從 1 開始，不含標題）"""
    ws = get_sheet(SHEET_SCHEDULE)
    data_row = row_index + 1  # +1 因為有標題列
    status_col  = SCHEDULE_HEADERS.index("狀態") + 1
    remark_col  = SCHEDULE_HEADERS.index("備註") + 1
    ws.update_cell(data_row, status_col, status)
    if remark:
        ws.update_cell(data_row, remark_col, remark)


# ── 銷貨單主檔 ────────────────────────────────────────────────────

def load_orders() -> list[dict]:
    ws = get_sheet(SHEET_ORDERS)
    return ws.get_all_records()


def append_order(order: dict):
    ws = get_sheet(SHEET_ORDERS)
    existing = ws.get_all_records()
    if any(str(r.get("銷貨單號")) == str(order.get("order_no")) for r in existing):
        return False  # 已存在
    now = datetime.now().strftime("%Y/%m/%d %H:%M")
    ws.append_row([
        order.get("order_no", ""),
        order.get("order_date", ""),
        order.get("seller_tax_id", ""),
        order.get("buyer_tax_id", ""),
        order.get("filename", "").split("-")[1].split(".")[0] if "-" in order.get("filename","") else "",
        order.get("customer_order_no", ""),
        order.get("contact", ""),
        now,
    ])
    return True


# ── 廠商標籤設定 ──────────────────────────────────────────────────

def load_label_config(customer: str) -> list[str] | None:
    """讀取廠商的標籤欄位順序，找不到回傳 None"""
    ws = get_sheet(SHEET_LABELS)
    records = ws.get_all_records()
    for r in records:
        if r.get("廠商名稱") == customer:
            fields_str = r.get("欄位順序", "")
            if fields_str:
                return [f.strip() for f in fields_str.split(",")]
    return None


def save_label_config(customer: str, fields: list[str]):
    """儲存廠商的標籤欄位順序"""
    ws = get_sheet(SHEET_LABELS)
    records = ws.get_all_records()
    now = datetime.now().strftime("%Y/%m/%d %H:%M")
    fields_str = ", ".join(fields)

    for i, r in enumerate(records):
        if r.get("廠商名稱") == customer:
            data_row = i + 2  # +1 標題 +1 因為 enumerate 從 0
            ws.update_cell(data_row, 2, fields_str)
            ws.update_cell(data_row, 3, now)
            return

    ws.append_row([customer, fields_str, now])


def load_all_label_configs() -> dict[str, list[str]]:
    """讀取所有廠商的標籤設定"""
    ws = get_sheet(SHEET_LABELS)
    records = ws.get_all_records()
    result = {}
    for r in records:
        name = r.get("廠商名稱", "")
        fields_str = r.get("欄位順序", "")
        if name and fields_str:
            result[name] = [f.strip() for f in fields_str.split(",")]
    return result

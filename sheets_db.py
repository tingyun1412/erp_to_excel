"""
Google Sheets 資料庫層
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

SHEET_SCHEDULE   = "出貨排程"
SHEET_LABELS     = "廠商標籤設定"
SHEET_ORDERS     = "銷貨單主檔"
SHEET_TEMPLATES  = "標籤模板"

SCHEDULE_HEADERS  = ["銷貨單號","出貨日期","客戶名稱","料號","品名","數量","單位","客戶料號","客戶訂單號","狀態","備註","匯入時間"]
LABELS_HEADERS    = ["廠商名稱","欄位順序","最後更新"]
ORDERS_HEADERS    = ["銷貨單號","銷貨日期","賣方統編","買方統編","客戶名稱","客戶訂單號","聯絡人","匯入時間"]
TEMPLATES_HEADERS = ["廠商名稱","模板名稱","設定JSON","最後更新"]

HEADERS_MAP = {
    SHEET_SCHEDULE:  SCHEDULE_HEADERS,
    SHEET_LABELS:    LABELS_HEADERS,
    SHEET_ORDERS:    ORDERS_HEADERS,
    SHEET_TEMPLATES: TEMPLATES_HEADERS,
}


@st.cache_resource
def get_client():
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
        if sheet_name in HEADERS_MAP:
            ws.append_row(HEADERS_MAP[sheet_name])
        return ws


def _rows_to_records(rows: list[list], headers: list[str]) -> list[dict]:
    """
    把 get_all_values() 轉成 dict list，自動偵測是否有標題列。
    """
    if not rows:
        return []
    first_row = rows[0]
    start = 1 if (first_row and str(first_row[0]).strip() == headers[0]) else 0
    records = []
    for row in rows[start:]:
        if not any(str(c).strip() for c in row):
            continue
        record = {h: (row[i] if i < len(row) else "") for i, h in enumerate(headers)}
        records.append(record)
    return records


# ── 出貨排程 ──────────────────────────────────────────────────────

def load_schedule() -> list[dict]:
    ws = get_sheet(SHEET_SCHEDULE)
    rows = ws.get_all_values()
    return _rows_to_records(rows, SCHEDULE_HEADERS)


def append_schedule_rows(rows: list[dict]):
    ws = get_sheet(SHEET_SCHEDULE)
    existing = _rows_to_records(ws.get_all_values(), SCHEDULE_HEADERS)
    existing_keys = {
        (str(r.get("銷貨單號","")), str(r.get("料號","")))
        for r in existing
    }
    now = datetime.now().strftime("%Y/%m/%d %H:%M")
    added = 0
    for row in rows:
        key = (str(row.get("銷貨單號","")), str(row.get("料號","")))
        if key in existing_keys:
            continue
        ws.append_row([
            row.get("銷貨單號",""), row.get("出貨日期",""), row.get("客戶名稱",""),
            row.get("料號",""), row.get("品名",""), row.get("數量",""),
            row.get("單位","PC"), row.get("客戶料號",""), row.get("客戶訂單號",""),
            row.get("狀態","待出貨"), row.get("備註",""), now,
        ])
        existing_keys.add(key)
        added += 1
    return added


def update_schedule_status(row_index: int, status: str, remark: str = ""):
    ws = get_sheet(SHEET_SCHEDULE)
    data_row = row_index + 1
    ws.update_cell(data_row, SCHEDULE_HEADERS.index("狀態") + 1, status)
    if remark:
        ws.update_cell(data_row, SCHEDULE_HEADERS.index("備註") + 1, remark)


# ── 銷貨單主檔 ────────────────────────────────────────────────────

def load_orders() -> list[dict]:
    ws = get_sheet(SHEET_ORDERS)
    rows = ws.get_all_values()
    return _rows_to_records(rows, ORDERS_HEADERS)


def append_order(order: dict):
    ws = get_sheet(SHEET_ORDERS)
    existing = _rows_to_records(ws.get_all_values(), ORDERS_HEADERS)
    if any(str(r.get("銷貨單號")) == str(order.get("order_no")) for r in existing):
        return False
    now = datetime.now().strftime("%Y/%m/%d %H:%M")
    customer = order.get("filename","").split("-")[1].split(".")[0] if "-" in order.get("filename","") else ""
    ws.append_row([
        order.get("order_no",""), order.get("order_date",""),
        order.get("seller_tax_id",""), order.get("buyer_tax_id",""),
        customer, order.get("customer_order_no",""),
        order.get("contact",""), now,
    ])
    return True


# ── 廠商標籤設定（欄位偏好） ──────────────────────────────────────

def load_label_config(customer: str) -> list[str] | None:
    ws = get_sheet(SHEET_LABELS)
    records = _rows_to_records(ws.get_all_values(), LABELS_HEADERS)
    for r in records:
        if r.get("廠商名稱") == customer:
            fields_str = r.get("欄位順序","")
            if fields_str:
                return [f.strip() for f in fields_str.split(",")]
    return None


def save_label_config(customer: str, fields: list[str]):
    ws = get_sheet(SHEET_LABELS)
    records = _rows_to_records(ws.get_all_values(), LABELS_HEADERS)
    now = datetime.now().strftime("%Y/%m/%d %H:%M")
    fields_str = ", ".join(fields)
    for i, r in enumerate(records):
        if r.get("廠商名稱") == customer:
            ws.update_cell(i + 2, 2, fields_str)
            ws.update_cell(i + 2, 3, now)
            return
    ws.append_row([customer, fields_str, now])


def load_all_label_configs() -> dict[str, list[str]]:
    ws = get_sheet(SHEET_LABELS)
    records = _rows_to_records(ws.get_all_values(), LABELS_HEADERS)
    result = {}
    for r in records:
        name = r.get("廠商名稱","")
        fields_str = r.get("欄位順序","")
        if name and fields_str:
            result[name] = [f.strip() for f in fields_str.split(",")]
    return result


# ── 標籤模板 ──────────────────────────────────────────────────────

def load_templates(customer: str = "") -> list[dict]:
    ws = get_sheet(SHEET_TEMPLATES)
    records = _rows_to_records(ws.get_all_values(), TEMPLATES_HEADERS)
    if customer:
        return [r for r in records if r.get("廠商名稱") == customer]
    return records


def save_template(customer: str, template_name: str, config_json: str):
    ws = get_sheet(SHEET_TEMPLATES)
    records = _rows_to_records(ws.get_all_values(), TEMPLATES_HEADERS)
    now = datetime.now().strftime("%Y/%m/%d %H:%M")
    for i, r in enumerate(records):
        if r.get("廠商名稱") == customer and r.get("模板名稱") == template_name:
            ws.update_cell(i + 2, 3, config_json)
            ws.update_cell(i + 2, 4, now)
            return
    ws.append_row([customer, template_name, config_json, now])


def delete_template(customer: str, template_name: str):
    ws = get_sheet(SHEET_TEMPLATES)
    records = _rows_to_records(ws.get_all_values(), TEMPLATES_HEADERS)
    for i, r in enumerate(records):
        if r.get("廠商名稱") == customer and r.get("模板名稱") == template_name:
            ws.delete_rows(i + 2)
            return
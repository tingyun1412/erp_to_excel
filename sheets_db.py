"""
Google Sheets 資料庫層
只保留：標籤模板管理
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

SPREADSHEET_ID   = "15RRGc0Kmxr6w8cithjEOYJGnmPuWFy9eTBXElVBaR0Y"
SHEET_TEMPLATES  = "標籤模板"
TEMPLATES_HEADERS = ["廠商名稱", "模板名稱", "設定JSON", "最後更新"]


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
        if sheet_name == SHEET_TEMPLATES:
            ws.append_row(TEMPLATES_HEADERS)
        return ws


def _rows_to_records(rows: list[list], headers: list[str]) -> list[dict]:
    """
    把 get_all_values() 轉成 dict list。
    若第一列是標題列，用實際標題對應欄位（容錯欄位順序）。
    """
    if not rows:
        return []
    first_row = [str(c).strip() for c in rows[0]]
    is_header_row = any(h in first_row for h in headers)
    if is_header_row:
        col_map = {cell: i for i, cell in enumerate(first_row) if cell}
        start = 1
    else:
        col_map = {h: i for i, h in enumerate(headers)}
        start = 0
    records = []
    for row in rows[start:]:
        if not any(str(c).strip() for c in row):
            continue
        record = {
            h: str(row[col_map[h]]).strip() if col_map.get(h, -1) >= 0 and col_map[h] < len(row) else ""
            for h in headers
        }
        records.append(record)
    return records


# ── 標籤模板 ──────────────────────────────────────────────────────

@st.cache_data(ttl=120)
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


def clear_cache():
    load_templates.clear()
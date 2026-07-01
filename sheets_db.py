"""
Google Sheets 資料庫層
標籤模板管理 + 模板 Excel 存 Google Drive + 廠商帳號管理
"""
import time
from datetime import datetime, timezone, timedelta

_TW = timezone(timedelta(hours=8))
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SPREADSHEET_ID    = "15RRGc0Kmxr6w8cithjEOYJGnmPuWFy9eTBXElVBaR0Y"
SHEET_TEMPLATES   = "標籤模板"
TEMPLATES_HEADERS = ["廠商名稱", "模板名稱", "設定JSON", "最後更新", "Excel檔案ID"]
SHEET_VENDORS     = "廠商帳號"
VENDORS_HEADERS   = ["公司名稱", "網址", "帳號", "密碼"]

_DRIVE_FOLDER     = "ERP標籤模板"
_EXCEL_MIME       = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _retry(fn, retries: int = 4):
    """Retry fn on 429 with exponential back-off (1 / 2 / 4 / 8 s)."""
    for attempt in range(retries):
        try:
            return fn()
        except Exception as e:
            if "429" in str(e) and attempt < retries - 1:
                time.sleep(2 ** attempt)
            else:
                raise


@st.cache_resource
def get_client():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)


@st.cache_resource
def _get_drive_session():
    """AuthorizedSession for raw Drive REST calls."""
    from google.auth.transport.requests import AuthorizedSession  # type: ignore
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return AuthorizedSession(creds)


@st.cache_resource
def get_sheet(sheet_name: str):
    """Cache the worksheet object — avoids repeated open_by_key + worksheet() calls."""
    client = get_client()
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    try:
        ws = spreadsheet.worksheet(sheet_name)
        # 確保 SHEET_TEMPLATES 有 Excel檔案ID 欄位標題
        if sheet_name == SHEET_TEMPLATES:
            try:
                first_row = ws.row_values(1)
                if "Excel檔案ID" not in first_row:
                    ws.update_cell(1, len(first_row) + 1, "Excel檔案ID")
            except Exception:
                pass
        return ws
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=20)
        if sheet_name == SHEET_TEMPLATES:
            ws.append_row(TEMPLATES_HEADERS)
        elif sheet_name == SHEET_VENDORS:
            ws.append_row(VENDORS_HEADERS)
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


# ── Google Drive 模板 Excel 存取 ──────────────────────────────────

@st.cache_data(ttl=3600)
def _drive_folder_id() -> str:
    """取得（或建立）ERP標籤模板 Drive 資料夾，回傳 folder ID。"""
    session = _get_drive_session()
    q = f"name='{_DRIVE_FOLDER}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    r = session.get("https://www.googleapis.com/drive/v3/files",
                    params={"q": q, "fields": "files(id)"})
    files = r.json().get("files", [])
    if files:
        return files[0]["id"]
    r = session.post(
        "https://www.googleapis.com/drive/v3/files",
        json={"name": _DRIVE_FOLDER, "mimeType": "application/vnd.google-apps.folder"},
        params={"fields": "id"},
    )
    r.raise_for_status()
    return r.json()["id"]


def upload_template_excel(excel_bytes: bytes, customer: str, template_name: str) -> str:
    """上傳模板 Excel 到 Google Drive，回傳 file ID。"""
    import json as _json
    session = _get_drive_session()
    folder_id = _drive_folder_id()
    filename = f"{customer}_{template_name}.xlsx"

    # 檢查是否已有同名檔案
    q = f"name='{filename}' and '{folder_id}' in parents and trashed=false"
    r = session.get("https://www.googleapis.com/drive/v3/files",
                    params={"q": q, "fields": "files(id)"})
    existing = r.json().get("files", [])

    if existing:
        # 更新既有檔案內容
        file_id = existing[0]["id"]
        session.patch(
            f"https://www.googleapis.com/upload/drive/v3/files/{file_id}",
            params={"uploadType": "media"},
            headers={"Content-Type": _EXCEL_MIME},
            data=excel_bytes,
        )
        return file_id

    # 新建：multipart upload
    boundary = "===XLBOUNDARY==="
    metadata = _json.dumps({"name": filename, "parents": [folder_id]})
    body = (
        f"--{boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n"
        + metadata
        + f"\r\n--{boundary}\r\nContent-Type: {_EXCEL_MIME}\r\n\r\n"
    ).encode("utf-8") + excel_bytes + f"\r\n--{boundary}--".encode("utf-8")

    r = session.post(
        "https://www.googleapis.com/upload/drive/v3/files",
        params={"uploadType": "multipart", "fields": "id"},
        headers={"Content-Type": f"multipart/related; boundary={boundary}"},
        data=body,
    )
    r.raise_for_status()
    return r.json().get("id", "")


@st.cache_data(ttl=3600)
def download_template_excel(file_id: str) -> bytes:
    """從 Google Drive 下載模板 Excel bytes（每個 session 快取 1 小時）。"""
    session = _get_drive_session()
    r = session.get(
        f"https://www.googleapis.com/drive/v3/files/{file_id}",
        params={"alt": "media"},
    )
    r.raise_for_status()
    return r.content


# ── 標籤模板 ──────────────────────────────────────────────────────

@st.cache_data(ttl=120)
def load_templates(customer: str = "") -> list[dict]:
    ws = get_sheet(SHEET_TEMPLATES)
    records = _rows_to_records(_retry(ws.get_all_values), TEMPLATES_HEADERS)
    if customer:
        return [r for r in records if r.get("廠商名稱") == customer]
    return records


def save_template(customer: str, template_name: str, config_json: str,
                  excel_bytes: bytes = None):
    """儲存模板設定到 Sheets；若有 excel_bytes 則同步上傳到 Drive。"""
    ws = get_sheet(SHEET_TEMPLATES)
    records = _rows_to_records(_retry(ws.get_all_values), TEMPLATES_HEADERS)
    now = datetime.now(_TW).strftime("%Y/%m/%d %H:%M")

    # 上傳 Excel 到 Drive（失敗不中斷存檔流程）
    file_id = ""
    if excel_bytes:
        try:
            file_id = upload_template_excel(excel_bytes, customer, template_name)
        except Exception:
            pass

    for i, r in enumerate(records):
        if r.get("廠商名稱") == customer and r.get("模板名稱") == template_name:
            keep_fid = file_id or r.get("Excel檔案ID", "")
            _retry(lambda: ws.update(f"C{i+2}:E{i+2}", [[config_json, now, keep_fid]]))
            return
    _retry(lambda: ws.append_row([customer, template_name, config_json, now, file_id]))


def delete_template(customer: str, template_name: str):
    ws = get_sheet(SHEET_TEMPLATES)
    records = _rows_to_records(_retry(ws.get_all_values), TEMPLATES_HEADERS)
    for i, r in enumerate(records):
        if r.get("廠商名稱") == customer and r.get("模板名稱") == template_name:
            _retry(lambda: ws.delete_rows(i + 2))
            return


def clear_cache():
    load_templates.clear()
    try:
        load_vendors.clear()
    except Exception:
        pass
    try:
        download_template_excel.clear()
        _drive_folder_id.clear()
    except Exception:
        pass


# ── 廠商帳號 ──────────────────────────────────────────────────────

@st.cache_data(ttl=120)
def load_vendors() -> list[dict]:
    ws = get_sheet(SHEET_VENDORS)
    return _rows_to_records(_retry(ws.get_all_values), VENDORS_HEADERS)


def save_vendor(name: str, url: str, username: str, password: str):
    ws = get_sheet(SHEET_VENDORS)
    records = _rows_to_records(_retry(ws.get_all_values), VENDORS_HEADERS)
    for i, r in enumerate(records):
        if r.get("公司名稱") == name:
            _retry(lambda: ws.update(f"B{i+2}:D{i+2}", [[url, username, password]]))
            return
    _retry(lambda: ws.append_row([name, url, username, password]))


def delete_vendor(name: str):
    ws = get_sheet(SHEET_VENDORS)
    records = _rows_to_records(_retry(ws.get_all_values), VENDORS_HEADERS)
    for i, r in enumerate(records):
        if r.get("公司名稱") == name:
            _retry(lambda: ws.delete_rows(i + 2))
            return

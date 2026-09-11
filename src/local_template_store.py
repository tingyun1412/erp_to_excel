"""
標籤模板 + LSCR 基礎模板：本機檔案儲存。
取代原本的 Google Sheets（模板清單）+ Google Drive（模板 Excel 檔案），
不再需要任何 Google 憑證。

存放位置：跟這支程式同一層目錄下的 local_data/ 資料夾。整個 App 資料夾
（程式碼＋資料）搬到公司共用碟使用時，資料會跟著一起搬過去，
不需要另外設定路徑。

沿用原本 Google 版本的函式簽章（包含把 Drive 檔案 ID 欄位名稱
「Excel檔案ID」留著、值改存本機檔名），讓呼叫端（app.py）幾乎不用改。

本機檔案讀寫很快，不像原本呼叫 Google API 那樣需要靠快取避免延遲，
所以這裡不用 st.cache_data——每次都直接讀最新內容，也不需要額外呼叫任何
clear_cache 之類的函式來刷新。
"""
import json
import re
from datetime import datetime, timezone, timedelta
from pathlib import Path

_TW = timezone(timedelta(hours=8))

_BASE_DIR = Path(__file__).resolve().parent / "local_data"
_TEMPLATES_DIR = _BASE_DIR / "templates"
_TEMPLATES_META = _TEMPLATES_DIR / "templates.json"
_TEMPLATES_FILES_DIR = _TEMPLATES_DIR / "files"
_LSCR_BASE_FILE = _BASE_DIR / "LSCR_模板0.xlsx"

TEMPLATES_HEADERS = ["廠商名稱", "模板名稱", "設定JSON", "最後更新", "Excel檔案ID"]


def _safe_filename(customer: str, template_name: str) -> str:
    raw = f"{customer}_{template_name}".strip() or "模板"
    return re.sub(r'[\\/:*?"<>|]', "_", raw) + ".xlsx"


def _load_meta() -> list[dict]:
    if not _TEMPLATES_META.exists():
        return []
    try:
        return json.loads(_TEMPLATES_META.read_text(encoding="utf-8"))
    except Exception:
        return []


def _save_meta(records: list[dict]):
    _TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
    _TEMPLATES_META.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")


def load_templates(customer: str = "") -> list[dict]:
    records = _load_meta()
    if customer:
        return [r for r in records if r.get("廠商名稱") == customer]
    return records


def save_template(customer: str, template_name: str, config_json: str,
                   excel_bytes: bytes = None,
                   original_customer: str = None,
                   original_template_name: str = None) -> str | None:
    """
    儲存模板設定到本機清單；若有 excel_bytes 則同步寫入本機檔案。
    參數與回傳值語意跟原本 sheets_db.save_template 相同（見該處註解），
    只是「Drive 上傳失敗」換成了「本機寫檔失敗」。
    """
    records = _load_meta()
    now = datetime.now(_TW).strftime("%Y/%m/%d %H:%M")

    match_customer = original_customer if original_customer is not None else customer
    match_name     = original_template_name if original_template_name is not None else template_name

    write_error: str | None = None
    new_filename = ""
    if excel_bytes:
        try:
            _TEMPLATES_FILES_DIR.mkdir(parents=True, exist_ok=True)
            new_filename = _safe_filename(customer, template_name)
            (_TEMPLATES_FILES_DIR / new_filename).write_bytes(excel_bytes)
        except Exception as e:
            write_error = str(e)

    for r in records:
        if r.get("廠商名稱") == match_customer and r.get("模板名稱") == match_name:
            keep_filename = new_filename or r.get("Excel檔案ID", "")
            # 改名但這次沒有重新上傳 Excel：把舊檔案實體改名，避免清單名字換了、
            # 檔案卻還留著舊名字，兩邊對不起來。
            if not new_filename and keep_filename and (customer != match_customer or template_name != match_name):
                renamed = _safe_filename(customer, template_name)
                old_path = _TEMPLATES_FILES_DIR / keep_filename
                if old_path.exists():
                    try:
                        old_path.rename(_TEMPLATES_FILES_DIR / renamed)
                        keep_filename = renamed
                    except Exception:
                        pass
            r["廠商名稱"] = customer
            r["模板名稱"] = template_name
            r["設定JSON"] = config_json
            r["最後更新"] = now
            r["Excel檔案ID"] = keep_filename
            _save_meta(records)
            return write_error

    records.append({
        "廠商名稱": customer, "模板名稱": template_name, "設定JSON": config_json,
        "最後更新": now, "Excel檔案ID": new_filename,
    })
    _save_meta(records)
    return write_error


def delete_template(customer: str, template_name: str):
    """刪除模板設定，若有對應本機檔案也一併刪除，不留孤兒檔案。"""
    records = _load_meta()
    for i, r in enumerate(records):
        if r.get("廠商名稱") == customer and r.get("模板名稱") == template_name:
            filename = r.get("Excel檔案ID", "")
            if filename:
                try:
                    (_TEMPLATES_FILES_DIR / filename).unlink(missing_ok=True)
                except Exception:
                    pass
            del records[i]
            _save_meta(records)
            return


def download_template_excel(filename: str) -> bytes:
    """讀本機模板 Excel bytes（沿用舊名稱，參數其實是本機檔名不是 Drive file ID）。"""
    if not filename:
        return b""
    path = _TEMPLATES_FILES_DIR / filename
    return path.read_bytes() if path.exists() else b""


# ── LSCR 基礎模板（模板0）─────────────────────────────────────────


def find_lscr_base_template_id() -> str:
    """回傳非空字串代表本機已有 LSCR 模板0（沿用舊介面命名，值本身沒有實際意義）。"""
    return "local" if _LSCR_BASE_FILE.exists() else ""


def save_lscr_base_template(excel_bytes: bytes) -> str:
    _BASE_DIR.mkdir(parents=True, exist_ok=True)
    _LSCR_BASE_FILE.write_bytes(excel_bytes)
    return "local"


def download_lscr_base_template() -> bytes:
    return _LSCR_BASE_FILE.read_bytes() if _LSCR_BASE_FILE.exists() else b""

"""
電子發票逐張送出
把銷貨單資料逐張餵進 e-invoice.com.tw「一般會員發票開立」表單並送出。

登入頁有圖形驗證碼，無法全自動：先用 launch_login_browser() 開一個有畫面的
瀏覽器讓人工完成登入，再用 submit_batch() 連回同一個瀏覽器逐張送單。

只支援「一般會員」流程。核心會員（日月光等）表單結構尚未確認，
resolve_member_type() 會把這些客戶標成 "core"，呼叫端應排除、不要送進 submit_batch。
"""
import socket
import subprocess
import sys
from pathlib import Path

LOGIN_URL = "https://www.e-invoice.com.tw/j2iv/mgt/mgt_logon.jsp"
GENERAL_ENTRY_SELECTOR = 'a[href="/j2iv/FJ2IVC1OIOCOMMON01.do"]'
ISSUE_BUTTON_SELECTOR  = 'input[type="button"][value="開立發票"]'
CANCEL_BUTTON_SELECTOR = 'input[type="button"][value="放棄開立"]'
ADD_ITEM_SELECTOR      = "#idAddButton"

# 目前已知的核心會員（日月光等 8 家），表單流程尚未實作，僅用來排除、不進 submit_batch
CORE_MEMBER_CUSTOMERS = [
    "台塑", "南亞", "台化", "台塑石化",  # 台塑集團
    "景碩", "富采", "日月光", "欣興", "南茂", "中華精測", "台亞半導體",
]

_PROFILE_DIR = Path.home() / ".cache" / "einvoice_chrome_profile"


def _ensure_chromium_installed():
    """惰性安裝 Chromium：只有真的要用電子發票自動化時才檢查/安裝。"""
    chrome_dir = Path.home() / ".cache/ms-playwright"
    if not chrome_dir.exists():
        subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            check=True,
        )


def _free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def resolve_member_type(customer_name: str) -> str:
    """回傳 "core"（核心會員，尚未支援自動送出）或 "general"（一般會員）。"""
    name = (customer_name or "").strip()
    if not name:
        return "general"
    for core_name in CORE_MEMBER_CUSTOMERS:
        if core_name in name:
            return "core"
    return "general"


def launch_login_browser() -> dict:
    """
    開一個有畫面的 Chromium，導到登入頁，讓使用者手動輸入帳密與圖形驗證碼。
    user-data-dir 用固定目錄（不是臨時目錄），下次開啟有機會還保留登入 session，
    減少要重新過驗證碼的次數。
    回傳 {"port": int, "pid": int}。
    """
    _ensure_chromium_installed()
    from playwright.sync_api import sync_playwright

    port = _free_port()
    _PROFILE_DIR.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as pw:
        exe = pw.chromium.executable_path

    proc = subprocess.Popen(
        [
            exe,
            f"--remote-debugging-port={port}",
            f"--user-data-dir={_PROFILE_DIR}",
            "--no-first-run",
            "--no-default-browser-check",
            LOGIN_URL,
        ],
        creationflags=subprocess.DETACHED_PROCESS if sys.platform == "win32" else 0,
    )
    return {"port": port, "pid": proc.pid}


def is_login_alive(port: int) -> bool:
    """試連 CDP，確認瀏覽器還在且已登入（不在登入頁）。"""
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as pw:
            browser = pw.chromium.connect_over_cdp(f"http://localhost:{port}", timeout=5_000)
            ctx = browser.contexts[0]
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            still_login_page = page.locator('input[type="password"]').count() > 0
            return not still_login_page
    except Exception:
        return False


def close_login_browser(pid: int):
    """盡力關閉登入用的瀏覽器 process，失敗就吞掉。"""
    try:
        if sys.platform == "win32":
            subprocess.run(["taskkill", "/PID", str(pid), "/F", "/T"], capture_output=True)
        else:
            import os
            import signal
            os.kill(pid, signal.SIGTERM)
    except Exception:
        pass


# ── 表單填寫 ────────────────────────────────────────────────────────

def _fill_item_row(page, index: int, item: dict, order_no: str):
    """填第 index 項（0 起算）的品項欄位。index > 0 前要先按「新增項次」。"""
    if index > 0:
        page.click(ADD_ITEM_SELECTOR)
        page.wait_for_timeout(300)

    name = (item.get("name") or "").strip()
    spec = (item.get("description") or "").strip()
    unit = (item.get("unit") or "PCS").strip()
    qty = item.get("quantity", 0) or 0
    price = item.get("unit_price", 0) or 0
    remark = (item.get("remark") or "").strip()

    page.fill(f'input[name="dc_mtnm_1_{index}"]', name or item.get("item_no", ""))
    if spec:
        page.fill(f'input[name="dc_dsr2_1_{index}"]', spec)
    page.fill(f'input[name="dc_un1_1_{index}"]', unit)
    page.fill(f'input[name="dc_up_1_{index}"]', str(price))
    page.fill(f'input[name="dc_qty1_1_{index}"]', str(qty))
    page.fill(f'input[name="dc_relno1_1_{index}"]', order_no)
    if remark:
        page.fill(f'input[name="dc_relno2_1_{index}"]', remark)


def submit_one_order(page, order: dict, dry_run: bool = True) -> dict:
    """
    送出單一銷貨單的一般會員發票。
    回傳 {"success": bool, "invoice_no": str | None, "error": str | None, "name_mismatch": bool}
    dry_run=True 時，表單填完後按「放棄開立」而不是真的送出。
    """
    order_no = order.get("order_no", "")
    buyer_tax_id = (order.get("buyer_tax_id") or "").strip()
    result = {"success": False, "invoice_no": None, "error": None, "name_mismatch": False}

    if not buyer_tax_id:
        result["error"] = "銷貨單沒有受票人統一編號"
        return result

    items = order.get("items", [])
    if not items:
        result["error"] = "銷貨單沒有品項"
        return result

    try:
        page.click(GENERAL_ENTRY_SELECTOR)
        page.wait_for_load_state("networkidle", timeout=15_000)

        page.locator('input[type="text"]').first.fill(buyer_tax_id)
        page.click("text=下一步")
        page.wait_for_load_state("networkidle", timeout=15_000)

        # 粗略比對受票人名稱，不符不擋，只記錄供人工審核
        page_text = page.locator("body").inner_text()
        customer_name = (order.get("customer_name") or "").strip()
        if customer_name and customer_name not in page_text:
            result["name_mismatch"] = True

        for idx, item in enumerate(items):
            _fill_item_row(page, idx, item, order_no)

        page.fill("#dc_relno_0", order_no)

        if dry_run:
            page.click(CANCEL_BUTTON_SELECTOR)
            page.wait_for_timeout(500)
            result["error"] = "dry_run"
            return result

        page.click(ISSUE_BUTTON_SELECTOR)
        page.wait_for_load_state("networkidle", timeout=20_000)
        result["success"] = True
        return result

    except Exception as e:
        result["error"] = str(e)
        return result


def lookup_invoice_no_by_relno(page, order_no: str) -> str | None:
    """
    到查詢作業畫面，用相關號碼＝order_no 查回實際發票號碼。
    注意：查詢作業畫面的實際欄位/選擇器尚未經過現場驗證（探索時 session 逾時中斷），
    這裡先用最合理的猜測寫法，第一次真的送出發票後務必實際測試這個函式是否查得到。
    查不到就回傳 None，呼叫端要有 None 的容錯處理，不能假設一定查得到。
    """
    try:
        page.click("text=查詢作業")
        page.wait_for_load_state("networkidle", timeout=15_000)

        relno_input = page.locator(
            'input[name*="relno" i], input[id*="relno" i]'
        ).first
        if relno_input.count() == 0:
            return None
        relno_input.fill(order_no)

        search_btn = page.locator(
            'input[type="submit"], input[type="button"][value*="查詢"], button:has-text("查詢")'
        ).first
        if search_btn.count() == 0:
            return None
        search_btn.click()
        page.wait_for_load_state("networkidle", timeout=15_000)

        # 發票號碼通常是 2 碼英文 + 8 碼數字
        import re
        text = page.locator("body").inner_text()
        m = re.search(r"\b[A-Z]{2}\d{8}\b", text)
        return m.group(0) if m else None
    except Exception:
        return None


def submit_batch(port: int, orders: list[dict], dry_run: bool = True, on_progress=None) -> list[dict]:
    """
    連回已登入的瀏覽器，逐張送出 orders（皆須為 resolve_member_type=="general"）。
    單張失敗不中斷後面的（比照 erp_downloader.download_label_pdfs 的作法）。
    on_progress(index, total, order, result) 會在每張處理完後呼叫，方便呼叫端即時更新畫面/寫入紀錄。
    回傳每張的結果 list（跟 orders 同順序），每個 dict 額外帶 "order_no"/"customer_name"。
    """
    from playwright.sync_api import sync_playwright

    results = []
    total = len(orders)

    with sync_playwright() as pw:
        browser = pw.chromium.connect_over_cdp(f"http://localhost:{port}")
        ctx = browser.contexts[0]
        page = ctx.pages[0] if ctx.pages else ctx.new_page()

        for i, order in enumerate(orders):
            try:
                result = submit_one_order(page, order, dry_run=dry_run)
                if result["success"] and not result.get("invoice_no"):
                    result["invoice_no"] = lookup_invoice_no_by_relno(page, order.get("order_no", ""))
            except Exception as e:
                result = {"success": False, "invoice_no": None, "error": str(e), "name_mismatch": False}

            result["order_no"] = order.get("order_no", "")
            result["customer_name"] = order.get("customer_name", "")
            results.append(result)

            if on_progress:
                on_progress(i, total, order, result)

        # 注意：這裡刻意不呼叫 browser.close()——瀏覽器是外部登入用的 subprocess，
        # 只是連線斷開（跳出 with 區塊時 playwright driver 會自動斷線），
        # 瀏覽器本身留著給下一批送單用，避免每批都要重新登入過驗證碼。

    return results

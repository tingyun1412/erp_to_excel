"""
電子發票逐張送出
把銷貨單資料逐張餵進 e-invoice.com.tw「一般會員發票開立」表單。

登入頁有圖形驗證碼，無法全自動：先用 launch_login_browser() 開一個有畫面的
瀏覽器讓人工完成登入（user-data-dir 用固定目錄，之後一段時間內重開還留著登入
session，不用每次都重新過驗證碼），之後用 fill_one_order() 連回同一個瀏覽器
逐張把表單填好。

重要：真正送出（dry_run=False）時，填完表單後絕對不會自動按「開立發票」——
這一步一定要由人工在瀏覽器視窗裡親自確認、親自按下去，程式只負責填欄位。
只有測試用的 dry_run=True 模式會自動按「放棄開立」收尾，不會留下任何紀錄。

只支援「一般會員」流程。核心會員（目前只有日月光）表單結構尚未確認，
resolve_member_type() 會把這些客戶標成 "core"，呼叫端應排除、不要送進 fill_one_order。
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

# e-invoice.com.tw 帳密（跟公司內部系統帳密不同層級，直接寫死不用另外設定）
LOGIN_ACCOUNT  = "24405403I1"
LOGIN_PASSWORD = "Gc24405403"

# 核心會員只有日月光，其餘一律走一般會員流程
CORE_MEMBER_CUSTOMERS = ["日月光"]

_PROFILE_DIR = Path.home() / ".cache" / "einvoice_chrome_profile"


def _ensure_chromium_installed():
    """惰性安裝 Chromium：只有真的要用電子發票自動化時才檢查/安裝。
    直接問 Playwright 實際會去哪裡找執行檔（會尊重 PLAYWRIGHT_BROWSERS_PATH），
    而不是猜一個固定路徑——固定路徑在 Windows 上跟預設快取位置對不起來，
    會導致明明已經裝好還是每次都想重新下載。"""
    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        exe = Path(pw.chromium.executable_path)
    if not exe.exists():
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


def _active_page(ctx):
    """
    從 context 現有的分頁裡選出目前該操作的那一個：優先挑「不在登入頁」的
    分頁（網址不是 mgt_logon.jsp 且畫面上沒有密碼欄位）——e-invoice.com.tw
    登入成功後有時會另外開一個新分頁承接後續內容，舊的登入分頁不會自動關掉，
    所以不能直接用第一個分頁。較晚開的分頁通常才是登入後的正確分頁，所以
    從後往前找。都不符合就回傳 None（呼叫端自行決定要不要重試），絕對不會
    在這裡另外生一個新分頁出來，避免分頁越開越多。
    """
    pages = ctx.pages
    if not pages:
        return None
    for page in reversed(pages):
        try:
            on_login_url = "mgt_logon.jsp" in page.url
            has_password_field = page.locator('input[type="password"]').count() > 0
        except Exception:
            continue
        if not (on_login_url or has_password_field):
            return page
    return None


def launch_login_browser() -> dict:
    """
    開一個有畫面的 Chromium，導到登入頁，並自動把帳號密碼填好——
    只剩圖形驗證碼要人工輸入，輸完按登入即可。
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

    # 等瀏覽器真的開起來、把帳密先填好，讓使用者只需要輸入圖形驗證碼。
    # 如果已經是登入後的狀態（沿用了舊 session），畫面上不會有密碼欄位，填不到也沒關係。
    import time
    for _ in range(20):
        try:
            with sync_playwright() as pw:
                browser = pw.chromium.connect_over_cdp(f"http://localhost:{port}", timeout=2_000)
                ctx = browser.contexts[0]
                page = ctx.pages[0] if ctx.pages else None
                if page is not None and page.locator("#inputAcno").count() > 0:
                    page.fill("#inputAcno", LOGIN_ACCOUNT)
                    page.fill("#inputPswd", LOGIN_PASSWORD)
            if page is not None:
                break
        except Exception:
            pass
        time.sleep(0.5)

    return {"port": port, "pid": proc.pid}


def is_login_alive(port: int) -> bool:
    """試連 CDP，確認瀏覽器還在且已登入（不在登入頁）。"""
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as pw:
            browser = pw.chromium.connect_over_cdp(f"http://localhost:{port}", timeout=5_000)
            ctx = browser.contexts[0]
            return _active_page(ctx) is not None
    except Exception:
        return False


def close_login_browser(pid: int):
    """盡力關閉登入用的瀏覽器 process，失敗就吞掉。"""
    try:
        if sys.platform == "win32":
            # 這支 App 是用 pythonw.exe 跑（沒有主控台），taskkill.exe 預設還是會
            # 自己彈一個主控台視窗閃一下；加 CREATE_NO_WINDOW 讓它安靜執行。
            subprocess.run(
                ["taskkill", "/PID", str(pid), "/F", "/T"],
                capture_output=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
        else:
            import os
            import signal
            os.kill(pid, signal.SIGTERM)
    except Exception:
        pass


# ── 表單填寫 ────────────────────────────────────────────────────────

def _fill_item_row(page, index: int, item: dict, customer_order_no: str):
    """
    填第 index 項（0 起算）的品項欄位。index > 0 前要先按「新增項次」。
    品名1＝品名+規格合併，品名2＝客戶料號（備註）；
    相關號碼一＝客戶訂單號碼，相關號碼二不填。
    """
    if index > 0:
        page.click(ADD_ITEM_SELECTOR)
        page.wait_for_timeout(300)

    name = (item.get("name") or "").strip()
    spec = (item.get("description") or "").strip()
    name_and_spec = (name + spec).strip() or item.get("item_no", "")
    unit = (item.get("unit") or "PCS").strip()
    qty = item.get("quantity", 0) or 0
    price = item.get("unit_price", 0) or 0
    customer_part_no = (item.get("remark") or "").strip()  # 客戶料號
    # 品項自己有指定的話優先用品項的（例如月結彙總發票，每個品項來自不同出貨單號）
    item_customer_order_no = (item.get("customer_order_no") or customer_order_no or "").strip()

    page.fill(f'input[name="dc_mtnm_1_{index}"]', name_and_spec)
    if customer_part_no:
        page.fill(f'input[name="dc_dsr2_1_{index}"]', customer_part_no)
    page.fill(f'input[name="dc_un1_1_{index}"]', unit)
    page.fill(f'input[name="dc_up_1_{index}"]', str(price))
    page.fill(f'input[name="dc_qty1_1_{index}"]', str(qty))
    if item_customer_order_no:
        page.fill(f'input[name="dc_relno1_1_{index}"]', item_customer_order_no)


def fill_one_order(page, order: dict, dry_run: bool = True) -> dict:
    """
    把單一銷貨單的資料填進「一般會員發票開立」表單。
    回傳 {"state": ..., "error": str | None, "name_mismatch": bool}
    state：
      "filled_awaiting_manual_submit" —— 已填完，表單留在畫面上等人工在瀏覽器裡親自按「開立發票」
      "dry_run_cancelled"             —— 測試模式，已自動按「放棄開立」收尾，沒有留下任何紀錄
      "error"                         —— 失敗，看 error 欄位
    這個函式本身絕對不會按「開立發票」——dry_run=False 時填完就停在畫面上，
    要不要送出、什麼時候送出，都是人工在瀏覽器裡自己決定。
    """
    order_no = order.get("order_no", "")
    customer_order_no = (order.get("customer_order_no") or order_no or "").strip()
    buyer_tax_id = (order.get("buyer_tax_id") or "").strip()
    result = {"state": "error", "error": None, "name_mismatch": False}

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
            _fill_item_row(page, idx, item, customer_order_no)

        # 整張發票層級的相關號碼：填我們自己的銷貨單號，之後用它反查真實發票號碼
        page.fill("#dc_relno_0", order_no)

        if dry_run:
            page.click(CANCEL_BUTTON_SELECTOR)
            page.wait_for_timeout(500)
            result["state"] = "dry_run_cancelled"
            return result

        result["state"] = "filled_awaiting_manual_submit"
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


def _with_page(port: int, fn):
    """連回已登入的瀏覽器、取得 page、執行 fn(page)、斷線（不 close，瀏覽器留給下次用）。"""
    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        browser = pw.chromium.connect_over_cdp(f"http://localhost:{port}")
        ctx = browser.contexts[0]
        page = _active_page(ctx)
        if page is None:
            raise RuntimeError("找不到已登入的分頁，請重新登入")
        return fn(page)
    # 注意：這裡刻意不呼叫 browser.close()——瀏覽器是外部登入用的 subprocess，
    # 只是連線斷開，瀏覽器本身留著給下一次操作用，避免每次都要重新登入過驗證碼。


def fill_one_order_via_port(port: int, order: dict, dry_run: bool = True) -> dict:
    """連線→填一筆→斷線，包成一次呼叫，方便 Streamlit 每次按鈕互動各自獨立呼叫。"""
    result = _with_page(port, lambda page: fill_one_order(page, order, dry_run=dry_run))
    result["order_no"] = order.get("order_no", "")
    result["customer_name"] = order.get("customer_name", "")
    return result


def lookup_invoice_no_via_port(port: int, order_no: str) -> str | None:
    return _with_page(port, lambda page: lookup_invoice_no_by_relno(page, order_no))


def dry_run_batch(port: int, orders: list[dict], on_progress=None) -> list[dict]:
    """
    測試模式專用：連回瀏覽器，逐張自動填寫＋自動按「放棄開立」，不會留下任何真實紀錄。
    只用來驗證欄位填寫邏輯是否正確；真實送出一律要逐張人工確認，見 fill_one_order_via_port。
    單張失敗不中斷後面的（比照 erp_downloader.download_label_pdfs 的作法）。
    """
    from playwright.sync_api import sync_playwright

    results = []
    total = len(orders)

    with sync_playwright() as pw:
        browser = pw.chromium.connect_over_cdp(f"http://localhost:{port}")
        ctx = browser.contexts[0]
        page = _active_page(ctx)
        if page is None:
            raise RuntimeError("找不到已登入的分頁，請重新登入")

        for i, order in enumerate(orders):
            try:
                result = fill_one_order(page, order, dry_run=True)
            except Exception as e:
                result = {"state": "error", "error": str(e), "name_mismatch": False}

            result["order_no"] = order.get("order_no", "")
            result["customer_name"] = order.get("customer_name", "")
            results.append(result)

            if on_progress:
                on_progress(i, total, order, result)

    return results

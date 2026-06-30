"""
ERP 標籤自動下載
登入 http://scm.honprec.com/hp/Index.aspx，依出貨單號下載標籤 PDF。

需求：
    pip install playwright
    python -m playwright install chromium
"""
import io
import zipfile
from pathlib import Path

ERP_INDEX = "http://scm.honprec.com/hp/Index.aspx"


def download_label_pdfs(
    order_nos: list,
    username: str = "BR026",
    password: str = "5403",
    debug_dir: str = None,
) -> dict:
    """
    登入 ERP，依出貨單號下載標籤 PDF。
    回傳 {order_no: pdf_bytes | None}
    debug_dir 非 None 時每步截圖到該目錄（方便除錯）。
    """
    from playwright.sync_api import sync_playwright

    results = {no: None for no in order_nos}

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        ctx = browser.new_context(accept_downloads=True)
        page = ctx.new_page()
        dbg = _make_debugger(page, debug_dir)

        try:
            _login(page, username, password, dbg)
        except Exception as e:
            dbg("error_login")
            browser.close()
            raise RuntimeError(f"ERP 登入失敗：{e}") from e

        for order_no in order_nos:
            try:
                _go_to_ship_orders(page, dbg)
                pdf = _download_one(page, order_no, dbg)
                results[order_no] = pdf
            except Exception as e:
                dbg(f"error_{order_no}")
                print(f"[ERP] {order_no} 下載失敗：{e}")

        browser.close()

    return results


def pack_zip(results: dict) -> bytes:
    """將 {order_no: pdf_bytes} 打包成 ZIP bytes。"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for order_no, pdf_bytes in results.items():
            if pdf_bytes:
                zf.writestr(f"標籤_{order_no}.pdf", pdf_bytes)
    buf.seek(0)
    return buf.read()


# ── 內部函式 ──────────────────────────────────────────────────────

def _make_debugger(page, debug_dir):
    counter = [0]
    if not debug_dir:
        return lambda name: None
    Path(debug_dir).mkdir(parents=True, exist_ok=True)
    def snap(name):
        counter[0] += 1
        try:
            page.screenshot(path=str(Path(debug_dir) / f"{counter[0]:02d}_{name}.png"))
        except Exception:
            pass
    return snap


def _login(page, username, password, dbg):
    page.goto(ERP_INDEX, timeout=30_000)
    page.wait_for_load_state("domcontentloaded")
    dbg("01_login")

    # 填帳號 — ASP.NET WebForms 常見 ID 模式
    _fill_first(page, username, [
        "input[name*='UserID']", "input[id*='UserID']",
        "input[name*='User']",  "input[id*='User']",
        "input[type='text']:visible",
    ])

    # 填密碼
    page.locator("input[type='password']:visible").first.fill(password)

    # 送出
    _click_first(page, [
        "input[type='submit']:visible",
        "button[type='submit']:visible",
        "input[value*='登入']:visible",
        "input[value*='確定']:visible",
        "a:has-text('登入'):visible",
    ])

    page.wait_for_load_state("networkidle", timeout=20_000)
    dbg("02_after_login")


def _go_to_ship_orders(page, dbg):
    page.locator("text=出貨單").first.click()
    page.wait_for_load_state("networkidle", timeout=20_000)
    dbg("03_order_list")


def _download_one(page, order_no: str, dbg) -> bytes:
    # 找含出貨單號的列
    row = page.locator(f"tr:has-text('{order_no}')").first
    if row.count() == 0:
        raise RuntimeError(f"找不到出貨單 {order_no}")

    # 點「標籤」按鈕
    row.locator("text=標籤").first.click()
    page.wait_for_load_state("networkidle", timeout=20_000)
    dbg(f"04_label_{order_no}")

    # 勾選全部 checkbox（如果有）
    for cb in page.locator("input[type='checkbox']:visible").all():
        try:
            if not cb.is_checked():
                cb.check()
        except Exception:
            pass

    # 點「下載選取標籤」
    with page.expect_download(timeout=30_000) as dl_info:
        page.locator("text=下載選取標籤").click()

    dl = dl_info.value
    pdf_bytes = Path(dl.path()).read_bytes()
    dbg(f"05_done_{order_no}")
    return pdf_bytes


def _fill_first(page, value, selectors):
    for sel in selectors:
        loc = page.locator(sel)
        if loc.count() > 0:
            loc.first.fill(value)
            return
    raise RuntimeError("找不到帳號輸入欄位")


def _click_first(page, selectors):
    for sel in selectors:
        loc = page.locator(sel)
        if loc.count() > 0:
            loc.first.click()
            return
    raise RuntimeError("找不到登入按鈕")

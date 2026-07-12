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

ERP_INDEX  = "http://scm.honprec.com/hp/Index.aspx"
ERP_ORDERS = "https://scm.honprec.com/HP/MA10.aspx"   # 出貨單列表（觀察自錯誤訊息 URL）


def download_label_pdfs(
    order_nos: list,
    username: str = "BR026",
    password: str = "5403",
    debug_dir: str = None,
) -> tuple[dict, dict]:
    """
    登入 ERP，依出貨單號下載標籤 PDF。
    回傳 (results, errors)
      results: {order_no: pdf_bytes | None}
      errors:  {order_no: error_str}  — 失敗原因
    """
    from playwright.sync_api import sync_playwright

    results = {no: None for no in order_nos}
    errors: dict = {}

    with sync_playwright() as pw:
        # Streamlit Cloud 容器沒有沙盒所需權限、/dev/shm 也很小，Chromium 預設啟動
        # 會在沙盒初始化或共用記憶體時直接 segfault（而非丟出乾淨的例外）。
        browser = pw.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu"],
        )
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
                _go_to_ship_orders(page, dbg, username, password)
                pdf = _download_one(page, order_no, dbg)
                results[order_no] = pdf
            except Exception as e:
                dbg(f"error_{order_no}")
                errors[order_no] = str(e)

        browser.close()

    return results, errors


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


def _is_login_page(page) -> bool:
    return page.locator("input[type='password']").count() > 0


def _click_nav_ship_orders(page) -> bool:
    """
    點「出貨單」導覽按鈕（ASP.NET PostBack: id=ctl00_btnMA06）。
    先用 ID 精確定位，找不到再用文字 fallback。
    點擊後等待 PostBack 完成（networkidle）。
    """
    for frame in page.frames:
        try:
            if (frame.url or "") in ("about:blank", ""):
                continue
            # 優先：已知 ASP.NET 控制項 ID
            loc = frame.locator("#ctl00_btnMA06")
            if loc.count() == 0:
                # Fallback：文字包含「出貨單」的 <a>
                loc = frame.locator("a:has-text('出貨單')")
            if loc.count() > 0:
                loc.first.click(timeout=5_000)
                # PostBack 是同步 form submit，等待頁面重新載入
                try:
                    page.wait_for_load_state("networkidle", timeout=20_000)
                except Exception:
                    pass
                return True
        except Exception:
            pass
    return False


def _go_to_ship_orders(page, dbg, username: str = "", password: str = ""):
    """確保 session 有效後點「出貨單」，PostBack 後等頁面穩定。"""
    # ── session 失效偵測 ──
    if _is_login_page(page):
        dbg("relogin_needed")
        _login(page, username, password, dbg)
        page.wait_for_load_state("networkidle", timeout=20_000)

    dbg("03a_dashboard")

    # ── 點「出貨單」──
    if not _click_nav_ship_orders(page):
        # 若目前頁面沒有導覽列（如 PDF 下載完成頁），先回首頁
        page.goto(ERP_INDEX, timeout=30_000)
        page.wait_for_load_state("networkidle", timeout=15_000)
        if _is_login_page(page):
            _login(page, username, password, dbg)
            page.wait_for_load_state("networkidle", timeout=20_000)
        if not _click_nav_ship_orders(page):
            raise RuntimeError("找不到「出貨單」按鈕（#ctl00_btnMA06 / a:has-text）")

    dbg("03_after_postback")


def _find_order_row(page, order_no: str):
    """
    在所有非 about:blank frame（含主 frame）裡找出貨單號對應的 <tr>。
    回傳 (frame, row_locator) 或 (page, None)。
    """
    for f in page.frames:
        try:
            url = f.url or ""
            if url == "about:blank":
                continue
            row = f.locator(f"tr:has-text('{order_no}')")
            if row.count() > 0:
                return f, row.first
        except Exception:
            pass
    return page, None


def _download_one(page, order_no: str, dbg) -> bytes:
    import time

    # ── 輪詢等待出貨單號出現（最多 25 秒）──
    row = None
    deadline = time.time() + 25
    while time.time() < deadline:
        _, row = _find_order_row(page, order_no)
        if row is not None:
            break
        page.wait_for_timeout(500)

    # ── 若輪詢後仍找不到，嘗試搜尋欄位 ──
    if row is None:
        for f in page.frames:
            try:
                si = f.locator("input[type='text']:visible")
                if si.count() > 0:
                    si.first.fill(order_no)
                    si.first.press("Enter")
                    page.wait_for_load_state("networkidle", timeout=10_000)
                    _, row = _find_order_row(page, order_no)
                    break
            except Exception:
                pass

    if row is None:
        url = page.url
        title = page.title()
        frames_info = " | ".join(f.url for f in page.frames)
        preview = ""
        for f in page.frames:
            try:
                t = f.inner_text("body")[:300].replace("\n", " ")
                if t.strip():
                    preview += f"[{f.url[-40:]}] {t}  "
            except Exception:
                pass
        raise RuntimeError(
            f"找不到出貨單 {order_no}。"
            f"當前頁：{title} | {url}\n"
            f"Frames: {frames_info}\n"
            f"頁面內容預覽：{preview[:500]}"
        )

    # ── 點「標籤」按鈕 ──
    row.locator("text=標籤").first.click()
    page.wait_for_load_state("networkidle", timeout=20_000)
    dbg(f"04_label_{order_no}")

    # ── 勾選全部 checkbox ──
    for f in page.frames:
        try:
            if (f.url or "") == "about:blank":
                continue
            for cb in f.locator("input[type='checkbox']:visible").all():
                if not cb.is_checked():
                    cb.check()
        except Exception:
            pass

    # ── 點「下載選取標籤」並接收下載 ──
    with page.expect_download(timeout=30_000) as dl_info:
        for f in page.frames:
            try:
                if (f.url or "") == "about:blank":
                    continue
                btn = f.locator("text=下載選取標籤")
                if btn.count() > 0:
                    btn.first.click()
                    break
            except Exception:
                pass

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

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


def _go_to_ship_orders(page, dbg, username: str = "", password: str = ""):
    import re as _re
    page.goto(ERP_INDEX, timeout=30_000)
    page.wait_for_load_state("domcontentloaded")

    # PDF 下載後 ERP session 可能失效，自動重新登入
    if _is_login_page(page):
        dbg("relogin_needed")
        if username and password:
            _login(page, username, password, dbg)
        else:
            raise RuntimeError("ERP session 已過期，無法自動重新登入（缺少帳密）")

    # 取出「出貨單」連結的 href，直接 goto（比 click 可靠）
    nav_url = None
    try:
        link = page.locator("a:has-text('出貨單')").first
        href = link.get_attribute("href", timeout=3_000)
        if href and not href.lower().startswith("javascript") and href != "#":
            base = _re.match(r'https?://[^/]+', page.url).group()
            if href.startswith("http"):
                nav_url = href
            elif href.startswith("/"):
                nav_url = base + href
            else:
                # 相對路徑 e.g. "MA10.aspx"
                folder = _re.match(r'(https?://[^/]+(?:/[^/]+/)?)', page.url).group(1)
                nav_url = folder + href
    except Exception:
        pass

    if nav_url:
        page.goto(nav_url, timeout=30_000)
    else:
        # Fallback：直接 click
        for sel in ["a:has-text('出貨單')", "button:has-text('出貨單')", "text=出貨單"]:
            loc = page.locator(sel)
            if loc.count() > 0:
                loc.first.click()
                break

    page.wait_for_load_state("networkidle", timeout=20_000)
    # 等待 table 出現（動態載入）
    try:
        page.wait_for_selector("table", timeout=10_000)
    except Exception:
        pass
    dbg("03_order_list")


def _active_frame(page, order_no: str = ""):
    """
    ERP 頁面常用 iframe 架構：主頁是外框，內容在子 frame。
    先找含出貨單號的 frame；若找不到則找最大的非主 frame；最後 fallback 主 page。
    """
    frames = page.frames
    if len(frames) <= 1:
        return page  # 沒有 iframe，直接用 main frame

    if order_no:
        for f in frames[1:]:
            try:
                if f.locator(f"text={order_no}").count() > 0:
                    return f
            except Exception:
                pass

    # fallback：第一個非主 frame（通常是內容區）
    return frames[1]


def _download_one(page, order_no: str, dbg) -> bytes:
    # 等 table 出現（動態載入頁可能需要額外時間）
    try:
        page.wait_for_selector("table tr, iframe", timeout=15_000)
    except Exception:
        pass

    frame = _active_frame(page, order_no)

    # 在正確 frame 裡找出貨單列
    row = frame.locator(f"tr:has-text('{order_no}')").first
    if row.count() == 0:
        # 嘗試搜尋欄位
        for s_sel in [
            f"input[type='text'][id*='search' i]:visible",
            "input[type='text']:visible",
        ]:
            search_inputs = frame.locator(s_sel)
            if search_inputs.count() > 0:
                search_inputs.first.fill(order_no)
                search_inputs.first.press("Enter")
                page.wait_for_load_state("networkidle", timeout=15_000)
                frame = _active_frame(page, order_no)
                row = frame.locator(f"tr:has-text('{order_no}')").first
                if row.count() > 0:
                    break

    if row.count() == 0:
        url = page.url
        title = page.title()
        frames_info = " | ".join(f.url for f in page.frames)
        # 嘗試從所有 frame 取得頁面文字以輔助除錯
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

    # 點「標籤」按鈕
    row.locator("text=標籤").first.click()
    page.wait_for_load_state("networkidle", timeout=20_000)
    dbg(f"04_label_{order_no}")

    # 勾選全部 checkbox（如果有）——同樣在正確 frame 裡找
    frame2 = _active_frame(page)
    for cb in frame2.locator("input[type='checkbox']:visible").all():
        try:
            if not cb.is_checked():
                cb.check()
        except Exception:
            pass

    # 點「下載選取標籤」
    with page.expect_download(timeout=30_000) as dl_info:
        frame2.locator("text=下載選取標籤").click()

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

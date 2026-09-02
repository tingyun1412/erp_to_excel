"""
均華(GMM)供應商作業平台 — 出貨標籤下載
登入 http://SLP.gmmcorp.com.tw:8080/apps/lg_001.nsf，依出貨單號抓取標籤資料，
直接截取均華網站原本的畫面（保留原本的顏色/框線/字型，不重畫），組成 PDF，
跟 erp_downloader.download_label_pdfs() 用同樣的 {order_no: pdf_bytes} 回傳
格式，接上 app.py 現有的「複製截圖」流程（PDF 每頁轉一張圖、垂直合併、
按鈕一鍵複製到剪貼簿貼進 BarTender 之類的標籤軟體）。

這個網站是 Lotus Domino 應用程式，用 HTTP Basic Auth 保護在資料庫層級
（不是表單登入）。找出貨單號對應文件的過程純用 requests（快、不用開瀏覽器）；
只有最後「擷取標籤畫面」這一步才用 Playwright 開真的瀏覽器，直接對均華
網站算圖後的結果截圖，畫面跟人工操作看到的一模一樣，不是重新畫的版本。

流程（對應人工操作的每一步）：
  1. 「出貨-依建立日」列表（FVEWFV02）— 依出貨單號找到文件 UID。
     出貨單號本身帶建立日期（如 S2260902108＝2026/09/02 第108筆），列表本身
     就是依建立日新到舊排序，正常使用（列印剛建立沒多久的出貨單）第一頁就
     找得到；找不到才跟著「下一頁」按鈕用的 Domino Click token 翻頁。
  2. 開啟該筆文件（EditDocument）— 取得 SeleNum（預設全選的標籤序號清單）。
  3. 呼叫「列印選取標籤(標籤機)」按鈕實際打的網址（RunAgent，
     Function=PrtSele2DBarcode4Supplier）— 用 Playwright 開啟這個網址，
     對每一張標籤（<table class='lable'>）個別截圖，裁切精準，畫面跟均華
     網站本身完全一致。
"""
import io
import re
import subprocess
import sys
import zipfile
from pathlib import Path

import requests

BASE = "http://SLP.gmmcorp.com.tw:8080/apps/lg_001.nsf"
VIEW_URL = f"{BASE}/FVEWFV02?OpenForm"


def _ensure_chromium_installed():
    """惰性安裝 Chromium：只有真的要用這個下載功能時才檢查/安裝。"""
    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        exe = Path(pw.chromium.executable_path)
    if not exe.exists():
        subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            check=True,
        )


def _get(session: requests.Session, url: str, **kw) -> requests.Response:
    """統一在這裡設定 timeout 跟強制 UTF-8（伺服器有標 charset=UTF-8，
    但 requests 有時猜錯成別的編碼，中文會變亂碼）。"""
    resp = session.get(url, timeout=kw.pop("timeout", 20), **kw)
    resp.raise_for_status()
    resp.encoding = "utf-8"
    return resp


def _post(session: requests.Session, url: str, data: dict, **kw) -> requests.Response:
    resp = session.post(url, data=data, timeout=kw.pop("timeout", 20), **kw)
    resp.raise_for_status()
    resp.encoding = "utf-8"
    return resp


def _find_doc_uid(session: requests.Session, order_no: str, max_pages: int = 15) -> str | None:
    """在「出貨-依建立日」列表裡找出貨單號對應的文件 UID（32 碼十六進位）。"""
    row_pat = re.compile(
        re.escape(order_no) + r".*?/0/([0-9A-Fa-f]{32})\?EditDocument",
        re.DOTALL,
    )
    next_pat = re.compile(r"下一頁[^>]*_doClick\('([^']+)'")

    click_token = None
    for _ in range(max_pages):
        resp = (_post(session, VIEW_URL, {"__Click": click_token})
                if click_token else _get(session, VIEW_URL))
        page_html = resp.text
        m = row_pat.search(page_html)
        if m:
            return m.group(1).upper()
        nm = next_pat.search(page_html)
        if not nm:
            return None
        click_token = nm.group(1)
    return None


def _get_sele_num(session: requests.Session, doc_uid: str) -> str:
    resp = _get(session, f"{BASE}/0/{doc_uid}?EditDocument")
    m = re.search(r'name="SeleNum"\s+value="([^"]*)"', resp.text)
    return m.group(1) if m else ""


def _screenshot_labels(username: str, password: str, doc_uid: str, sele_num: str) -> list[bytes]:
    """
    用 Playwright 開啟「列印選取標籤(標籤機)」實際打的網址，對每一張標籤
    （<table class='lable'>）個別截圖，回傳 PNG bytes 的清單，畫面跟均華
    網站本身完全一致（顏色/框線/字型都不變），不是重新畫的版本。
    """
    _ensure_chromium_installed()
    from playwright.sync_api import sync_playwright

    url = (f"{BASE}/RunAgent?OpenAgent&Function=PrtSele2DBarcode4Supplier"
           f"&ParentUNID={doc_uid}&SeleNum={sele_num}")

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        ctx = browser.new_context(http_credentials={"username": username, "password": password})
        page = ctx.new_page()
        page.goto(url, timeout=30_000, wait_until="networkidle")
        # 頁面 onload 會自動彈 window.print()（無頭模式下不會真的印，但穩妥起見擋掉）
        page.evaluate("window.print = function(){}")

        tables = page.locator("table.lable")
        count = tables.count()
        shots = [tables.nth(i).screenshot() for i in range(count)]

        browser.close()
    return shots


def download_label_pdfs(
    order_nos: list,
    username: str = "110488",
    password: str = "4667044110488",
) -> tuple[dict, dict]:
    """
    登入均華供應商平台，依出貨單號抓取標籤，直接截取原始畫面組成 PDF
    （每張標籤一頁）。回傳 (results, errors)：
      results: {order_no: pdf_bytes | None}
      errors:  {order_no: error_str}
    格式跟 erp_downloader.download_label_pdfs 一致，可以直接沿用 app.py
    既有的「複製截圖」／打包 zip 流程，呼叫端不用另外寫一套。
    """
    from PIL import Image

    results = {no: None for no in order_nos}
    errors: dict = {}

    session = requests.Session()
    session.auth = (username, password)

    for order_no in order_nos:
        try:
            doc_uid = _find_doc_uid(session, order_no)
            if not doc_uid:
                raise RuntimeError(
                    f"在「出貨-依建立日」列表裡找不到出貨單號 {order_no}"
                    "（可能太舊、翻頁翻不到，或單號打錯）"
                )
            sele_num = _get_sele_num(session, doc_uid)
            shots = _screenshot_labels(username, password, doc_uid, sele_num)
            if not shots:
                raise RuntimeError("查得到這張出貨單，但解析不到任何標籤內容")

            imgs = [Image.open(io.BytesIO(s)).convert("RGB") for s in shots]
            buf = io.BytesIO()
            imgs[0].save(buf, format="PDF", save_all=True, append_images=imgs[1:])
            results[order_no] = buf.getvalue()
        except requests.exceptions.HTTPError as e:
            if e.response is not None and e.response.status_code == 401:
                errors[order_no] = "帳號密碼錯誤（HTTP 401）"
            else:
                errors[order_no] = str(e)
        except Exception as e:
            errors[order_no] = str(e)

    return results, errors


def pack_zip(results: dict) -> bytes:
    """跟 erp_downloader.pack_zip 一樣：把 {order_no: pdf_bytes} 打包成 ZIP。"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for order_no, pdf_bytes in results.items():
            if pdf_bytes:
                zf.writestr(f"標籤_{order_no}.pdf", pdf_bytes)
    buf.seek(0)
    return buf.read()

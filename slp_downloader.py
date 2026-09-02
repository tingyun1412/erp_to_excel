"""
均華(GMM)供應商作業平台 — 出貨標籤下載
登入 http://SLP.gmmcorp.com.tw:8080/apps/lg_001.nsf，依出貨單號抓取標籤資料，
自己重新畫成標籤圖片（含 QR Code），組成 PDF，跟 erp_downloader.download_label_pdfs()
用同樣的 {order_no: pdf_bytes} 回傳格式，直接接上 app.py 現有的「複製截圖」流程
（PDF 每頁轉一張圖、垂直合併、按鈕一鍵複製到剪貼簿貼進 BarTender 之類的標籤軟體）。

這個網站是 Lotus Domino 應用程式，用 HTTP Basic Auth 保護在資料庫層級
（不是表單登入），純用 requests 就能把整個流程走完，不需要瀏覽器/Playwright。

流程（對應人工操作的每一步）：
  1. 「出貨-依建立日」列表（FVEWFV02）— 依出貨單號找到文件 UID。
     出貨單號本身帶建立日期（如 S2260902108＝2026/09/02 第108筆），列表本身
     就是依建立日新到舊排序，正常使用（列印剛建立沒多久的出貨單）第一頁就
     找得到；找不到才跟著「下一頁」按鈕用的 Domino Click token 翻頁。
  2. 開啟該筆文件（EditDocument）— 取得 SeleNum（預設全選的標籤序號清單）。
  3. 呼叫「列印選取標籤(標籤機)」按鈕實際打的網址（RunAgent，
     Function=PrtSele2DBarcode4Supplier）— 拿到乾淨的標籤 HTML，每張標籤
     是一個 <table class='lable'>，內含料號/Project/數量/PO/標籤序號(BCode#)/
     品名規格，以及一個指向 QR Code 圖片 servlet 的 <iframe src>。
  4. 把每張標籤重新畫成一張圖片（文字用 Pillow 畫，QR Code 直接下載原圖貼上，
     兩者內容完全一致，不是重新產生的假 QR），組成一份 PDF。
"""
import html
import io
import re
import zipfile

import requests
from PIL import Image, ImageDraw, ImageFont

BASE = "http://SLP.gmmcorp.com.tw:8080/apps/lg_001.nsf"
VIEW_URL = f"{BASE}/FVEWFV02?OpenForm"

_FONT_PATH = r"C:\Windows\Fonts\msjh.ttc"  # 微軟正黑體，Windows 內建


def _font(size: int):
    try:
        return ImageFont.truetype(_FONT_PATH, size)
    except Exception:
        return ImageFont.load_default()


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


def _fetch_labels(session: requests.Session, doc_uid: str, sele_num: str) -> list[dict]:
    """呼叫「列印選取標籤(標籤機)」實際打的網址，解析出每一張標籤的資料。"""
    url = (f"{BASE}/RunAgent?OpenAgent&Function=PrtSele2DBarcode4Supplier"
           f"&ParentUNID={doc_uid}&SeleNum={sele_num}")
    page_html = _get(session, url, timeout=30).text

    labels = []
    for block in re.findall(r"<table class='lable'.*?</table>", page_html, re.DOTALL):
        def _grab(pat, default=""):
            m = re.search(pat, block, re.DOTALL)
            return html.unescape(m.group(1).strip()) if m else default

        qr_m = re.search(r"src='([^']+QRCode[^']+)'", block)
        qr_url = qr_m.group(1) if qr_m else ""
        if qr_url and not qr_url.startswith("http"):
            qr_url = "http://SLP.gmmcorp.com.tw:8080" + qr_url

        labels.append({
            "item":     _grab(r"Item\s*:\s*([^<]+)</td>"),
            "project":  _grab(r"Project\s*:\s*([^<]+)</td>"),
            "qty_line": _grab(r"Qty\s*:\s*([^<]+)</td>"),
            "bcode":    _grab(r"BCode#\s*:\s*([^<]+)<br>"),
            "dates":    _grab(r"BCode#[^<]*<br>([^<]+)<br>"),
            "desc":     _grab(r"Desc\s*:\s*([^<]+)</td>"),
            "qr_url":   qr_url,
        })
    return labels


def _render_label(session: requests.Session, label: dict) -> Image.Image:
    """把一張標籤的文字資料＋QR Code 畫成一張乾淨的圖片（貼進 BarTender 用）。"""
    W, H = 900, 420
    img = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(img)

    pad = 22
    f_normal = _font(30)
    f_small = _font(24)
    qr_size = 220

    lines = [
        (f"料號：{label['item']}", f_normal),
        (f"Project：{label['project']}", f_small),
        (label["qty_line"], f_small),
        (f"標籤序號：{label['bcode']}", f_small),
        (label["dates"], f_small),
    ]
    y = pad
    for text, font in lines:
        draw.text((pad, y), text, fill="black", font=font)
        y += font.size + 10

    desc = label.get("desc", "")
    if desc:
        y += 8
        max_chars = 22
        for i in range(0, len(desc), max_chars):
            draw.text((pad, y), desc[i:i + max_chars], fill="black", font=f_small)
            y += f_small.size + 6

    if label["qr_url"]:
        try:
            qr_img = Image.open(io.BytesIO(_get(session, label["qr_url"], timeout=15).content))
            qr_img = qr_img.convert("RGB").resize((qr_size, qr_size))
            img.paste(qr_img, (W - qr_size - pad, pad))
        except Exception:
            pass  # QR 抓不到就留白，不擋整張標籤的其他文字

    return img


def download_label_pdfs(
    order_nos: list,
    username: str = "110488",
    password: str = "4667044110488",
) -> tuple[dict, dict]:
    """
    登入均華供應商平台，依出貨單號抓取標籤資料並重新畫成 PDF（每張標籤一頁）。
    回傳 (results, errors)：
      results: {order_no: pdf_bytes | None}
      errors:  {order_no: error_str}
    格式跟 erp_downloader.download_label_pdfs 一致，可以直接沿用 app.py
    既有的「複製截圖」／打包 zip 流程，呼叫端不用另外寫一套。
    """
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
            labels = _fetch_labels(session, doc_uid, sele_num)
            if not labels:
                raise RuntimeError("查得到這張出貨單，但解析不到任何標籤內容")

            imgs = [_render_label(session, lb) for lb in labels]
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

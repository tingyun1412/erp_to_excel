"""
銷貨單 RTF 解析模組
RTF 文字框值的模式：\\loch\\f18 VALUE（每個值重複兩次，取一次）

Raw values 觀察順序（以 202606170012-ASEK.rtf 為例）：
  C0177-1        ← 客戶代號（header 區，忽略或存為 customer_code）
  76027628       ← 賣方統編
  107            ← 忽略（頁碼）
  1              ← 忽略
  2026/06/17     ← 銷貨日期
  202606170012   ← 銷貨單號
  6100836338     ← 買方統編
  BA30737334     ← 買方備用統編
  07-3617131     ← 電話
  07-3617352     ← 傳真
  1 / 1          ← 頁數（忽略）
  GC             ← 業務代號（忽略）
  202512050003   ← 客戶訂單號
  #15313         ← 聯絡人工號
  RD=0.89,...    ← 品名規格
  0001           ← 項次（忽略）
  3              ← 稅率（忽略）
  BAS089109BREB  ← 料號
  LSCR-043-043-S ← 客戶料號（緊跟在料號後）
  200            ← 數量
"""
import re
from pathlib import Path


def extract_field_values(rtf_bytes: bytes) -> list[str]:
    """從 RTF bytes 擷取所有文字框的值（去除重複）"""
    text = rtf_bytes.decode("cp950", errors="replace")
    pattern = re.compile(r"\\loch\\f18 ([^\r\n\\}]+)")
    seen = set()
    values = []
    for m in pattern.findall(text):
        v = m.strip()
        if v and v not in seen and v != ":":
            seen.add(v)
            values.append(v)
    return values


def _looks_like_item_no(v: str) -> bool:
    """自家料號：2碼以上大寫英文開頭，接數字/大寫/連字號，總長6碼以上"""
    return bool(re.match(r"^[A-Z]{2,}[0-9A-Z\-]{4,}$", v)) and len(v) >= 6


def parse_sales_order_rtf(file_path) -> dict:
    """
    解析銷貨單 RTF，回傳結構化字典

    回傳欄位：
        filename, order_no, order_date, customer_order_no,
        seller_tax_id, buyer_tax_id, phone, fax, contact,
        items: list of {
            item_no, description, quantity, unit,
            unit_price, amount, remark, ship_date, customer
        }
        raw_values（除錯用）
    """
    path = Path(file_path)
    raw  = path.read_bytes()
    values = extract_field_values(raw)

    result = {
        "filename":          path.name,
        "order_no":          "",
        "order_date":        "",
        "customer_order_no": "",
        "seller_tax_id":     "",
        "buyer_tax_id":      "",
        "phone":             "",
        "fax":               "",
        "contact":           "",
        "items":             [],
        "raw_values":        values,
    }

    # 從檔名抓客戶（格式：單號-客戶.rtf）
    name_parts = path.stem.split("-", 1)
    customer_from_filename = name_parts[1] if len(name_parts) > 1 else ""

    # ── 兩段式解析 ────────────────────────────────────────────
    # header 區：遇到第一個規格字串（含 =）或料號之前
    # 品項區：遇到料號之後

    tax_ids     = []
    header_done = False   # 進入品項區的旗標
    pending_desc = ""     # 目前累積的品名規格

    for v in values:
        # ── 任何區都處理的 header 欄位 ───────────────────────

        # 日期 YYYY/MM/DD
        if re.match(r"^\d{4}/\d{2}/\d{2}$", v):
            result["order_date"] = v.replace("/", "")
            continue

        # 12碼數字（單號 / 客戶訂單號）
        if re.match(r"^\d{12}$", v):
            if not result["order_no"]:
                result["order_no"] = v
            elif not result["customer_order_no"] and v != result["order_no"]:
                result["customer_order_no"] = v
            continue

        # 統編（8-10碼數字，或英文前綴統編）
        if re.match(r"^\d{8,10}$", v):
            tax_ids.append(v)
            continue
        if re.match(r"^[A-Z]{1,2}\d{8}$", v):
            tax_ids.append(v)
            continue

        # 電話/傳真
        if re.match(r"^\d{2,4}-\d{6,8}$", v):
            if not result["phone"]:   result["phone"] = v
            elif not result["fax"]:  result["fax"]   = v
            continue

        # 聯絡人工號
        if re.match(r"^#\d+$", v):
            result["contact"] = v
            continue

        # 頁數格式（1 / 1）
        if re.match(r"^\d+\s*/\s*\d+$", v):
            continue

        # header 區的小數字（項次、稅率、頁碼）
        if re.match(r"^\d{1,4}$", v) and not header_done:
            continue

        # 品名規格（含 =）→ 進入品項準備區
        if "=" in v:
            pending_desc = v
            header_done = True
            continue

        # ── 品項區 ──────────────────────────────────────────
        if _looks_like_item_no(v):
            header_done = True
            item = {
                "item_no":     v,
                "description": pending_desc,
                "quantity":    0,
                "unit":        "PC",
                "unit_price":  0,
                "amount":      0,
                "remark":      "",   # 客戶料號，在料號之後才填
                "ship_date":   result["order_date"],
                "customer":    customer_from_filename,
            }
            result["items"].append(item)
            pending_desc = ""
            continue

        # 客戶料號：字母開頭含連字號，但只在「已有料號、且在品項區」才接受
        if header_done and re.match(r"^[A-Za-z][A-Za-z0-9\-_\.]{3,}$", v) and "-" in v:
            if result["items"] and not result["items"][-1]["remark"]:
                result["items"][-1]["remark"] = v
            continue

        # 數量（只在品項區接受）
        if header_done and re.match(r"^\d{1,6}$", v):
            qty = int(v)
            if 1 <= qty <= 99999:
                for item in reversed(result["items"]):
                    if item["quantity"] == 0:
                        item["quantity"] = qty
                        break
            continue

    # ── 統編分配 ─────────────────────────────────────────────
    for tid in tax_ids:
        if re.match(r"^\d{8}$", tid) and not result["seller_tax_id"]:
            result["seller_tax_id"] = tid
        elif not result["buyer_tax_id"]:
            result["buyer_tax_id"] = tid

    # order_date fallback
    if not result["order_date"] and result["order_no"]:
        result["order_date"] = result["order_no"][:8]
    for item in result["items"]:
        if not item["ship_date"]:
            item["ship_date"] = result["order_date"]

    # ── 後處理：合併被誤認為料號的客戶料號 ───────────────────
    result["items"] = _postprocess_items(result["items"])

    return result


def _postprocess_items(items: list[dict]) -> list[dict]:
    """
    若某 item 沒有 description 且前一個 item 的 quantity == 0，
    代表這個 item_no 其實是前一個的客戶料號，合併並轉移數量。
    """
    if len(items) < 2:
        return items
    result = [items[0]]
    for item in items[1:]:
        prev = result[-1]
        if item["description"] == "" and prev["quantity"] == 0:
            if not prev["remark"]:
                prev["remark"] = item["item_no"]
            prev["quantity"] = item["quantity"]
        else:
            result.append(item)
    return result


def parse_multiple_rtf(file_list) -> list[dict]:
    results = []
    for f in file_list:
        try:
            results.append(parse_sales_order_rtf(f))
        except Exception as e:
            results.append({
                "filename":   str(f),
                "error":      str(e),
                "items":      [],
                "raw_values": [],
            })
    return results
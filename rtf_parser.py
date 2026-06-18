"""
銷貨單 RTF 解析模組
RTF 文字框值的模式：\\loch\\f18 VALUE（每個值重複兩次，取一次）

Raw values 觀察順序（以 202606170012-ASEK.rtf 為例）：
  C0177-1        ← 客戶料號（放在最前面）
  76027628       ← 賣方統編
  107            ← 忽略（頁碼/流水）
  1              ← 忽略
  2026/06/17     ← 銷貨日期
  202606170012   ← 銷貨單號
  6100836338     ← 買方統編
  BA30737334     ← 買方備用
  07-3617131     ← 電話
  07-3617352     ← 傳真
  1 / 1          ← 頁數
  GC             ← 業務代號（忽略）
  202512050003   ← 客戶訂單號
  #15313         ← 聯絡人
  RD=0.89,...    ← 品名規格
  0001           ← 項次（忽略）
  3              ← 稅率（忽略）
  BAS089109BREB  ← 料號
  LSCR-043-043-S ← 客戶料號（緊跟在料號後）
  200            ← 數量（最後）
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
    """料號：2碼以上大寫英文開頭，接數字/大寫/連字號，總長6碼以上"""
    return bool(re.match(r"^[A-Z]{2,}[0-9A-Z\-]{4,}$", v)) and len(v) >= 6


def _looks_like_qty(v: str, context_after_item: bool) -> bool:
    """數量：純數字，且在解析到料號「之後」才接受"""
    if not re.match(r"^\d{1,6}$", v):
        return False
    qty = int(v)
    # 排除常見非數量的小數字（項次 0001 在前面已過濾，這裡排除 1-9 的小數）
    # 若確認在料號後，接受 1 以上的合理數量
    return context_after_item and 1 <= qty <= 99999


def parse_sales_order_rtf(file_path) -> dict:
    """解析銷貨單 RTF，回傳結構化字典"""
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
    # 第一段：抓 header 欄位（日期、統編、電話等）
    # 第二段：抓品項（料號→客戶料號→數量 的重複組合）

    tax_ids      = []
    header_done  = False   # 遇到第一個料號後視為進入品項區
    pending_desc = ""
    pending_cusno_before = []  # 在料號之前碰到的疑似客戶料號

    for v in values:
        # ── Header 區通用規則（不管在哪都執行）──────────────
        # 日期
        if re.match(r"^\d{4}/\d{2}/\d{2}$", v):
            result["order_date"] = v.replace("/", "")
            continue

        # 12碼數字（單號/客戶訂單號）
        if re.match(r"^\d{12}$", v):
            if not result["order_no"]:
                result["order_no"] = v
            elif not result["customer_order_no"] and v != result["order_no"]:
                result["customer_order_no"] = v
            continue

        # 統編
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

        # 品名規格（含 =）
        if "=" in v:
            pending_desc = v
            continue

        # 忽略：頁數格式、純小數字（項次、稅率）
        if re.match(r"^\d+\s*/\s*\d+$", v):
            continue
        if re.match(r"^\d{1,4}$", v) and not header_done:
            continue  # header 區的小數字（頁碼、稅率）忽略

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
                "remark":      pending_cusno_before[-1] if pending_cusno_before else "",
                "ship_date":   result["order_date"],
                "customer":    customer_from_filename,
            }
            result["items"].append(item)
            pending_desc = ""
            pending_cusno_before = []
            continue

        # 客戶料號（字母開頭，含連字號，長度合理）
        if re.match(r"^[A-Za-z][A-Za-z0-9\-_\.]{3,}$", v) and "-" in v:
            if header_done and result["items"] and not result["items"][-1]["remark"]:
                result["items"][-1]["remark"] = v
            elif not header_done:
                pending_cusno_before.append(v)
            continue

        # 數量（只在品項區、料號後接受）
        if header_done and re.match(r"^\d{1,6}$", v):
            qty = int(v)
            if 1 <= qty <= 99999:
                # 找最後一個數量還是 0 的 item
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

    result["items"] = _postprocess_items(result["items"])
    return result


def parse_multiple_rtf(file_list) -> list[dict]:
    results = []
    for f in file_list:
        try:
            results.append(parse_sales_order_rtf(f))
        except Exception as e:
            results.append({"filename": str(f), "error": str(e), "items": [], "raw_values": []})
    return results

# ── 後處理：修正「客戶料號被誤認為料號」的情況 ──────────────────
# 若 item_no 候選的數量是 0，且前一個 item 的 remark 是空的，
# 且前一個 item 的數量也是 0，
# 則把這個 item 合併到前一個（當作客戶料號）
def _postprocess_items(items: list[dict]) -> list[dict]:
    """後處理：合併被誤認為料號的客戶料號

    規則：若某個 item 沒有 description（規格），且前一個 item 的 quantity == 0，
    則這個 item 很可能是前一個的客戶料號，把它的 item_no 設為前一個的 remark，
    並把數量轉移給前一個。
    """
    if len(items) < 2:
        return items
    result = [items[0]]
    for item in items[1:]:
        prev = result[-1]
        is_likely_customer_part_no = (
            item["description"] == "" and   # 沒有規格說明
            prev["quantity"] == 0            # 前一個料號還沒有數量
        )
        if is_likely_customer_part_no:
            # 把這個 item 合併到前一個
            if not prev["remark"]:
                prev["remark"] = item["item_no"]
            prev["quantity"] = item["quantity"]  # 數量轉移
        else:
            result.append(item)
    return result

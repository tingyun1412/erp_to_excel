"""
銷貨單 RTF 解析模組

品項欄位順序（固定，不靠前綴判斷）：

格式A（銷貨單1,3,4）：
  項次(0001) → 規格(0~多行) → 數量(xxx.000) → 你們料號 → 客戶料號(備註) → 批號(12碼)

格式B（銷貨單2，規格在項次前）：
  規格 → 項次 → 數量(xxx.000) → 你們料號 → 客戶料號(備註)

判斷你們料號 vs 客戶料號：純靠位置，先出現的是你們的，後出現的是客戶的。
"""
import re
from pathlib import Path


def extract_field_values(rtf_bytes: bytes) -> list[str]:
    """
    每個欄位在 RTF 中出現兩次（文字框格式），
    所以只跳過「相鄰重複」，不做全域去重。
    這樣 0004, 0006... 各自的 100.000 都能保留。
    """
    text = rtf_bytes.decode("cp950", errors="replace")
    pattern = re.compile(r"\\loch\\f18 ([^\r\n\\}]+)")
    values = []
    prev = None
    for m in pattern.findall(text):
        v = m.strip()
        if v and v != ":" and v != prev:
            values.append(v)
        prev = v
    return values


# ── 型別判斷（只判斷明確格式，不判斷料號前綴）────────────────────

def _is_seq(v):
    """項次：0001-0099"""
    return bool(re.match(r'^0[0-9]{3}$', v))

def _is_qty(v):
    """數量：xxx.000（含三位小數）"""
    return bool(re.match(r'^\d+\.\d{3}$', v))

def _is_lot_no(v):
    """批號：12碼純數字，或 R/P 開頭+數字"""
    if re.match(r'^\d{12}$', v): return True
    if re.match(r'^[RP]\d{9,}$', v): return True
    return False

def _is_date(v):
    return bool(re.match(r'^\d{4}/\d{2}/\d{2}$', v))

def _is_phone(v):
    return bool(re.match(r'^\d{2,4}-\d{4,8}$', v)) or bool(re.match(r'^\d{4}-\d{3}-\d{3}$', v))

def _is_page_fraction(v):
    return bool(re.match(r'^\d+\s*/\s*\d+$', v))

def _is_tax_id(v):
    if re.match(r'^\d{8}$', v): return True
    if re.match(r'^[A-Z]{2}\d{8}$', v): return True
    if re.match(r'^\d{3}-\d{9,}$', v): return True
    if re.match(r'^[A-Z]{3,5}-\d{8,}$', v): return True
    return False

def _is_part_no(v):
    """任何料號（你們的或客戶的）：英數字混合，長度>=5，可含連字號"""
    if not v or len(v) < 5: return False
    if _is_seq(v): return False
    if _is_qty(v): return False
    if _is_lot_no(v): return False
    if _is_page_fraction(v): return False
    if re.match(r'^\d+$', v): return False          # 純數字不是料號
    if re.match(r'^\d+[.,]\d+$', v): return False  # 數字含小數不是料號
    # 英數字組合（可含 - * . / 等符號）
    return bool(re.match(r'^[A-Za-z0-9][A-Za-z0-9\-\*\./_]{3,}$', v))

def _is_spec(v):
    """規格：含技術符號"""
    if _is_part_no(v) and not re.search(r'[=,*"um]', v): return False
    return bool(re.search(r'[=,]|\*\d|\d\*|um|mm|OD|ID|PD|RD|BeCu|\d+\.\d+mm', v))

def _clean_qty(v):
    try:
        return int(float(v))
    except Exception:
        return 0


# ══════════════════════════════════════════════════════════════════

def parse_sales_order_rtf(file_path) -> dict:
    path = Path(file_path)
    raw  = path.read_bytes()
    values = extract_field_values(raw)

    name_parts = path.stem.split("-", 1)
    customer_from_filename = name_parts[1] if len(name_parts) > 1 else ""

    result = {
        "filename":          path.name,
        "order_no":          "",
        "order_date":        "",
        "customer_order_no": "",
        "seller_tax_id":     "",
        "buyer_tax_id":      "",
        "buyer_id_raw":      "",
        "phone":             "",
        "fax":               "",
        "mobile":            "",
        "contact":           "",
        "customer_code":     "",
        "items":             [],
        "raw_values":        values,
    }

    # ── Header 區 ─────────────────────────────────────────────
    header_end = next((i for i, v in enumerate(values) if _is_seq(v)), len(values))

    tax_ids, phones = [], []
    header_seen = set()  # header 內去重，避免多頁重複header被重複解析
    for v in values[:header_end]:
        if v in header_seen:
            continue
        header_seen.add(v)
        if re.match(r'^C\d{4}(-\d+)?$', v):
            result["customer_code"] = v
        elif re.match(r'^#\d+$', v):
            result["contact"] = v
        elif _is_date(v):
            result["order_date"] = v.replace("/", "")
        elif re.match(r'^\d{12}$', v):
            if not result["order_no"]:
                result["order_no"] = v
            elif not result["customer_order_no"] and v != result["order_no"]:
                result["customer_order_no"] = v
        elif _is_tax_id(v):
            if v not in tax_ids:
                tax_ids.append(v)
        elif _is_phone(v):
            if v not in phones:
                phones.append(v)

    for tid in tax_ids:
        if re.match(r'^\d{8}$', tid) and not result["seller_tax_id"]:
            result["seller_tax_id"] = tid
        elif not result["buyer_tax_id"]:
            result["buyer_tax_id"] = tid
            result["buyer_id_raw"] = tid

    for i, p in enumerate(phones):
        if i == 0:   result["phone"]  = p
        elif i == 1: result["fax"]    = p
        elif i == 2: result["mobile"] = p

    if not result["order_date"] and result["order_no"]:
        result["order_date"] = result["order_no"][:8]

    # ── 品項區 ────────────────────────────────────────────────
    # 收集 header 中已知的值，多頁 RTF 會在品項區重複出現 header，要跳過
    header_known = set(values[:header_end])

    item_vals = [v for v in values[header_end:]
                 if v not in header_known or _is_seq(v) or _is_qty(v)]
    # 保留：項次、數量一定要留；其他 header 值過濾掉

    # 偵測格式B：第一個項次前有規格字串
    first_seq_pos = next((i for i, v in enumerate(item_vals) if _is_seq(v)), None)
    format_b = (
        first_seq_pos is not None
        and first_seq_pos > 0
        and any(_is_spec(v) for v in item_vals[:first_seq_pos])
    )

    if format_b:
        result["items"] = _parse_format_b(item_vals, result["order_date"], customer_from_filename)
    else:
        result["items"] = _parse_format_a(item_vals, result["order_date"], customer_from_filename)

    return result


def _blank_item(seq, ship_date, customer):
    return {
        "seq":         seq,
        "item_no":     "",   # 你們的料號（位置第一個）
        "description": "",   # 規格
        "quantity":    0,
        "unit":        "PCS",
        "unit_price":  0,
        "amount":      0,
        "remark":      "",   # 客戶料號（位置第二個）
        "lot_no":      "",   # 批號
        "ship_date":   ship_date,
        "customer":    customer,
    }


def _parse_format_a(vals, ship_date, customer):
    """
    格式A：項次 → 規格(0~多行) → 數量(xxx.000) → 你們料號 → 客戶料號 → 批號
    """
    seq_positions = [i for i, v in enumerate(vals) if _is_seq(v)]
    items = []

    for si, seq_pos in enumerate(seq_positions):
        end = seq_positions[si + 1] if si + 1 < len(seq_positions) else len(vals)
        seg = vals[seq_pos + 1 : end]

        item = _blank_item(vals[seq_pos], ship_date, customer)
        desc_parts = []
        # 狀態機：spec → after_qty（遇到數量後）→ got_item_no（拿到料號後）
        state = "spec"

        for v in seg:
            if _is_page_fraction(v) or _is_seq(v):
                continue

            if state == "spec":
                if _is_qty(v):
                    item["quantity"] = _clean_qty(v)
                    state = "after_qty"
                elif _is_lot_no(v):
                    item["lot_no"] = v
                elif _is_part_no(v) and not _is_spec(v):
                    # 沒數量直接碰到料號
                    item["item_no"] = v
                    state = "got_item_no"
                else:
                    desc_parts.append(v)

            elif state == "after_qty":
                if _is_lot_no(v):
                    item["lot_no"] = v
                elif _is_part_no(v):
                    item["item_no"] = v
                    state = "got_item_no"

            elif state == "got_item_no":
                if _is_lot_no(v):
                    item["lot_no"] = v
                elif _is_part_no(v) and not item["remark"]:
                    item["remark"] = v  # 客戶料號

        item["description"] = " ".join(desc_parts).strip()
        items.append(item)

    return items


def _parse_format_b(vals, ship_date, customer):
    """
    格式B：規格 → 項次 → [數量] → 你們料號 → 客戶料號
    規格在項次前面
    """
    seq_positions = [i for i, v in enumerate(vals) if _is_seq(v)]
    items = []
    consumed_up_to = 0

    for si, seq_pos in enumerate(seq_positions):
        # 規格：上一個品項消耗到的位置 ~ 本項次前
        spec_seg = vals[consumed_up_to : seq_pos]
        desc_parts = [v for v in spec_seg if _is_spec(v) and not _is_part_no(v)]

        # 本項次後的值
        next_seq = seq_positions[si + 1] if si + 1 < len(seq_positions) else len(vals)
        after_seg = vals[seq_pos + 1 : next_seq]

        item = _blank_item(vals[seq_pos], ship_date, customer)
        item["description"] = " ".join(desc_parts).strip()

        last_consumed = seq_pos + 1

        for j, v in enumerate(after_seg):
            abs_idx = seq_pos + 1 + j
            if _is_page_fraction(v) or _is_seq(v):
                continue
            if _is_qty(v) and item["quantity"] == 0:
                item["quantity"] = _clean_qty(v)
                last_consumed = abs_idx + 1
            elif _is_lot_no(v) and not item["lot_no"]:
                item["lot_no"] = v
                last_consumed = abs_idx + 1
            elif _is_part_no(v) and not item["item_no"]:
                item["item_no"] = v   # 你們的料號（第一個）
                last_consumed = abs_idx + 1
            elif _is_part_no(v) and not item["remark"]:
                item["remark"] = v    # 客戶料號（第二個）
                last_consumed = abs_idx + 1
            elif _is_spec(v) and not _is_part_no(v):
                # 遇到下一個品項的規格，停止
                break

        consumed_up_to = last_consumed
        items.append(item)

    return items


def parse_multiple_rtf(file_list) -> list[dict]:
    results = []
    for f in file_list:
        try:
            results.append(parse_sales_order_rtf(f))
        except Exception as e:
            results.append({"filename": str(f), "error": str(e), "items": [], "raw_values": []})
    return results
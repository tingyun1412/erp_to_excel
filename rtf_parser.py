"""
銷貨單 RTF 解析模組

格式A（銷貨單1,3,4）：項次 → 品名(中文) → 規格 → 數量 → 料號 → 客戶料號/批號
格式B（銷貨單2，規格在項次前）：規格 → 項次 → 數量 → 料號 → 客戶料號

批號 vs 客戶料號分類規則（兩階段）：
  - 若該品項段落含 12 位純數字批號 → R/P+數字視為客戶料號
  - 若無 12 位純數字批號 → R/P+數字視為批號
  - 遇到「以下空白」等頁尾標記後停止收集
"""
import re
from pathlib import Path

_FOOTER_STOP = {"以下空白"}

# 換頁時會重複印出的表頭/簽收區標籤，解析品名/規格時應跳過（不可當作品名）
_HEADER_NOISE_LABELS = {
    "客戶代號", "客戶名稱", "聯", "絡", "人", "統一編號", "送貨地址",
    "單據日期", "單據號碼", "訂單單號", "發票號碼", "送貨電話",
    "傳\u3000\u3000真", "行動電話", "聯絡電話", "頁\u3000\u3000次",
    "序號", "品名／規格", "單位", "數量", "銷貨單", "備註",
    "GC", "單號", "客戶簽收：", "業務：", "審核：", "經辦人：",
}


def _is_header_noise(v: str) -> bool:
    """判斷是否為換頁重印的表頭/簽收區文字（不應視為品名或規格）"""
    if v in _HEADER_NOISE_LABELS:
        return True
    if "公司" in v:
        return True  # 供應商/客戶公司全名重印
    if any(kw in v for kw in ("市", "區", "縣", "路", "街")):
        return True  # 送貨地址重印
    if v in {"號", "樓", "路", "街", "巷", "弄", "段", "市", "區", "縣"}:
        return True  # 地址片段重印
    if re.match(r'^\d{1,3}$', v):
        return True  # 地址門牌號等短數字片段
    if _is_date(v) or _is_phone(v) or _is_page_fraction(v) or _is_tax_id(v):
        return True
    if re.match(r'^#\d+$', v):
        return True
    if re.match(r'^C\d{4}(-\d+)?$', v):
        return True
    return False


def extract_field_values(rtf_bytes: bytes) -> list[str]:
    """
    提取 RTF 中所有欄位值（ASCII + 中文），按檔案位置排序，相鄰重複去除。
    ASCII 值來自 \\loch\\f18；中文值來自 \\hich\\af1\\dbch\\f18 的 Hex 編碼。
    """
    text = rtf_bytes.decode("cp950", errors="replace")
    matches = []

    for m in re.finditer(r"\\loch\\f18 ([^\r\n\\}]+)", text):
        v = m.group(1).strip()
        if v and v not in (":", " ", "  "):
            matches.append((m.start(), v))

    _p_dbch = re.compile(r"\\hich\\af1\\dbch\\f18 ((?:\\'[0-9a-fA-F]{2}[\r\n\s]*)+)")
    for m in _p_dbch.finditer(text):
        hex_part = m.group(1)
        try:
            hbytes = bytes(int(x, 16) for x in re.findall(r"'([0-9a-fA-F]{2})", hex_part))
            decoded = hbytes.decode("cp950", errors="replace").strip()
            if decoded:
                matches.append((m.start(), decoded))
        except Exception:
            pass

    matches.sort(key=lambda x: x[0])
    values: list[str] = []
    prev = None
    for _, v in matches:
        if v != prev:
            values.append(v)
        prev = v
    return values


def _extract_customer_name_from_values(values: list[str]) -> str:
    """
    從欄位值序列中提取客戶公司名稱。
    公司名稱在原始文件中常被拆成多個片段（例如「鼎元光電科技」「(」「股」「)」「公司竹南分公司」），
    所以從客戶代號之後開始，把連續片段接起來，直到累積字串第一次出現「公司」為止才停止
    （避免把後面重複列印的第二份名稱也接進來）。
    """
    code_idx = next(
        (i for i, v in enumerate(values) if re.match(r"^C\d{4}(-\d+)?$", v)),
        None,
    )
    if code_idx is None:
        return ""
    parts = []
    for v in values[code_idx + 1: code_idx + 1 + 10]:
        if v in ("客戶名稱", "客戶代號"):
            continue
        if not v or _is_date(v) or _is_phone(v) or re.match(r"^#\d+$", v):
            break
        parts.append(v)
        if "公司" in "".join(parts):
            break
    return "".join(parts)


# ── 型別判斷 ──────────────────────────────────────────────────────

def _is_seq(v):
    return bool(re.match(r'^0[0-9]{3}$', v))

def _is_qty(v):
    return bool(re.match(r'^\d+\.\d{3}$', v))

def _is_12digit_lot(v):
    return bool(re.match(r'^\d{12}$', v))

def _is_rp_lot(v):
    return bool(re.match(r'^[RP]\d{9,}$', v))

def _is_lot_no(v):
    return _is_12digit_lot(v) or _is_rp_lot(v)

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

def _is_chinese(v):
    return bool(re.search(r'[一-鿿㐀-䶿]', v))

def _is_part_no(v):
    if not v or len(v) < 5: return False
    if _is_seq(v): return False
    if _is_qty(v): return False
    if _is_lot_no(v): return False
    if _is_page_fraction(v): return False
    if re.match(r'^\d+$', v): return False
    if re.match(r'^\d+[.,]\d+$', v): return False
    return bool(re.match(r'^[A-Za-z0-9][A-Za-z0-9\-\*\./_]{3,}$', v))

def _is_spec(v):
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
        "customer_name":     "",
        "items":             [],
        "raw_values":        values,
    }

    header_end = next((i for i, v in enumerate(values) if _is_seq(v)), len(values))

    tax_ids, phones = [], []
    header_seen = set()
    customer_code_found = False
    for v in values[:header_end]:
        if v in header_seen:
            continue
        header_seen.add(v)
        if re.match(r'^C\d{4}(-\d+)?$', v):
            result["customer_code"] = v
            customer_code_found = True
        elif re.match(r'^#\d+$', v):
            result["contact"] = v
        elif _is_date(v):
            result["order_date"] = v.replace("/", "")
        elif _is_12digit_lot(v):
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

    result["customer_name"] = _extract_customer_name_from_values(values[:header_end])

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

    # 過濾表頭重印雜訊，同時排除 header 內出現的中文（例如聯絡人姓名）再次出現在品項中
    header_known_chinese = {v for v in values[:header_end] if _is_chinese(v)}
    item_vals = [v for v in values[header_end:]
                 if not _is_header_noise(v)
                 and not (_is_chinese(v) and v in header_known_chinese and not _is_seq(v))]

    first_seq_pos = next((i for i, v in enumerate(item_vals) if _is_seq(v)), None)
    format_b = (
        first_seq_pos is not None
        and first_seq_pos > 0
        and any(_is_spec(v) for v in item_vals[:first_seq_pos])
    )

    # 格式C：項次 → 數量 → 單位 → 料號 → 客戶料號 → 品名 → 規格
    # 判斷依據：項次後緊接著就是數量（格式A/B皆非如此）
    format_c = False
    if not format_b and first_seq_pos is not None:
        nxt = item_vals[first_seq_pos + 1] if first_seq_pos + 1 < len(item_vals) else ""
        format_c = _is_qty(nxt)

    if format_b:
        result["items"] = _parse_format_b(item_vals, result["order_date"], customer_from_filename)
    elif format_c:
        result["items"] = _parse_format_c(item_vals, result["order_date"], customer_from_filename)
    else:
        result["items"] = _parse_format_a(item_vals, result["order_date"], customer_from_filename)

    return result


def _blank_item(seq, ship_date, customer):
    return {
        "seq":         seq,
        "item_no":     "",
        "name":        "",
        "description": "",
        "quantity":    0,
        "unit":        "PCS",
        "unit_price":  0,
        "amount":      0,
        "remark":      "",
        "lot_no":      "",
        "ship_date":   ship_date,
        "customer":    customer,
    }


def _classify_post_item(item: dict, post_vals: list):
    """
    兩階段分類：
    有 12 位批號 → 12位=批號，R/P+數字=客戶料號
    無 12 位批號 → R/P+數字=批號，其他料號=客戶料號
    """
    has_12digit = any(_is_12digit_lot(v) for v in post_vals)

    for v in post_vals:
        if _is_12digit_lot(v):
            if not item["lot_no"]:
                item["lot_no"] = v
        elif _is_rp_lot(v):
            if has_12digit:
                if not item["remark"]:
                    item["remark"] = v
            else:
                if not item["lot_no"]:
                    item["lot_no"] = v
        elif _is_part_no(v):
            if not item["remark"]:
                item["remark"] = v


def _parse_format_a(vals, ship_date, customer):
    """格式A：項次 → 品名(中文) → 規格 → 數量 → 料號 → 客戶料號/批號"""
    seq_positions = [i for i, v in enumerate(vals) if _is_seq(v)]
    items = []

    for si, seq_pos in enumerate(seq_positions):
        end = seq_positions[si + 1] if si + 1 < len(seq_positions) else len(vals)
        seg = vals[seq_pos + 1 : end]

        item = _blank_item(vals[seq_pos], ship_date, customer)
        desc_parts = []
        desc_seen = set()   # 同段落規格去重（避免 RTF 文字框二次重複）
        post_vals = []
        state = "spec"
        _building_name = False  # 是否正在累積中文品名（連續中文串）

        for v in seg:
            if _is_page_fraction(v) or _is_seq(v):
                continue
            if v in _FOOTER_STOP:
                break

            if state == "spec":
                if _is_qty(v):
                    _building_name = False
                    item["quantity"] = _clean_qty(v)
                    state = "after_qty"
                elif _is_chinese(v) and v not in _FOOTER_STOP:
                    if len(v) >= 2:
                        if not item["name"]:
                            item["name"] = v
                            _building_name = True
                        elif _building_name:
                            item["name"] += v   # 連續中文 → 合併為品名
                        # else: 非連續中文 → 略過（第二份複本或其他）
                    # 單字中文（如「度」）不加入任何欄位
                elif _is_part_no(v) and not _is_spec(v):
                    _building_name = False
                    item["item_no"] = v
                    state = "got_item_no"
                else:
                    _building_name = False
                    if v not in desc_seen:
                        desc_parts.append(v)
                        desc_seen.add(v)

            elif state == "after_qty":
                if _is_chinese(v):
                    pass  # 品名應在 spec 階段已設定
                elif _is_part_no(v):
                    item["item_no"] = v
                    state = "got_item_no"

            elif state == "got_item_no":
                if v in _FOOTER_STOP:
                    break
                if not _is_chinese(v):
                    post_vals.append(v)

        _classify_post_item(item, post_vals)
        item["description"] = " ".join(desc_parts).strip()
        items.append(item)

    return items


def _parse_format_b(vals, ship_date, customer):
    """格式B：規格 → 項次 → 數量 → 料號 → 客戶料號"""
    seq_positions = [i for i, v in enumerate(vals) if _is_seq(v)]
    items = []
    consumed_up_to = 0

    for si, seq_pos in enumerate(seq_positions):
        spec_seg = vals[consumed_up_to : seq_pos]
        desc_parts = [v for v in spec_seg if _is_spec(v) and not _is_part_no(v)]

        next_seq = seq_positions[si + 1] if si + 1 < len(seq_positions) else len(vals)
        after_seg = vals[seq_pos + 1 : next_seq]

        item = _blank_item(vals[seq_pos], ship_date, customer)
        item["description"] = " ".join(desc_parts).strip()
        post_vals = []
        last_consumed = seq_pos + 1
        got_item_no = False

        for j, v in enumerate(after_seg):
            abs_idx = seq_pos + 1 + j
            if _is_page_fraction(v) or _is_seq(v):
                continue
            if v in _FOOTER_STOP:
                break
            if _is_chinese(v):
                if len(v) >= 2 and not item["name"] and v not in _FOOTER_STOP:
                    item["name"] = v
                continue
            if _is_qty(v) and item["quantity"] == 0:
                item["quantity"] = _clean_qty(v)
                last_consumed = abs_idx + 1
            elif not got_item_no and _is_part_no(v):
                item["item_no"] = v
                got_item_no = True
                last_consumed = abs_idx + 1
            elif got_item_no:
                post_vals.append(v)
                last_consumed = abs_idx + 1
            elif _is_spec(v) and not _is_part_no(v):
                break

        _classify_post_item(item, post_vals)
        consumed_up_to = last_consumed
        items.append(item)

    return items


def _parse_format_c(vals, ship_date, customer):
    """格式C：項次 → 數量 → 單位 → 料號 → 客戶料號 → 品名 → 規格（可能因換頁而重複印一次）"""
    seq_positions = [i for i, v in enumerate(vals) if _is_seq(v)]
    items = []

    for si, seq_pos in enumerate(seq_positions):
        end = seq_positions[si + 1] if si + 1 < len(seq_positions) else len(vals)
        seg = vals[seq_pos + 1 : end]

        item = _blank_item(vals[seq_pos], ship_date, customer)
        desc_parts, desc_seen, post_vals = [], set(), []
        state = "qty"

        for v in seg:
            if _is_page_fraction(v) or _is_seq(v):
                continue
            if v in _FOOTER_STOP:
                break
            if _is_header_noise(v):
                continue  # 換頁重印的表頭/簽收區文字，跳過但繼續找真正的品名

            if state == "qty":
                if _is_qty(v):
                    item["quantity"] = _clean_qty(v)
                    state = "unit"

            elif state == "unit":
                if _is_chinese(v) and len(v) <= 2:
                    state = "item_no"   # 跳過單位文字，固定使用 PCS
                elif _is_part_no(v):
                    item["item_no"] = v
                    state = "post_item"

            elif state == "item_no":
                if _is_part_no(v):
                    item["item_no"] = v
                    state = "post_item"

            elif state == "post_item":
                if _is_chinese(v):
                    if not item["name"]:
                        item["name"] = v
                        state = "spec"
                elif _is_part_no(v) and not item["remark"]:
                    item["remark"] = v
                else:
                    post_vals.append(v)

            elif state == "spec":
                if _is_chinese(v):
                    if v == item["name"]:
                        break  # 偵測到重複印的第二份，停止避免重複累加
                    elif len(v) <= 2 and v not in desc_seen:
                        desc_parts.append(v)
                        desc_seen.add(v)
                elif _is_spec(v):
                    if v not in desc_seen:
                        desc_parts.append(v)
                        desc_seen.add(v)
                elif _is_part_no(v):
                    post_vals.append(v)

        _classify_post_item(item, post_vals)
        item["description"] = " ".join(desc_parts).strip()
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
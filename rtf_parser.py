"""
銷貨單 RTF 解析模組

RTF 文件使用絕對定位文字框（\\shp），每個欄位是獨立的文字框。
欄位邊界由標題列（品名/規格、數量）的 X 座標決定：
  seq    = 序號欄（最左）
  item_no= 本公司料號（seq 與 品名 之間）
  name   = 品名（品名/規格欄位的 CJK 前綴）
  spec   = 規格（品名/規格欄位中非 CJK 部分，X 在 name_col_x ~ qty_col_x 之間）
  qty    = 數量（qty_col_x ± BAND）
  post   = 單位/備注/批號（qty_col_x 之後）

批號 vs 客戶料號分類規則（兩階段）：
  - 若該品項含 12 位純數字批號 → R/P+數字視為客戶料號，12位=批號
  - 若無 12 位純數字批號 → R/P+數字視為批號，其他料號=客戶料號
"""
import re
from pathlib import Path

_FOOTER_STOP = {"以下空白"}

# 簽收區的前綴文字 — 出現就代表進入簽收區，應停止收集品項資料
_SIGNING_PREFIXES = ("業務：", "審核：", "經辦人：", "客戶簽收：")

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


def _extract_raw_matches(text: str) -> list[tuple[int, str]]:
    """Extract all (pos, value) from \\loch and \\hich sequences."""
    matches = []
    for m in re.finditer(r"\\loch\\f18 ([^\r\n\\}]+)", text):
        v = m.group(1).strip()
        if v and v not in (":", " ", "  "):
            matches.append((m.start(), v))
    for m in re.compile(r"\\hich\\af1\\dbch\\f18 ((?:\\'[0-9a-fA-F]{2}[\r\n\s]*)+)").finditer(text):
        try:
            hb = bytes(int(x, 16) for x in re.findall(r"'([0-9a-fA-F]{2})", m.group(1)))
            decoded = hb.decode("cp950", errors="replace").strip()
            if decoded:
                matches.append((m.start(), decoded))
        except Exception:
            pass
    return matches


def extract_field_values(rtf_bytes: bytes) -> list[str]:
    """
    只從 RTF 表格欄位（\\trowd…\\row 區塊）提取文字值。
    文字框、頁首、簽收區等不在 table 裡的內容自動排除。
    若偵測不到任何 table，退回全文提取（相容性保底）。
    """
    text = rtf_bytes.decode("cp950", errors="replace")

    # 找出所有 \trowd..\row 區塊（table 列）
    row_spans: list[tuple[int, int]] = []
    for m in re.finditer(r'\\trowd\b', text):
        end_m = re.search(r'\\row\b', text[m.start():])
        if end_m:
            row_spans.append((m.start(), m.start() + end_m.end()))
    row_spans.sort()

    matches = _extract_raw_matches(text)

    if row_spans:
        # 二元搜尋，只保留在 table 列內的 matches
        def _in_table(pos: int) -> bool:
            lo, hi = 0, len(row_spans) - 1
            while lo <= hi:
                mid = (lo + hi) // 2
                s, e = row_spans[mid]
                if s <= pos < e:
                    return True
                elif pos < s:
                    hi = mid - 1
                else:
                    lo = mid + 1
            return False

        filtered = [(p, v) for p, v in matches if _in_table(p)]
        # 只有「全文有重複序號」（多頁文件換頁重印）時才啟用 table-only 過濾。
        # 其他情況（單頁、格式B）退回全文，避免漏掉品名/料號等在 table 外的值。
        all_seqs     = [v for _, v in matches  if re.match(r'^0[0-9]{3}$', v)]
        filt_seqs    = [v for _, v in filtered if re.match(r'^0[0-9]{3}$', v)]
        has_dup      = len(all_seqs) > len(set(all_seqs))
        seq_ok       = bool(filt_seqs) and set(filt_seqs) == set(all_seqs)
        # 只要 table 涵蓋所有序號就用 table-only（去掉文字框/頁眉等雜訊）
        if filtered and seq_ok:
            matches = filtered

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

def _is_invoice_no(v):
    """台灣發票號碼：2大寫英文 + 8數字"""
    return bool(re.match(r'^[A-Z]{2}\d{8}$', v))

def _is_tax_id(v):
    if re.match(r'^\d{8}$', v): return True
    # 注意：[A-Z]{2}\d{8} 是發票號碼，不是統編，已獨立用 _is_invoice_no 判斷
    if re.match(r'^\d{3}-\d{9,}$', v): return True
    if re.match(r'^[A-Z]{3,5}-\d{8,}$', v): return True
    return False

_CJK_PAT = re.compile(r'[一-鿿㐀-䶿豈-﫿]')

def _is_chinese(v):
    return bool(_CJK_PAT.search(v))

def _split_cjk_prefix(v: str) -> tuple[str, str]:
    """('頂針Φ') → ('頂針', 'Φ')  ── 返回純 CJK 前綴與剩餘部分。"""
    i = 0
    while i < len(v) and _CJK_PAT.match(v[i]):
        i += 1
    return v[:i], v[i:]

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
# 位置解析（Shape-based）
# ══════════════════════════════════════════════════════════════════

def _extract_shapes(text: str) -> list[dict]:
    """從 RTF 提取所有 \\shp 定位文字框（排除 \\shprslt 副本）。"""
    loch_pat = re.compile(r'\\loch\\f18 ([^\r\n\\}]+)')
    hich_pat = re.compile(r"\\hich\\af1\\dbch\\f18 ((?:\\'[0-9a-fA-F]{2}[\r\n\s]*)+)")

    def _vals(content: str) -> list[str]:
        raw: list[tuple[int, str]] = []
        for m in loch_pat.finditer(content):
            v = m.group(1).strip()
            if v and v not in (':', ' ', '  '):
                raw.append((m.start(), v))
        for m in hich_pat.finditer(content):
            try:
                hb = bytes(int(x, 16) for x in re.findall(r"'([0-9a-fA-F]{2})", m.group(1)))
                decoded = hb.decode('cp950', errors='replace').strip()
                if decoded:
                    raw.append((m.start(), decoded))
            except Exception:
                pass
        raw.sort(key=lambda x: x[0])
        result, prev = [], None
        for _, v in raw:
            if v != prev:
                result.append(v)
                prev = v
        return result

    shp_pat = re.compile(
        r'\\shp\{\\[*]\\shpinst\\shpleft(-?\d+)\\shptop(-?\d+)\\shpright(-?\d+)\\shpbottom(-?\d+)'
        r'.*?\\shptxt\s*(.*?)\}}}',
        re.DOTALL,
    )
    shapes = []
    for m in shp_pat.finditer(text):
        vals = _vals(m.group(5))
        if vals:
            shapes.append({
                'pos':  m.start(),
                'left': int(m.group(1)),
                'top':  int(m.group(2)),
                'vals': vals,
            })
    return shapes


def _find_col_x_from_header(shapes: list[dict]) -> dict[str, int]:
    """
    從所有 shapes（過濾噪音前）找標題列的欄位 X 座標。
    回傳 {'name': x, 'qty': x}（找到幾個回傳幾個）。
    """
    result: dict[str, int] = {}
    for s in shapes:
        for v in s['vals']:
            if '品名' in v and 'name' not in result:
                result['name'] = s['left']
            if v == '數量' and 'qty' not in result:
                result['qty'] = s['left']
    return result


def _parse_items_by_position(rtf_bytes: bytes, ship_date: str, customer: str) -> list[dict]:
    """
    按文字框的 X/Y 座標解析品項。
    主要路徑：先從標題列取得「品名/規格」和「數量」的 X 座標，以此為邊界，
      將邊界內所有文字框收入品名+規格（避免規格文字因 X 偏移被漏掉）。
    退路：找不到標題時退回原有 X bucket rank 方法。
    """
    text = rtf_bytes.decode('cp950', errors='replace')
    shapes = _extract_shapes(text)

    # ── 先從標題找欄位邊界（過濾噪音前）────────────────────────────
    col_x      = _find_col_x_from_header(shapes)
    name_col_x = col_x.get('name')
    qty_col_x  = col_x.get('qty')
    ALIGN_TOL  = 80    # twips，欄位邊界容忍（~1.4mm）
    use_header = (name_col_x is not None and qty_col_x is not None
                  and qty_col_x > name_col_x + ALIGN_TOL * 2)

    # ── 過濾純噪音文字框 ─────────────────────────────────────────────
    def _noise(vals: list[str]) -> bool:
        return all(_is_header_noise(v) or v in _FOOTER_STOP for v in vals)
    shapes = [s for s in shapes if not _noise(s['vals'])]

    Y_TOL = 60

    # ── 按 Y 分組 ────────────────────────────────────────────────────
    y_buckets: dict[int, list[dict]] = {}
    for s in shapes:
        key = next((k for k in y_buckets if abs(k - s['top']) <= Y_TOL), None)
        if key is None:
            key = s['top']
            y_buckets[key] = []
        y_buckets[key].append(s)

    items: list[dict] = []
    seen_seqs: set[str] = set()

    for top in sorted(y_buckets):
        row_shapes = y_buckets[top]

        # ── 找序號 shapes（排序依 byte 位置 = 頁次順序）────────────
        seq_shapes = sorted(
            [s for s in row_shapes if any(_is_seq(v) for v in s['vals'])],
            key=lambda s: s['pos'],
        )
        if not seq_shapes:
            continue

        n_items    = len(seq_shapes)
        seq_left_x = min(s['left'] for s in seq_shapes)

        if use_header:
            # 每頁的 byte 位置上界
            page_ends = [
                seq_shapes[i + 1]['pos'] if i + 1 < n_items else float('inf')
                for i in range(n_items)
            ]

            def _page_shapes(band: list[dict], pg: int) -> list[dict]:
                e_pos = page_ends[pg]
                s_pos = seq_shapes[pg]['pos']
                sel = [s for s in band
                       if (pg == 0 and s['pos'] < e_pos)
                       or (pg > 0 and s_pos <= s['pos'] < e_pos)]
                return sorted(sel, key=lambda s: s['left'])

            # seq 到 qty 之間（含 qty ± ALIGN_TOL）→ 料號 + 品名 + 規格 + 數量
            pre_post_band = [s for s in row_shapes
                             if seq_left_x < s['left'] <= qty_col_x + ALIGN_TOL]
            # qty 之後 → 單位 / 備注 / 批號
            post_band = [s for s in row_shapes if s['left'] > qty_col_x + ALIGN_TOL]

        else:
            # Fallback：X bucket rank（原方法）
            X_TOL = 30
            x_buckets: dict[int, list[dict]] = {}
            for s in row_shapes:
                key = next((k for k in x_buckets if abs(k - s['left']) <= X_TOL), None)
                if key is None:
                    key = s['left']
                    x_buckets[key] = []
                x_buckets[key].append(s)
            for k in x_buckets:
                x_buckets[k].sort(key=lambda s: s['pos'])
            sorted_x = sorted(x_buckets)

        for item_idx in range(n_items):
            seq = next((v for v in seq_shapes[item_idx]['vals'] if _is_seq(v)), None)
            if not seq or seq in seen_seqs:
                continue
            seen_seqs.add(seq)

            item = _blank_item(seq, ship_date, customer)

            if use_header:
                pre_post_pg = _page_shapes(pre_post_band, item_idx)
                post_pg     = _page_shapes(post_band, item_idx)

                ns:        list[str] = []
                post_vals: list[str] = []
                item_no_found = False
                qty_found     = False

                # 依 X 由左到右掃：第一個料號值 → item_no；數量值 → qty；其餘 → ns
                # 注意：這裡不過濾短 CJK（避免品名片段如「電木」「度」被誤刪）
                for s in pre_post_pg:   # _page_shapes 已按 left 排序
                    for v in s['vals']:
                        if not item_no_found and _is_part_no(v) and not _is_spec(v):
                            item['item_no'] = v
                            item_no_found = True
                        elif _is_qty(v) and not qty_found:
                            item['quantity'] = _clean_qty(v)
                            qty_found = True
                        else:
                            ns.append(v)

                # 後置欄位：批號 / 備注 / 單位（短中文跳過）
                for s in post_pg:
                    for v in s['vals']:
                        if _is_chinese(v) and len(v) <= 3:
                            continue
                        post_vals.append(v)

                _classify_post_item(item, post_vals)

                # 後置中不是批號/料號的值（如規格溢位）→ 補入規格
                for v in post_vals:
                    if not _is_12digit_lot(v) and not _is_rp_lot(v) and not _is_part_no(v):
                        ns.append(v)

            else:
                def gcv(col_rank: int, _ii: int = item_idx) -> list[str]:
                    if col_rank >= len(sorted_x):
                        return []
                    bucket = x_buckets[sorted_x[col_rank]]
                    return bucket[_ii]['vals'] if _ii < len(bucket) else []

                item['item_no'] = gcv(1)[0] if gcv(1) else ''
                ns       = gcv(2)
                qty_vals = gcv(3)
                post_vals = [v for v in gcv(3) if not _is_qty(v)]
                for ci in range(4, len(sorted_x)):
                    for v in gcv(ci):
                        if _is_chinese(v) and len(v) <= 3:
                            continue
                        post_vals.append(v)
                _classify_post_item(item, post_vals)
                item['quantity'] = _clean_qty(next((v for v in qty_vals if _is_qty(v)), '')) \
                    if any(_is_qty(v) for v in qty_vals) else 0

            # ── 品名 vs 規格分割 ─────────────────────────────────────
            name_parts: list[str] = []
            desc_vals:  list[str] = []
            _name_done = False
            for v in ns:
                if _name_done:
                    desc_vals.append(v)
                    continue
                cjk, rest = _split_cjk_prefix(v)
                if cjk:
                    name_parts.append(cjk)
                    if rest:
                        desc_vals.append(rest)
                        _name_done = True
                else:
                    _name_done = True
                    desc_vals.append(v)
            item['name']        = ''.join(name_parts)
            item['description'] = ' '.join(desc_vals).strip()

            items.append(item)

    items.sort(key=lambda x: x['seq'])
    return items


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
        "invoice_no":        "",
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
    _pending_tax_id   = False   # 剛看到「統一編號」標籤，等下一個值
    _pending_inv_no   = False   # 剛看到「發票號碼」標籤，等下一個值
    for v in values[:header_end]:
        if v in header_seen:
            continue
        header_seen.add(v)

        # 標籤追蹤（context-aware）
        if v == "統一編號":
            _pending_tax_id, _pending_inv_no = True, False
            continue
        if v == "發票號碼":
            _pending_inv_no, _pending_tax_id = True, False
            continue

        # 標籤後的值對應
        if _pending_tax_id and re.match(r'^\d{8}$', v) and not result["buyer_tax_id"]:
            result["buyer_tax_id"] = v
            result["buyer_id_raw"] = v
            _pending_tax_id = False
            continue
        if _pending_inv_no and _is_invoice_no(v) and not result["invoice_no"]:
            result["invoice_no"] = v
            _pending_inv_no = False
            continue

        # 其他欄位（同時重置 pending 旗標，表示找到的不是期望值）
        _pending_tax_id = _pending_inv_no = False

        if re.match(r'^C\d{4}(-\d+)?$', v):
            result["customer_code"] = v
        elif re.match(r'^#\d+$', v):
            result["contact"] = v
        elif _is_date(v):
            result["order_date"] = v.replace("/", "")
        elif _is_12digit_lot(v):
            if not result["order_no"]:
                result["order_no"] = v
            elif not result["customer_order_no"] and v != result["order_no"]:
                result["customer_order_no"] = v
        elif _is_invoice_no(v) and not result["invoice_no"]:
            result["invoice_no"] = v   # fallback：不需標籤也捕捉
        elif _is_tax_id(v):
            if v not in tax_ids:
                tax_ids.append(v)
        elif _is_phone(v):
            if v not in phones:
                phones.append(v)

    result["customer_name"] = _extract_customer_name_from_values(values[:header_end])

    # 統編：受票人（客戶）統編從 tax_ids 取（已優先從「統一編號」label 取）
    for tid in tax_ids:
        if not result["buyer_tax_id"]:
            result["buyer_tax_id"] = tid
            result["buyer_id_raw"] = tid

    for i, p in enumerate(phones):
        if i == 0:   result["phone"]  = p
        elif i == 1: result["fax"]    = p
        elif i == 2: result["mobile"] = p

    if not result["order_date"] and result["order_no"]:
        result["order_date"] = result["order_no"][:8]

    # 找出簽收區人名，排除出品項（兩種情況：獨立 label、合併字串 "業務：許雲雀"）
    _person_skip: set[str] = set()
    for _pi, _pv in enumerate(values):
        # 情況1："業務：" 獨立 value，下一個 value 是姓名
        if _pv in set(_SIGNING_PREFIXES) and _pi + 1 < len(values):
            _nx = values[_pi + 1]
            if _is_chinese(_nx) and len(_nx) <= 4:
                _person_skip.add(_nx)
        # 情況2："業務：許雲雀" 合併成單一 value
        for _pfx in _SIGNING_PREFIXES:
            if _pv.startswith(_pfx) and len(_pv) > len(_pfx):
                _name = _pv[len(_pfx):].strip()
                if _is_chinese(_name) and 1 <= len(_name) <= 4:
                    _person_skip.add(_name)
                break

    result["items"] = _parse_items_by_position(raw, result["order_date"], customer_from_filename)

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
            # 頁尾標記或簽收區（業務：/審核：等）→ 停止收集品項資料
            if v in _FOOTER_STOP or any(v.startswith(p) for p in _SIGNING_PREFIXES):
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
                    if item["item_no"]:
                        # 第二個料號 → 歸入後置（客戶料號/批號）
                        post_vals.append(v)
                        if item["name"]:
                            state = "got_item_no"
                    else:
                        item["item_no"] = v
                        # 品名已出現才切換狀態；品名尚未出現則留在 spec 繼續收集
                        if item["name"]:
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
                    if item["item_no"]:
                        post_vals.append(v)
                        state = "got_item_no"
                    else:
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


def _parse_format_b(vals, ship_date, customer, pre_vals=None):
    """格式B：規格 → 項次 → 數量 → 料號 → 客戶料號"""
    seq_positions = [i for i, v in enumerate(vals) if _is_seq(v)]
    items = []
    consumed_up_to = 0

    for si, seq_pos in enumerate(seq_positions):
        spec_seg = vals[consumed_up_to : seq_pos]
        # 第一個品項：規格在第一個序號「之前」，需從 pre_vals 補入
        if si == 0 and pre_vals:
            spec_seg = list(pre_vals) + list(spec_seg)
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
            if v in _FOOTER_STOP or any(v.startswith(p) for p in _SIGNING_PREFIXES):
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


def debug_shapes(file_path) -> list[dict]:
    """回傳所有 shp 文字框的 left/top/vals，供除錯用。"""
    text = Path(file_path).read_bytes().decode('cp950', errors='replace')
    shapes = _extract_shapes(text)
    return [{'left': s['left'], 'top': s['top'], 'vals': s['vals']} for s in shapes]


def parse_multiple_rtf(file_list) -> list[dict]:
    results = []
    for f in file_list:
        try:
            results.append(parse_sales_order_rtf(f))
        except Exception as e:
            results.append({"filename": str(f), "error": str(e), "items": [], "raw_values": []})
    return results
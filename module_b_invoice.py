"""
模組 B：月結（驗收資訊）→ 電子發票

驗收資訊 xlsx 彙整成一張發票，跟一般銷貨單一樣改用 einvoice_submitter.py
逐項填進 e-invoice.com.tw 網站、由人工按下開立（由網站配發發票號碼），
不再產生舊版 V1.6 格式的 .xls 整批上傳檔。
"""
from io import BytesIO

SELLER_TAX_ID = "24405403"

# ── 月結驗收資訊 ─────────────────────────────────────────────────


_ACCEPTANCE_COLS = {
    "出貨單號": 1, "單號": 2, "品號": 3, "品名": 4, "規格": 5,
    "出貨數量": 9, "單價": 12, "幣別": 13, "金額(未稅)": 15,
}


def parse_acceptance_excel(content: bytes) -> list[dict]:
    """
    解析「驗收資訊」xlsx（月結用），回傳每行 dict：
      order_no, line_no, part_no, name, spec, qty, unit_price, amount, currency
    """
    import openpyxl as _xl
    wb = _xl.load_workbook(BytesIO(content), data_only=True)
    ws = wb.active

    def _s(r, c):
        return str(ws.cell(r, c).value or "").strip()

    def _n(r, c):
        try:
            return float(ws.cell(r, c).value or 0)
        except (TypeError, ValueError):
            return 0.0

    # 自動偵測欄位位置（比對第一列標頭）
    header_map = {}
    for c in range(1, ws.max_column + 1):
        h = str(ws.cell(1, c).value or "").strip()
        if h in _ACCEPTANCE_COLS:
            header_map[h] = c

    col = lambda key: header_map.get(key, _ACCEPTANCE_COLS[key])

    result = []
    for r in range(2, ws.max_row + 1):
        order_no = _s(r, col("出貨單號"))
        if not order_no:
            continue
        result.append({
            "order_no":   order_no,
            "line_no":    _s(r, col("單號")),
            "part_no":    _s(r, col("品號")),
            "name":       _s(r, col("品名")),
            "spec":       _s(r, col("規格")),
            "qty":        _n(r, col("出貨數量")),
            "unit_price": _n(r, col("單價")),
            "amount":     _n(r, col("金額(未稅)")),
            "currency":   _s(r, col("幣別")) or "NTD",
        })
    return result


def acceptance_rows_to_order(
    rows: list[dict],
    order_no: str,
    customer_name: str,
    buyer_tax_id: str,
) -> dict:
    """
    把 parse_acceptance_excel 的 rows 彙整成一張「訂單」，
    餵給 einvoice_submitter.fill_one_order 逐項填進網站表單。
    每個品項各自帶自己那一行的出貨單號／單號，分別對應相關號碼一／二，
    跟舊版 V1.6 上傳檔的欄位對應一致。
    """
    items = []
    for row in rows:
        items.append({
            "item_no": row.get("part_no", ""),
            "name": row.get("name", ""),
            "description": row.get("spec", ""),
            "quantity": row.get("qty", 0),
            "unit_price": row.get("unit_price", 0),
            "unit": "PCS",
            "customer_order_no": row.get("order_no", ""),  # 相關號碼一：該行自己的出貨單號
            "remark": row.get("line_no", ""),               # 相關號碼二：該行自己的單號
        })
    return {
        "order_no": order_no,
        "customer_name": customer_name,
        "buyer_tax_id": buyer_tax_id,
        "items": items,
    }

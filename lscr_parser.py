"""
LSCR 出貨明細確認單 Excel 解析器
工作表 'list' → order/item 格式（與 RTF 解析器相容）
"""
import re
import openpyxl


def _v(ws, row, col) -> str:
    v = ws.cell(row=row, column=col).value
    return str(v).strip() if v is not None else ""


def parse_lscr_excel_wb(wb: openpyxl.Workbook) -> list[dict]:
    """
    從已開啟的 Workbook 解析 'list' 工作表。
    回傳 list[order_dict]，每個 PO NO 一筆。
    """
    if "list" not in wb.sheetnames:
        raise ValueError("找不到 'list' 工作表")
    ws = wb["list"]

    # ── 標頭資訊（前 5 行）──────────────────────────────────────────
    customer = ""
    ship_date = ""
    for r in range(1, 6):
        for c in range(1, 20):
            v = _v(ws, r, c)
            if re.search(r'客戶[：:]', v):
                customer = re.sub(r'^.*客戶[：:]\s*', '', v).strip()
            if re.search(r'出貨日期[：:]', v):
                d = re.sub(r'^.*出貨日期[：:]\s*', '', v).strip()
                ship_date = re.sub(r'[-/年月]', '', d).strip('日').replace(' ', '')

    # ── 資料行（跳過標頭，偵測有 Item 欄位的行）────────────────────
    # 欄位：C5=PO NO, C6=LOT NO, C7=Item(料號), C8=成品圖號,
    #       C9=品名, C10=規格, C11=總數量, C12=大包裝qty, C13=大包裝unit,
    #       C14=小包裝qty, C15=小包裝unit
    orders_map: dict[str, dict] = {}
    max_row = ws.max_row

    for r in range(6, max_row + 1):
        item_col = _v(ws, r, 7)   # Item (料號)
        qty_col  = _v(ws, r, 11)  # 總數量
        if not item_col or not qty_col:
            continue
        # 跳過 header-like 列
        if item_col.lower() in ("item", "料號", "品名", "no."):
            continue

        po_no    = _v(ws, r, 5) or "N/A"
        lot_no   = _v(ws, r, 6)
        remark   = _v(ws, r, 8)   # 成品圖號
        name     = _v(ws, r, 9)
        desc     = _v(ws, r, 10)
        pkg_qty  = _v(ws, r, 14)   # 小包裝 qty
        pkg_unit = _v(ws, r, 15)   # 小包裝 unit

        qty  = pkg_qty  if pkg_qty  else qty_col
        unit = pkg_unit if pkg_unit else "PCS"

        item = {
            "item_no":     item_col,
            "name":        name,
            "description": desc,
            "quantity":    qty,
            "unit":        unit,
            "lot_no":      lot_no,
            "remark":      remark,
            "ship_date":   ship_date,
        }

        if po_no not in orders_map:
            orders_map[po_no] = {
                "order_no":          po_no,
                "customer_name":     customer,
                "customer_order_no": po_no,
                "ship_date":         ship_date,
                "items":             [],
            }
        orders_map[po_no]["items"].append(item)

    return list(orders_map.values())


def parse_lscr_excel(path: str) -> list[dict]:
    wb = openpyxl.load_workbook(path, data_only=True)
    return parse_lscr_excel_wb(wb)

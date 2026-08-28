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

    每個 item 含標準欄位外，另有私有欄位：
      _total_qty  : 總出貨數量 (str)
      _large_qty  : 大包裝數量 (str)
      _large_unit : 大包裝單位
      _small_qty  : 小包裝數量 (str)
      _small_unit : 小包裝單位
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

    # ── 資料行
    # 欄位：C1=ID1, C2=ID2,
    #       C5=PO NO, C6=LOT NO, C7=Item(料號), C8=成品圖號,
    #       C9=品名, C10=規格,
    #       C11=總數量
    # C12~C15 有兩種版本的確認單格式，逐行判斷用哪一種（見下方）：
    #   舊版：C12=大包裝qty, C13=大包裝unit, C14=小包裝qty, C15=小包裝unit
    #   新版（C14 沒有值時）：C12=標籤(小包裝)qty, C13=標籤(小包裝)unit，
    #                        大包裝固定＝該行自己的訂單總數量（C11）
    orders_map: dict[str, dict] = {}

    for r in range(6, ws.max_row + 1):
        item_col = _v(ws, r, 7)    # Item (料號)
        total_col = _v(ws, r, 11)  # 總數量
        if not item_col or not total_col:
            continue
        if item_col.lower() in ("item", "料號", "品名", "no."):
            continue

        id1 = _v(ws, r, 1)
        id2 = _v(ws, r, 2)
        po_no = _v(ws, r, 5) or "N/A"
        lot_no = _v(ws, r, 6)
        remark = _v(ws, r, 8)   # 成品圖號
        name = _v(ws, r, 9)
        desc = _v(ws, r, 10)
        col12 = _v(ws, r, 12)
        col13 = _v(ws, r, 13)
        col14 = _v(ws, r, 14)
        col15 = _v(ws, r, 15)

        if col14:
            # 舊版格式：C14 有值，照舊當小包裝欄位
            large_qty = col12
            large_unit = col13 or "PCS"
            small_qty = col14
            small_unit = col15 or "PCS"
        else:
            # 新版格式：沒有獨立的大包裝欄位，C12/C13 是小包裝（固定兩張），
            # 大標籤直接印這一行自己的訂單總數量
            large_qty = total_col
            large_unit = "PCS"
            small_qty = col12
            small_unit = col13 or "PCS"

        # 品名加入 ID1/ID2 尺寸
        if id1 and id2:
            name = f"{name}（ID1={id1}mm*ID2={id2}mm）"

        # 預設用小包裝數量顯示
        qty = small_qty if small_qty else total_col
        unit = small_unit if small_qty else large_unit

        item = {
            "item_no":      item_col,
            "name":         name,
            "description":  desc,
            "quantity":     qty,
            "unit":         unit,
            "lot_no":       lot_no,
            "remark":       remark,
            "ship_date":    ship_date,
            # 私有欄位供展開邏輯使用
            "_total_qty":   total_col,
            "_large_qty":   large_qty or total_col,
            "_large_unit":  large_unit,
            "_small_qty":   small_qty,
            "_small_unit":  small_unit,
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

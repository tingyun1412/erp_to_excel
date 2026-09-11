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
      _total_qty  : 訂單數量/總數量 (str)
      _small_qty  : 標籤（小包裝）數量 (str) —— 標籤實際印的就是這個數字
      _small_unit : 標籤單位，固定 "PCS"
      _pack_qty   : 裝箱數量 (str) —— 純粹紀錄用，不影響標籤內容
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
    #       C11=訂單數量(總數量)，
    #       C12=標籤（小包裝）數量 —— 標籤只印這一欄，
    #       C13=裝箱數量 —— 純粹紀錄用，不管填多少都不影響標籤內容
    #         （之前把 C13 誤當成「單位文字」接到標籤數量後面，
    #           變成裝箱的數字混進標籤顯示的數量，已改成完全不參與計算）。
    #       C14/C15 是同樣「標籤/裝箱」的重複表頭，目前版本的確認單一律空白，不使用。
    orders_map: dict[str, dict] = {}

    for r in range(6, ws.max_row + 1):
        item_col = _v(ws, r, 7)    # Item (料號)
        total_col = _v(ws, r, 11)  # 訂單數量(總數量)
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
        label_qty = _v(ws, r, 12)  # 標籤（小包裝）數量
        pack_qty  = _v(ws, r, 13)  # 裝箱數量：僅供參考，不影響標籤

        # 品名加入 ID1/ID2 尺寸
        if id1 and id2:
            name = f"{name}（ID1={id1}mm*ID2={id2}mm）"

        qty = label_qty if label_qty else total_col

        item = {
            "item_no":      item_col,
            "name":         name,
            "description":  desc,
            "quantity":     qty,
            "unit":         "PCS",
            "lot_no":       lot_no,
            "remark":       remark,
            "ship_date":    ship_date,
            # 私有欄位供 write_lscr_labels／預覽表使用
            "_total_qty":   total_col,
            "_small_qty":   qty,
            "_small_unit":  "PCS",
            "_pack_qty":    pack_qty,
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

"""
模組 B：銷貨單 → 電子發票上傳 Excel
依照 e-invoice.com.tw 的 V1.6 格式產生可直接上傳的 Excel
"""
from datetime import datetime
from io import BytesIO

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side


# ── 樣式 ──────────────────────────────────────────────────────────
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(color="FFFFFF", bold=True)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center")


def _tw_date(date_str: str) -> str:
    """
    YYYYMMDD → 民國年 YYYMMDD（7碼）
    例：20260617 → 1150617
    """
    if len(date_str) == 8:
        try:
            y = int(date_str[:4]) - 1911
            return f"{y}{date_str[4:]}"
        except ValueError:
            pass
    today = datetime.today()
    return f"{today.year - 1911}{today.month:02d}{today.day:02d}"


def _now_time() -> str:
    """HHMMSS"""
    return datetime.now().strftime("%H%M%S")


def generate_invoice_excel(
    orders: list[dict],
    seller_tax_id: str = "",
    invoice_prefix: str = "AA",
    start_number: int = 1,
) -> BytesIO:
    """
    接收多張銷貨單資料，產生電子發票上傳 Excel（主檔 + 明細）

    orders: list of dict，每筆一張銷貨單，欄位：
        order_no, order_date, seller_tax_id, buyer_tax_id,
        items: list of {description, quantity, unit, unit_price, amount, remark}

    seller_tax_id: 賣方統編（若訂單內有就優先用訂單的）
    invoice_prefix: 發票字軌（2碼英文，預設 AA）
    start_number: 起始發票號碼（8位數字）
    """
    wb = openpyxl.Workbook()
    ws_main   = wb.active
    ws_main.title = "發票主檔"
    ws_detail = wb.create_sheet("發票明細")

    # ── 主檔標題 ──────────────────────────────────────────────
    main_headers = [
        "發票號碼(IVNO)", "發票日期(IVDAT)", "發票時間(IVTM)",
        "未稅金額(IVAMT)", "稅率別(TAXRID)", "營業稅額(SALTAXAMT)",
        "發票人統一編號(IVPESRFNO)", "受票人統一編號(TAIVPESRFNO)",
        "款項別(CAID)", "相關號碼(RELNO)", "原幣金額(OCRYAMT)",
        "匯率(EXR)", "幣別(CUCY)", "彙開(GROPMK)", "通關方式(CSTMMK)",
        "買方聯絡人(BUYRCTM)", "買方聯絡人部門(BUYRCTMDP)",
        "買受人電子郵件(CUEMAIL)", "總備註(COMT5)",
        "發票開立自動通知(OPNAUTNTI)", "作廢發票自動通知(CANCELAUTNTI)",
        "零稅率原因(ZEROTAXRATEREASON)",
    ]
    detail_headers = [
        "發票號碼(IVNO)", "項次(IT)", "品名(DSR)", "品名2(DSR2)",
        "數量(QTY1)", "單位(UN1)", "單價(UP)", "金額(AMT)",
        "相關號碼一(RELNO1)", "相關號碼二(RELNO2)",
    ]

    for col, h in enumerate(main_headers, 1):
        cell = ws_main.cell(row=1, column=col, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = BORDER
        ws_main.column_dimensions[cell.column_letter].width = max(len(h) * 1.5, 12)

    for col, h in enumerate(detail_headers, 1):
        cell = ws_detail.cell(row=1, column=col, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = BORDER
        ws_detail.column_dimensions[cell.column_letter].width = max(len(h) * 1.5, 12)

    # ── 填入資料 ──────────────────────────────────────────────
    main_row   = 2
    detail_row = 2
    inv_num    = start_number

    for order in orders:
        items = order.get("items", [])
        if not items:
            continue

        # 計算金額
        total_amount = sum(
            (it.get("amount") or (it.get("quantity", 0) * it.get("unit_price", 0)))
            for it in items
        )
        tax_amount = round(total_amount * 0.05)

        inv_no  = f"{invoice_prefix}{inv_num:08d}"
        inv_date = _tw_date(order.get("order_date", ""))
        inv_time = _now_time()
        sid = order.get("seller_tax_id") or seller_tax_id
        bid = order.get("buyer_tax_id", "")
        relno = order.get("order_no", "")

        main_values = [
            inv_no, inv_date, inv_time,
            total_amount, 1, tax_amount,
            sid, bid,
            "Z", relno, "", 1, "TWD",
            "", "", "", "", "", "",
            "", "", "",
        ]
        for col, val in enumerate(main_values, 1):
            cell = ws_main.cell(row=main_row, column=col, value=val)
            cell.border = BORDER
            cell.alignment = CENTER

        # 明細
        for idx, item in enumerate(items, 1):
            qty       = item.get("quantity", 1)
            unit      = item.get("unit", "PC")
            up        = item.get("unit_price", 0)
            amt       = item.get("amount") or (qty * up)
            desc      = item.get("description", item.get("item_no", ""))
            desc2     = item.get("remark", "")  # 備註/客戶貨號放品名2
            relno1    = item.get("item_no", "")

            detail_values = [
                inv_no, idx, desc, desc2 if desc2 else None,
                qty, unit, up, amt,
                relno1, "",
            ]
            for col, val in enumerate(detail_values, 1):
                cell = ws_detail.cell(row=detail_row, column=col, value=val)
                cell.border = BORDER
                cell.alignment = CENTER
            detail_row += 1

        main_row += 1
        inv_num  += 1

    # 凍結首列
    ws_main.freeze_panes   = "A2"
    ws_detail.freeze_panes = "A2"

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

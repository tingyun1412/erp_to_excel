"""
模組 C：銷貨單 → 標籤 Excel
每張銷貨單的每個品項產生一張標籤
支援多種標籤格式（透過 LABEL_TEMPLATES 字典擴充）

新增標籤格式方法（給開發者）：
    1. 在 LABEL_TEMPLATES 加一個 key（格式名稱）
    2. value 是一個 function(ws, row_start, item, order) → 回傳下一個可用 row
    3. 重啟 app 即可在下拉選單看到新格式

新增標籤格式方法（給非技術使用者，未來可考慮做 UI）：
    目前需請技術人員協助，後續版本可加入「自訂欄位對應」UI
"""
from io import BytesIO

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ── 共用樣式 ──────────────────────────────────────────────────────
THIN  = Side(style="thin",   color="000000")
THICK = Side(style="medium", color="000000")
FULL_BORDER   = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
OUTER_BORDER  = Border(left=THICK, right=THICK, top=THICK, bottom=THICK)
CENTER        = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT          = Alignment(horizontal="left",   vertical="center", wrap_text=True)
TITLE_FONT    = Font(bold=True, size=14)
LABEL_FONT    = Font(bold=True, size=9)
VALUE_FONT    = Font(size=10)
HEADER_FILL   = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT   = Font(color="FFFFFF", bold=True, size=9)
ALT_FILL      = PatternFill("solid", fgColor="EBF3FB")


def _set(ws, row, col, value, font=None, fill=None, alignment=None, border=None):
    cell = ws.cell(row=row, column=col, value=value)
    if font:      cell.font      = font
    if fill:      cell.fill      = fill
    if alignment: cell.alignment = alignment
    if border:    cell.border    = border
    return cell


def _merge_set(ws, r1, c1, r2, c2, value, **kwargs):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    return _set(ws, r1, c1, value, **kwargs)


# ══════════════════════════════════════════════════════════════════
#  標籤格式定義
#  每個 function 簽名：(ws, row_start, item, order) → next_row (int)
# ══════════════════════════════════════════════════════════════════

def _label_standard(ws, row_start, item, order):
    """
    標準格式：包含銷貨單上所有重要資訊
    大小：約 10 行 × 5 欄
    """
    r = row_start

    # 外框（先設定後面的資料行會蓋上）
    title_val = f"出貨標籤"
    _merge_set(ws, r, 1, r, 5, title_val,
               font=Font(bold=True, size=13, color="FFFFFF"),
               fill=HEADER_FILL,
               alignment=CENTER)
    r += 1

    # 銷貨單號
    _set(ws, r, 1, "銷貨單號", font=LABEL_FONT, fill=ALT_FILL, alignment=CENTER, border=FULL_BORDER)
    _merge_set(ws, r, 2, r, 5, order.get("order_no",""),
               font=VALUE_FONT, alignment=LEFT, border=FULL_BORDER)
    r += 1

    # 客戶名稱
    _set(ws, r, 1, "客戶", font=LABEL_FONT, fill=ALT_FILL, alignment=CENTER, border=FULL_BORDER)
    _merge_set(ws, r, 2, r, 5, item.get("customer",""),
               font=VALUE_FONT, alignment=LEFT, border=FULL_BORDER)
    r += 1

    # 料號
    _set(ws, r, 1, "料號", font=LABEL_FONT, fill=ALT_FILL, alignment=CENTER, border=FULL_BORDER)
    _merge_set(ws, r, 2, r, 5, item.get("item_no",""),
               font=Font(bold=True, size=11), alignment=LEFT, border=FULL_BORDER)
    r += 1

    # 品名
    _set(ws, r, 1, "品名", font=LABEL_FONT, fill=ALT_FILL, alignment=CENTER, border=FULL_BORDER)
    _merge_set(ws, r, 2, r, 5, item.get("description",""),
               font=VALUE_FONT, alignment=LEFT, border=FULL_BORDER)
    r += 1

    # 數量 / 單位
    _set(ws, r, 1, "數量", font=LABEL_FONT, fill=ALT_FILL, alignment=CENTER, border=FULL_BORDER)
    _merge_set(ws, r, 2, r, 3, item.get("quantity",""),
               font=Font(bold=True, size=12), alignment=CENTER, border=FULL_BORDER)
    _set(ws, r, 4, "單位", font=LABEL_FONT, fill=ALT_FILL, alignment=CENTER, border=FULL_BORDER)
    _set(ws, r, 5, item.get("unit","PC"),
               font=VALUE_FONT, alignment=CENTER, border=FULL_BORDER)
    r += 1

    # 客戶貨號/備註
    remark = item.get("remark","")
    if remark:
        _set(ws, r, 1, "客戶貨號", font=LABEL_FONT, fill=ALT_FILL, alignment=CENTER, border=FULL_BORDER)
        _merge_set(ws, r, 2, r, 5, remark,
                   font=VALUE_FONT, alignment=LEFT, border=FULL_BORDER)
        r += 1

    # 出貨日期
    _set(ws, r, 1, "出貨日期", font=LABEL_FONT, fill=ALT_FILL, alignment=CENTER, border=FULL_BORDER)
    sd = item.get("ship_date","")
    if len(sd) == 8:
        sd_fmt = f"{sd[:4]}/{sd[4:6]}/{sd[6:8]}"
    else:
        sd_fmt = sd
    _merge_set(ws, r, 2, r, 5, sd_fmt,
               font=VALUE_FONT, alignment=LEFT, border=FULL_BORDER)
    r += 1

    # 客戶訂單號
    cust_order = order.get("customer_order_no","")
    if cust_order:
        _set(ws, r, 1, "客戶訂單", font=LABEL_FONT, fill=ALT_FILL, alignment=CENTER, border=FULL_BORDER)
        _merge_set(ws, r, 2, r, 5, cust_order,
                   font=VALUE_FONT, alignment=LEFT, border=FULL_BORDER)
        r += 1

    # 空白分隔行
    r += 1
    return r


def _label_simple(ws, row_start, item, order):
    """
    簡易格式：只顯示料號、品名、數量，適合小標籤
    大小：約 5 行 × 5 欄
    """
    r = row_start

    _merge_set(ws, r, 1, r, 5, item.get("item_no",""),
               font=Font(bold=True, size=14), alignment=CENTER,
               fill=PatternFill("solid", fgColor="2E75B6"),
               border=FULL_BORDER)
    ws.cell(r, 1).font = Font(bold=True, size=14, color="FFFFFF")
    r += 1

    _merge_set(ws, r, 1, r+1, 5, item.get("description",""),
               font=Font(size=10), alignment=CENTER, border=FULL_BORDER)
    r += 2

    _merge_set(ws, r, 1, r, 2, "QTY",
               font=Font(bold=True, size=10), alignment=CENTER,
               fill=ALT_FILL, border=FULL_BORDER)
    _merge_set(ws, r, 3, r, 5, item.get("quantity",""),
               font=Font(bold=True, size=12), alignment=CENTER, border=FULL_BORDER)
    r += 1

    r += 1  # 分隔
    return r


def _label_with_barcode_placeholder(ws, row_start, item, order):
    """
    含條碼位置的格式：左側資訊、右側留條碼空間
    （條碼需另外工具產生，這裡只留位置）
    大小：約 8 行 × 8 欄
    """
    r = row_start

    # 標題
    _merge_set(ws, r, 1, r, 8, f"出貨標籤 - {order.get('order_no','')}",
               font=Font(bold=True, size=11, color="FFFFFF"),
               fill=HEADER_FILL, alignment=CENTER)
    r += 1

    # 左側資訊（欄 1-5），右側條碼位置（欄 6-8）
    info = [
        ("料號",   item.get("item_no","")),
        ("品名",   item.get("description","")),
        ("數量",   f'{item.get("quantity","")} {item.get("unit","PC")}'),
        ("客戶",   item.get("customer","")),
        ("出貨日", item.get("ship_date","")),
    ]
    barcode_start = r
    for label, value in info:
        _set(ws, r, 1, label, font=LABEL_FONT, fill=ALT_FILL,
             alignment=CENTER, border=FULL_BORDER)
        _merge_set(ws, r, 2, r, 5, value,
                   font=VALUE_FONT, alignment=LEFT, border=FULL_BORDER)
        r += 1

    # 條碼區塊
    _merge_set(ws, barcode_start, 6, barcode_start + len(info) - 1, 8,
               "[條碼區]",
               font=Font(size=9, color="888888"),
               alignment=CENTER, border=FULL_BORDER)

    r += 1  # 分隔
    return r


# ══════════════════════════════════════════════════════════════════
#  標籤格式登錄表（新增格式只需在這裡加）
# ══════════════════════════════════════════════════════════════════
LABEL_TEMPLATES: dict[str, callable] = {
    "標準格式（含完整資訊）":     _label_standard,
    "簡易格式（料號+品名+數量）": _label_simple,
    "含條碼位置格式":            _label_with_barcode_placeholder,
}


def generate_labels_excel(
    orders: list[dict],
    template_name: str = "標準格式（含完整資訊）",
    labels_per_row: int = 1,
) -> BytesIO:
    """
    產生標籤 Excel
    orders: list of dict
    template_name: LABEL_TEMPLATES 的 key
    labels_per_row: 每行幾個標籤並排（目前實作 1，多欄可擴充）
    """
    template_fn = LABEL_TEMPLATES.get(template_name, _label_standard)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "出貨標籤"

    # 設定欄寬
    col_widths = [12, 8, 8, 8, 8, 8, 8, 8]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    current_row = 1
    for order in orders:
        for item in order.get("items", []):
            current_row = template_fn(ws, current_row, item, order)

    # 凍結首列無意義，標籤不凍結
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def list_templates() -> list[str]:
    """回傳所有可用的標籤格式名稱"""
    return list(LABEL_TEMPLATES.keys())

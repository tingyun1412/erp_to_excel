"""
標籤模板引擎
核心邏輯：
1. 分析使用者上傳的舊標籤 Excel → 找出「一個標籤單元」的結構
2. 使用者標記哪些格子是動態欄位（料號/品名/數量/日期/流水號等）
3. 存模板設定到 Google Sheets
4. 之後輸入銷貨單資料 → 複製模板格式 → 填入新資料 → 產出 Excel
"""
import copy
import json
import re
from datetime import datetime, date
from io import BytesIO

import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import (
    Alignment, Border, Font, PatternFill, Side,
    GradientFill,
)
from openpyxl.utils import get_column_letter, column_index_from_string


# ── 可對應的動態欄位 ──────────────────────────────────────────────
DYNAMIC_FIELDS = {
    "料號":     "item_no",
    "品名":     "description",
    "規格":     "description",   # 有些廠商叫規格
    "數量":     "quantity",
    "出貨日期": "ship_date",
    "客戶料號": "remark",
    "銷貨單號": "order_no",
    "客戶訂單": "customer_order_no",
    "Lot No":  "lot_no",        # 自動流水號
    "流水號":   "lot_no",
    "固定文字": "__fixed__",     # 不變的文字，如公司名稱、TEL
}

FIELD_LABELS = list(DYNAMIC_FIELDS.keys())


def _fmt_date(d: str) -> str:
    """YYYYMMDD → YYYY.MM.DD"""
    if len(d) == 8:
        try:
            return f"{d[:4]}.{d[4:6]}.{d[6:8]}"
        except Exception:
            pass
    return d


def _fmt_lot_no(order_no: str, seq: int, col_idx: int = 0) -> str:
    """
    產生 Lot No / 流水號
    order_no: 銷貨單號（12碼）
    seq: 第幾個標籤（1-based）
    col_idx: 同一列第幾欄（0-based，用於同列多個流水號遞增）
    """
    base = order_no[:12] if order_no else datetime.now().strftime("%Y%m%d%H%M")
    number = (seq - 1) * 2 + col_idx + 1  # 兩欄並排時遞增
    return f"{base[-8:]}{number:04d}"


def _get_cell_value(item: dict, order: dict, field_key: str, lot_seq: int = 1, col_idx: int = 0) -> str:
    """根據欄位 key 從訂單資料取值"""
    if field_key == "__fixed__" or not field_key:
        return None  # 固定文字保持原樣
    if field_key == "lot_no":
        return _fmt_lot_no(order.get("order_no", ""), lot_seq, col_idx)
    if field_key == "sequence":
        return f"{(lot_seq - 1) * 2 + col_idx + 1:04d}"
    if field_key == "ship_date":
        return _fmt_date(item.get("ship_date", ""))
    if field_key == "quantity":
        qty = item.get("quantity", "")
        unit = item.get("unit", "PCS")
        return f"{qty} {unit}" if qty else ""
    
    val = item.get(field_key) or order.get(field_key, "")
    return str(val) if val else ""


# ══════════════════════════════════════════════════════════════════
#  模板分析
# ══════════════════════════════════════════════════════════════════

def analyze_template(wb: openpyxl.Workbook, sheet_name: str) -> dict:
    """
    分析模板 Excel 的結構：
    - 偵測「一個標籤單元」有幾行幾欄
    - 偵測並排幾欄（每個單元佔幾欄，中間有無空欄）
    - 列出所有有值的格子和其值
    
    回傳：
    {
        "sheet_name": str,
        "unit_rows": int,       # 一個標籤幾行
        "columns_per_unit": int,# 一個標籤幾欄
        "gap_cols": int,        # 單元之間空幾欄
        "units_per_row": int,   # 每列幾個並排
        "cells": [              # 單元左上角為 (1,1) 的相對座標
            {"row": 1, "col": 1, "value": "晶晟精密科技股分有限公司", "field": "__fixed__"},
            ...
        ],
        "col_widths": {...},
        "row_heights": {...},
    }
    """
    ws = wb[sheet_name]
    all_values = ws.get_all_values() if hasattr(ws, 'get_all_values') else None
    
    # 用 openpyxl 讀所有有值的格
    cells_with_value = []
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None and str(cell.value).strip():
                cells_with_value.append(cell)
    
    if not cells_with_value:
        return {}
    
    # 找最大行列
    max_row = max(c.row for c in cells_with_value)
    max_col = max(c.column for c in cells_with_value)
    
    # 偵測空欄（用來找並排分隔）
    col_has_value = set(c.column for c in cells_with_value)
    
    # 找分隔空欄（寬度很窄的欄 or 完全沒值的欄）
    separator_cols = []
    for col_idx in range(1, max_col + 1):
        if col_idx not in col_has_value:
            separator_cols.append(col_idx)
        else:
            col_letter = get_column_letter(col_idx)
            width = ws.column_dimensions[col_letter].width if col_letter in ws.column_dimensions else 8
            if width < 5:  # 窄欄視為分隔
                separator_cols.append(col_idx)
    
    # 判斷每個標籤單元的欄範圍
    # 找所有「非分隔欄」的連續區段
    unit_col_groups = []
    current_group = []
    for col_idx in range(1, max_col + 1):
        if col_idx in separator_cols:
            if current_group:
                unit_col_groups.append(current_group)
                current_group = []
        else:
            current_group.append(col_idx)
    if current_group:
        unit_col_groups.append(current_group)
    
    units_per_row = len(unit_col_groups) if unit_col_groups else 1
    
    # 以第一個單元為基準
    first_unit_cols = unit_col_groups[0] if unit_col_groups else list(range(1, max_col + 1))
    columns_per_unit = len(first_unit_cols)
    gap_cols = len(separator_cols) // max(units_per_row - 1, 1) if units_per_row > 1 else 0
    
    # 偵測標籤單元的行數（找重複的 pattern）
    first_col_vals = []
    for row_idx in range(1, max_row + 1):
        cell = ws.cell(row=row_idx, column=first_unit_cols[0])
        first_col_vals.append(str(cell.value) if cell.value else "")
    
    # 偵測是否為「橫向流水號」模式（標題1行+資料行重複）
    # 特徵：第2行和第3行在同一欄的值格式相同（只有流水號遞增）
    header_rows = 0
    is_row_repeat_mode = False
    if max_row >= 3:
        row2_vals = [str(ws.cell(row=2, column=c).value or "") for c in first_unit_cols]
        row3_vals = [str(ws.cell(row=3, column=c).value or "") for c in first_unit_cols]
        # 如果第2行和第3行的非數字部分相同，視為橫向模式
        def strip_nums(s): return "".join(c for c in s if not c.isdigit())
        if strip_nums("".join(row2_vals)) == strip_nums("".join(row3_vals)):
            is_row_repeat_mode = True
            header_rows = 1
            unit_rows = 1
    
    if not is_row_repeat_mode:
        unit_rows = _detect_unit_rows(first_col_vals)
    
    # 擷取格子資訊
    cells_info = []
    # header 行（固定標題）
    for row_idx in range(1, header_rows + 1):
        for col_offset, col_idx in enumerate(first_unit_cols):
            cell = ws.cell(row=row_idx, column=col_idx)
            if cell.value is not None:
                val = str(cell.value).strip()
                field = _guess_field(val)
                cells_info.append({
                    "row":        row_idx,
                    "col":        col_offset + 1,
                    "value":      val,
                    "field":      field,
                    "is_header":  True,
                    "font_bold":  cell.font.bold if cell.font else False,
                    "alignment":  cell.alignment.horizontal if cell.alignment else "left",
                })
    # 資料行（重複單元）
    data_start = header_rows + 1
    for row_idx in range(data_start, data_start + unit_rows):
        for col_offset, col_idx in enumerate(first_unit_cols):
            cell = ws.cell(row=row_idx, column=col_idx)
            if cell.value is not None:
                val = str(cell.value).strip()
                field = _guess_field(val)
                cells_info.append({
                    "row":        row_idx - header_rows,  # 相對於資料區的行號
                    "col":        col_offset + 1,
                    "value":      val,
                    "field":      field,
                    "is_header":  False,
                    "font_bold":  cell.font.bold if cell.font else False,
                    "alignment":  cell.alignment.horizontal if cell.alignment else "left",
                })
    
    # 欄寬和行高
    col_widths = {}
    for col_idx in first_unit_cols:
        letter = get_column_letter(col_idx)
        w = ws.column_dimensions[letter].width if letter in ws.column_dimensions else 10
        col_widths[str(col_idx - first_unit_cols[0] + 1)] = w
    
    row_heights = {}
    for row_idx in range(1, unit_rows + 1):
        h = ws.row_dimensions[row_idx].height if row_idx in ws.row_dimensions else 15
        row_heights[str(row_idx)] = h or 15
    
    return {
        "sheet_name":        sheet_name,
        "unit_rows":         unit_rows,
        "header_rows":       header_rows,
        "is_row_repeat_mode": is_row_repeat_mode,
        "columns_per_unit":  columns_per_unit,
        "gap_cols":          gap_cols,
        "units_per_row":     units_per_row,
        "cells":             cells_info,
        "col_widths":        col_widths,
        "row_heights":       row_heights,
        "max_row":           max_row,
        "max_col":           max_col,
    }


def _detect_unit_rows(col_vals: list[str]) -> int:
    """
    從第一欄的值列表，找標籤重複的周期（單元行數）
    策略：找第一個非空值在後面重複出現的位置
    """
    if not col_vals:
        return 8  # 預設
    
    first_val = next((v for v in col_vals if v), "")
    if not first_val:
        return 8
    
    # 從第二次出現的位置判斷週期
    for i in range(1, len(col_vals)):
        if col_vals[i] == first_val:
            return i
    
    return len(col_vals)


def _guess_field(value: str) -> str:
    """根據格子的值自動猜測是哪個動態欄位"""
    v = value.lower()
    
    # 固定模式
    if "tel" in v or "fax" in v or "made in" in v:
        return "__fixed__"
    
    # 含冒號的標籤格（如 "料號：BSE50053"）→ 整格是動態
    if "料號" in v or "品號" in v:
        return "item_no_inline"  # 整行包含標籤+值
    if "品名" in v:
        return "description_inline"
    if "規格" in v:
        return "description_inline"
    if "數量" in v:
        return "quantity_inline"
    if "出貨日期" in v or "出廠日期" in v:
        return "ship_date_inline"
    if "lot no" in v:
        return "lot_no_inline"
    if "流水碼" in v or "流水號" in v:
        return "sequence"
    if "年月日" in v:
        return "ship_date"
    if "訂單號碼" in v or "訂單號" in v:
        return "customer_order_no_inline"
    if "思達料號" in v or "客戶料號" in v:
        return "remark_inline"
    if "思達品名" in v or "客戶品名" in v:
        return "description_inline"
    
    # 純值格（無冒號）
    if re.match(r'^\d{6,8}$', value):       return "ship_date"
    if re.match(r'^[A-Z]{2,}[0-9A-Z\-]+$', value) and len(value) >= 6:
        return "item_no"
    if re.match(r'^\d{4}$', value):          return "sequence"
    
    return "__fixed__"


# ══════════════════════════════════════════════════════════════════
#  標籤產生
# ══════════════════════════════════════════════════════════════════

def _fill_inline_value(template_str: str, field_key: str, item: dict, order: dict,
                        lot_seq: int = 1, col_idx: int = 0) -> str:
    """
    處理「標籤：值」合在一格的情況
    例：「料號：BAS089109BREB」→ 「料號：{新料號}」
    """
    # 找冒號位置，保留冒號前面的標籤
    for sep in ["：", ":"]:
        if sep in template_str:
            label = template_str.split(sep)[0] + sep
            new_val = _get_cell_value(item, order, field_key.replace("_inline",""), lot_seq, col_idx)
            return label + (new_val or "")
    return template_str


def generate_from_template(
    template_info: dict,
    orders: list[dict],
    template_wb: openpyxl.Workbook,
) -> BytesIO:
    """
    根據模板設定和訂單資料產生標籤 Excel
    
    template_info: analyze_template() 的回傳值（可能被使用者修改過欄位對應）
    orders: 解析後的銷貨單 list
    template_wb: 原始模板 workbook（用來複製樣式）
    """
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "出貨標籤"
    
    # 複製原始模板的樣式
    sheet_name = template_info.get("sheet_name", "")
    ws_tmpl = template_wb[sheet_name] if sheet_name in template_wb.sheetnames else None
    
    unit_rows        = template_info["unit_rows"]
    columns_per_unit = template_info["columns_per_unit"]
    gap_cols         = template_info.get("gap_cols", 1)
    units_per_row    = template_info.get("units_per_row", 2)
    cells_info       = template_info["cells"]
    col_widths       = template_info.get("col_widths", {})
    row_heights      = template_info.get("row_heights", {})
    
    # 設定欄寬（每個並排單元重複）
    for unit_col_idx in range(units_per_row):
        base_col = unit_col_idx * (columns_per_unit + gap_cols)
        for rel_col, width in col_widths.items():
            abs_col = base_col + int(rel_col)
            ws_out.column_dimensions[get_column_letter(abs_col)].width = float(width)
    
    # 收集所有要產生的品項
    all_items = []
    for order in orders:
        for item in order.get("items", []):
            all_items.append((item, order))
    
    current_row = 1
    
    for item_idx, (item, order) in enumerate(all_items):
        qty = int(item.get("quantity", 1) or 1)
        
        # 計算這個品項需要幾批（每批 units_per_row 個並排）
        # 共 qty 個標籤，每列 units_per_row 個
        total_labels = qty
        rows_needed  = (total_labels + units_per_row - 1) // units_per_row
        
        for label_row in range(rows_needed):
            # 設定行高
            for rel_row, height in row_heights.items():
                abs_row = current_row + int(rel_row) - 1
                ws_out.row_dimensions[abs_row].height = float(height)
            
            # 填入每個並排單元
            for unit_col_idx in range(units_per_row):
                label_seq = label_row * units_per_row + unit_col_idx + 1
                if label_seq > total_labels:
                    break  # 最後一列可能不滿
                
                base_col = unit_col_idx * (columns_per_unit + gap_cols)
                
                # 複製模板格子的樣式和值
                for cell_info in cells_info:
                    abs_row = current_row + cell_info["row"] - 1
                    abs_col = base_col + cell_info["col"]
                    
                    out_cell = ws_out.cell(row=abs_row, column=abs_col)
                    
                    field = cell_info.get("field", "__fixed__")
                    tmpl_val = cell_info.get("value", "")
                    
                    # 複製原始樣式
                    if ws_tmpl:
                        # 找對應的原始模板格
                        orig_row = cell_info["row"]
                        orig_col = cell_info["col"]
                        orig_cell = ws_tmpl.cell(row=orig_row, column=orig_col)
                        _copy_cell_style(orig_cell, out_cell)
                    
                    # 填值
                    if field == "__fixed__":
                        out_cell.value = tmpl_val
                    elif field.endswith("_inline"):
                        base_field = field.replace("_inline", "")
                        out_cell.value = _fill_inline_value(
                            tmpl_val, base_field, item, order, label_seq, unit_col_idx
                        )
                    elif field == "sequence":
                        out_cell.value = f"{label_seq:04d}"
                    elif field == "ship_date":
                        out_cell.value = _fmt_date(item.get("ship_date", ""))
                    elif field == "lot_no":
                        out_cell.value = _fmt_lot_no(order.get("order_no",""), label_seq, unit_col_idx)
                    elif field == "item_no":
                        out_cell.value = item.get("item_no", "")
                    elif field == "description":
                        out_cell.value = item.get("description", "")
                    elif field == "quantity":
                        out_cell.value = f"{item.get('quantity','')} {item.get('unit','PCS')}"
                    else:
                        new_val = _get_cell_value(item, order, field, label_seq, unit_col_idx)
                        out_cell.value = new_val if new_val else tmpl_val
            
            current_row += unit_rows
        
        # 品項間空一行
        current_row += 1
    
    buf = BytesIO()
    wb_out.save(buf)
    buf.seek(0)
    return buf


def _copy_cell_style(src, dst):
    """複製格子樣式"""
    try:
        if src.font:
            dst.font = copy.copy(src.font)
        if src.fill and src.fill.fill_type != "none":
            dst.fill = copy.copy(src.fill)
        if src.border:
            dst.border = copy.copy(src.border)
        if src.alignment:
            dst.alignment = copy.copy(src.alignment)
        if src.number_format:
            dst.number_format = src.number_format
    except Exception:
        pass


# ══════════════════════════════════════════════════════════════════
#  模板設定序列化（存進 Google Sheets）
# ══════════════════════════════════════════════════════════════════

def template_to_json(template_info: dict) -> str:
    """把模板設定轉成 JSON 字串存進 Sheets"""
    # 只存必要欄位，排除 workbook 物件
    save_keys = ["sheet_name","unit_rows","columns_per_unit","gap_cols",
                 "units_per_row","cells","col_widths","row_heights"]
    data = {k: template_info[k] for k in save_keys if k in template_info}
    return json.dumps(data, ensure_ascii=False)


def template_from_json(json_str: str) -> dict:
    return json.loads(json_str)


def get_field_options() -> list[str]:
    """回傳所有可選的動態欄位（給使用者在 UI 選擇）"""
    return FIELD_LABELS

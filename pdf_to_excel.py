"""
PDF 標籤 → Excel
每頁縮放到固定 8 格儲存格高度後插入 Excel。
"""
from io import BytesIO

import fitz  # PyMuPDF
from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage

TARGET_ROWS   = 8      # 每張標籤佔幾格
ROW_HEIGHT_PT = 15.0   # Excel 預設列高（points）
# 在 96dpi 下：1 pt = 96/72 px；TARGET_ROWS 格的像素高度
_TARGET_H_PX  = int(TARGET_ROWS * ROW_HEIGHT_PT * 96 / 72)   # ≈ 160 px


def pdf_to_excel(pdf_bytes: bytes) -> BytesIO:
    pdf = fitz.open(stream=pdf_bytes, filetype="pdf")

    wb = Workbook()
    ws = wb.active
    ws.title = "Labels"

    for page_idx, page in enumerate(pdf):
        # 用較高 DPI 渲染以保留品質，再縮小
        pix = page.get_pixmap(dpi=150)
        pil = PILImage.open(BytesIO(pix.tobytes("png")))

        # 縮放到目標高度，寬度等比
        orig_w, orig_h = pil.size
        new_h = _TARGET_H_PX
        new_w = max(1, int(orig_w * new_h / orig_h))
        pil = pil.resize((new_w, new_h), PILImage.LANCZOS)

        img_buf = BytesIO()
        pil.save(img_buf, format="PNG", optimize=True)
        img_buf.seek(0)

        xl_img = XLImage(img_buf)
        xl_img.height = new_h   # openpyxl: pixels
        xl_img.width  = new_w

        row_start = page_idx * TARGET_ROWS + 1
        ws.add_image(xl_img, f"A{row_start}")

        # 設定這 8 列的高度（points）
        for r in range(row_start, row_start + TARGET_ROWS):
            ws.row_dimensions[r].height = ROW_HEIGHT_PT

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

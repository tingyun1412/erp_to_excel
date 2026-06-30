from io import BytesIO

import fitz  # PyMuPDF
from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage


def pdf_to_excel(pdf_bytes: bytes) -> BytesIO:
    """
    PDF(bytes)
        ↓
    每頁轉圖片
        ↓
    放入 Excel
        ↓
    回傳 Excel(BytesIO)
    """

    # 開啟 PDF
    pdf = fitz.open(stream=pdf_bytes, filetype="pdf")

    wb = Workbook()
    ws = wb.active
    ws.title = "Labels"

    current_row = 1

    for page in pdf:

        # 300dpi 高解析
        pix = page.get_pixmap(dpi=300)

        img_bytes = pix.tobytes("png")

        # PIL 讀取
        pil_img = PILImage.open(BytesIO(img_bytes))

        img_stream = BytesIO()
        pil_img.save(img_stream, format="PNG")
        img_stream.seek(0)

        excel_img = XLImage(img_stream)

        # --------------------------
        # 自動縮放
        # --------------------------
        max_width = 520

        ratio = max_width / excel_img.width

        excel_img.width = int(excel_img.width * ratio)
        excel_img.height = int(excel_img.height * ratio)

        # 放到 Excel
        ws.add_image(excel_img, f"A{current_row}")

        # 調整列高
        row_height = excel_img.height * 0.75

        rows_used = int(row_height / 20) + 2

        for r in range(current_row, current_row + rows_used):
            ws.row_dimensions[r].height = 20

        current_row += rows_used

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output
"""
ERP 資料庫查詢模組

直接用「單據號碼」查詢公司 ERP 的 SQL Server 資料庫，取代舊的
「上傳銷貨單 RTF → 解析文字框座標」流程：批號、單價、金額都能直接查到，
不用等銷貨單印出來再上傳。

回傳的 dict 結構跟舊版 rtf_parser.parse_sales_order_rtf() 完全相容
（同樣的 key），所以 app.py 其他地方（標籤產生、電子發票）不用跟著改。

資料表對照（都在同一個 ERP 資料庫裡，透過「單據號碼」串起來）：
  WD4DT10A  銷貨單明細，DT10004=單據號碼（一張單可能有多列＝多個品項）
  WD2SU01   客戶主檔，SU01001=客戶代號

連線設定放在 .streamlit/secrets.toml 的 [mssql] 區塊（不會推上 GitHub）：
  [mssql]
  server   = "ntserver"
  database = "24405403"
  user     = "xxx"
  password = "xxx"
"""
import re

import streamlit as st

try:
    import pymssql
except ImportError:
    pymssql = None

from rtf_parser import _split_cjk_prefix

# WD4DT10A 欄位對照（單據明細）
_DT10_COLS = [
    "DT10004",  # 單據號碼
    "DT10005",  # 序號
    "DT10006",  # 料號
    "DT10008",  # 品名
    "DT10010",  # 數量
    "DT10011",  # 單價
    "DT10013",  # 金額
    "DT10021",  # 備註／客戶料號
    "DT10027",  # 客戶代號
    "DT10032",  # 批號
    "DT10801",  # 訂單單號（客戶 PO）
]

# WD2SU01 欄位對照（客戶主檔）
_SU01_COLS = [
    "SU01004",  # 客戶名稱（全稱）
    "SU01007",  # 統一編號
    "SU01010",  # 送貨地址
    "SU01013",  # 電話
    "SU01016",  # 傳真
    "SU01019",  # 手機
    "SU01022",  # 聯絡人
]


class ErpDbError(RuntimeError):
    """查詢 ERP 資料庫失敗時丟出，訊息已經是可以直接顯示給使用者看的中文。"""


def _get_conn():
    if pymssql is None:
        raise ErpDbError(
            "缺少 pymssql 套件，請先執行「pip install -r requirements.txt」再重新啟動程式。"
        )
    cfg = st.secrets.get("mssql")
    if not cfg:
        raise ErpDbError(
            "找不到 ERP 資料庫連線設定，請在 .streamlit/secrets.toml 加入 [mssql] "
            "區塊（server／database／user／password）。"
        )
    try:
        return pymssql.connect(
            server=cfg["server"],
            database=cfg["database"],
            user=cfg["user"],
            password=cfg["password"],
            timeout=15,
            login_timeout=15,
        )
    except KeyError as e:
        raise ErpDbError(f"[mssql] 設定缺少 {e} 這個欄位。") from e
    except Exception as e:
        raise ErpDbError(
            f"連不到 ERP 資料庫（是否在公司內網／已開 VPN？）：{e}"
        ) from e


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


def _looks_like_spec(text: str) -> bool:
    """
    粗略判斷一段文字像不像真正的品名規格：含 = 這種尺寸標記，
    且不是用「、」條列一堆客戶名稱的備註格式。
    WD4INVNA 有些記錄其實是「共用模具」的客戶清單備註（例如
    「鼎元、聯勝、光鋐;模具ID=0.12*OD=0.28mm」），不是這張訂單這個
    品項真正的品名規格，要濾掉避免誤用比銷貨單明細本身還離譜。
    """
    return bool(text) and "=" in text and "、" not in text


def _fetch_item_name(cursor, item_no: str, fallback_raw: str) -> tuple[str, str]:
    """
    查品保/驗收紀錄表 WD4INVNA 取得該料號比較可靠的品名/規格文字，
    用來取代銷貨單明細（WD4DT10A.DT10008）裡業務手動輸入、偶爾會漏字的版本
    （例如「二階」打成「二」）。取最新一筆，優先用 INVN006，看起來不像
    規格就改用 INVN005，兩個都不像規格（或這張表查不到這個料號）就
    退回用銷貨單明細自己的文字，不會硬套一個看起來就不對的值。
    回傳 (name, description)，用開頭連續中文字跟其餘部分切開。
    """
    raw = fallback_raw
    if item_no:
        try:
            cursor.execute(
                "SELECT TOP 1 INVN005, INVN006 FROM WD4INVNA "
                "WHERE INVN002 = %s ORDER BY INVN055 DESC",
                (item_no,),
            )
            row = cursor.fetchone()
            if row:
                invn005 = (row[0] or "").strip()
                invn006 = (row[1] or "").strip()
                if _looks_like_spec(invn006):
                    raw = invn006
                elif _looks_like_spec(invn005):
                    raw = invn005
        except Exception:
            pass  # 查不到就安靜退回用原本的文字，不擋整筆訂單
    return _split_cjk_prefix((raw or "").strip())


def _fetch_customer(cursor, customer_code: str) -> dict:
    """查客戶主檔，查不到就回傳全空字典（不擋單據本身的查詢結果）。"""
    if not customer_code:
        return {}
    cursor.execute(
        f"SELECT {', '.join(_SU01_COLS)} FROM WD2SU01 WHERE SU01001 = %s",
        (customer_code,),
    )
    row = cursor.fetchone()
    if not row:
        return {}
    name, tax_id, addr, phone, fax, mobile, contact_raw = row
    m = re.search(r"#\d+", contact_raw or "")
    contact = m.group(0) if m else (contact_raw or "")
    return {
        "customer_name": (name or "").strip(),
        "buyer_tax_id":  (tax_id or "").strip(),
        "address":       (addr or "").strip(),
        "phone":         (phone or "").strip(),
        "fax":           (fax or "").strip(),
        "mobile":        (mobile or "").strip(),
        "contact":       contact.strip(),
    }


def lookup_order_by_no(order_no: str) -> dict:
    """
    用「單據號碼」查詢完整銷貨單資料（含所有品項）。
    找不到資料時丟出 ErpDbError，訊息可直接顯示給使用者。
    """
    order_no = (order_no or "").strip()
    if not order_no:
        raise ErpDbError("單據號碼不能是空的。")

    conn = _get_conn()
    try:
        cursor = conn.cursor()
        cursor.execute(
            f"SELECT {', '.join(_DT10_COLS)} FROM WD4DT10A "
            f"WHERE DT10004 = %s ORDER BY DT10005",
            (order_no,),
        )
        rows = cursor.fetchall()
        if not rows:
            raise ErpDbError(f"查無單據號碼「{order_no}」，請確認號碼是否正確。")

        order_date = order_no[:8] if len(order_no) >= 8 else ""
        customer_code = (rows[0][8] or "").strip()  # DT10027

        result = {
            "filename":          f"DB-{order_no}",
            "order_no":          order_no,
            "order_date":        order_date,
            "customer_order_no": "",
            "invoice_no":        "",
            "seller_tax_id":     "",
            "buyer_tax_id":      "",
            "buyer_id_raw":      "",
            "phone":             "",
            "fax":               "",
            "mobile":            "",
            "contact":           "",
            "customer_code":     customer_code,
            "customer_name":     "",
            "items":             [],
            "raw_values":        [],
        }
        result.update(_fetch_customer(cursor, customer_code))
        result["buyer_id_raw"] = result["buyer_tax_id"]

        items = []
        for row in rows:
            (_doc_no, seq, item_no, name_raw, qty, unit_price, amount,
             remark, _cust_code, lot_no, cust_order_no) = row

            item = _blank_item(seq or "", order_date, result["customer_name"])
            item["item_no"]    = (item_no or "").strip()
            item["quantity"]   = int(qty or 0)
            item["unit_price"] = float(unit_price or 0)
            item["amount"]     = float(amount or 0)
            item["remark"]     = (remark or "").strip()
            item["lot_no"]     = (lot_no or "").strip()

            name_cjk, rest = _fetch_item_name(cursor, item["item_no"], name_raw)
            item["name"]        = name_cjk
            item["description"] = rest.strip()

            if cust_order_no and not result["customer_order_no"]:
                result["customer_order_no"] = (cust_order_no or "").strip()

            items.append(item)

        result["items"] = items
        return result
    finally:
        conn.close()


def lookup_orders_by_no(order_nos: list[str]) -> tuple[list[dict], list[str]]:
    """批次查詢多個單據號碼，回傳 (orders, errors)，跟舊版 parse_multiple_rtf 用法一致。"""
    orders, errors = [], []
    for no in order_nos:
        try:
            orders.append(lookup_order_by_no(no))
        except ErpDbError as e:
            errors.append(str(e))
        except Exception as e:
            errors.append(f"{no}: 查詢失敗（{e}）")
    return orders, errors

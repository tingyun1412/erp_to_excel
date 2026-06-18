"""
倉庫出貨自動化工具 v2
- Google Sheets 當資料庫（共享、持久）
- 出貨行事曆（網頁版，不用下載 Excel）
- 標籤欄位自訂 + 記憶廠商設定
"""
import calendar
from collections import defaultdict
from datetime import date, datetime

import streamlit as st

from rtf_parser import parse_sales_order_rtf
from module_b_invoice import generate_invoice_excel
from module_c_labels import generate_labels_excel, get_all_field_names, DEFAULT_FIELDS
from sheets_db import (
    append_schedule_rows, load_schedule, update_schedule_status,
    append_order,
    load_label_config, save_label_config, load_all_label_configs,
)

import tempfile

st.set_page_config(page_title="出貨自動化工具", page_icon="📦", layout="wide")

st.title("📦 出貨自動化工具")

if "parsed_orders" not in st.session_state:
    st.session_state.parsed_orders = []

WEEKDAY_ZH = ["一", "二", "三", "四", "五", "六", "日"]
STATUS_OPTIONS = ["待出貨", "已出貨", "延誤", "取消"]
STATUS_COLOR = {
    "待出貨": "#2E86C1",
    "已出貨": "#27AE60",
    "延誤":   "#E67E22",
    "取消":   "#95A5A6",
}


def _fmt_date_display(d: str) -> str:
    if len(d) == 8:
        return f"{d[:4]}/{d[4:6]}/{d[6:8]}"
    return d


def _parse_date(d: str):
    for fmt in ["%Y/%m/%d", "%Y%m%d"]:
        try:
            return datetime.strptime(d, fmt).date()
        except ValueError:
            pass
    return None


# ════════════════════════════════════════════════════════════════
#  側邊欄：上傳銷貨單
# ════════════════════════════════════════════════════════════════
with st.sidebar:
    st.header("上傳銷貨單")
    uploaded_files = st.file_uploader(
        "選擇 RTF 銷貨單（可多選）",
        type=["rtf"],
        accept_multiple_files=True,
    )

    if uploaded_files:
        if st.button("解析並匯入", type="primary", use_container_width=True):
            orders, errors = [], []
            with st.spinner("解析中..."):
                for uf in uploaded_files:
                    with tempfile.NamedTemporaryFile(suffix=".rtf", delete=False) as tmp:
                        tmp.write(uf.read())
                        tmp_path = tmp.name
                    try:
                        data = parse_sales_order_rtf(tmp_path)
                        data["filename"] = uf.name
                        orders.append(data)
                    except Exception as e:
                        errors.append(f"{uf.name}: {e}")

            if errors:
                for err in errors:
                    st.error(err)

            if orders:
                # 匯入 Sheets
                schedule_rows = []
                for order in orders:
                    append_order(order)
                    for item in order.get("items", []):
                        schedule_rows.append({
                            "銷貨單號":   order.get("order_no", ""),
                            "出貨日期":   _fmt_date_display(item.get("ship_date", "")),
                            "客戶名稱":   item.get("customer", ""),
                            "料號":       item.get("item_no", ""),
                            "品名":       item.get("description", ""),
                            "數量":       item.get("quantity", ""),
                            "單位":       item.get("unit", "PC"),
                            "客戶料號":   item.get("remark", ""),
                            "客戶訂單號": order.get("customer_order_no", ""),
                            "狀態":       "待出貨",
                            "備註":       "",
                        })

                added = append_schedule_rows(schedule_rows)
                st.session_state.parsed_orders = orders
                st.success(f"新增 {added} 筆，已存在的自動略過")
                st.rerun()

    st.divider()
    if st.session_state.parsed_orders:
        st.success(f"本次解析：{len(st.session_state.parsed_orders)} 張")
        if st.button("清除暫存", use_container_width=True):
            st.session_state.parsed_orders = []
            st.rerun()


# ════════════════════════════════════════════════════════════════
#  主頁籤
# ════════════════════════════════════════════════════════════════
tab_cal, tab_import, tab_b, tab_c = st.tabs([
    "📅 出貨行事曆",
    "📋 匯入記錄",
    "🧾 電子發票",
    "🏷 出貨標籤",
])


# ── 出貨行事曆 ────────────────────────────────────────────────────
with tab_cal:
    st.subheader("出貨行事曆")

    col_nav1, col_nav2, col_nav3, _ = st.columns([1, 2, 1, 4])
    today = date.today()

    if "cal_year"  not in st.session_state: st.session_state.cal_year  = today.year
    if "cal_month" not in st.session_state: st.session_state.cal_month = today.month

    with col_nav1:
        if st.button("◀ 上個月"):
            if st.session_state.cal_month == 1:
                st.session_state.cal_month = 12
                st.session_state.cal_year -= 1
            else:
                st.session_state.cal_month -= 1
            st.rerun()

    with col_nav2:
        tw_year = st.session_state.cal_year - 1911
        st.markdown(
            f"<h3 style='text-align:center;margin:0'>"
            f"{st.session_state.cal_year}/{st.session_state.cal_month:02d}"
            f"　（民國 {tw_year} 年）</h3>",
            unsafe_allow_html=True,
        )

    with col_nav3:
        if st.button("下個月 ▶"):
            if st.session_state.cal_month == 12:
                st.session_state.cal_month = 1
                st.session_state.cal_year += 1
            else:
                st.session_state.cal_month += 1
            st.rerun()

    # 載入資料
    with st.spinner("載入出貨排程..."):
        try:
            schedule = load_schedule()
        except Exception as e:
            st.error(f"無法連接資料庫：{e}")
            schedule = []

    # 依日期分組
    by_date = defaultdict(list)
    for i, row in enumerate(schedule):
        d = _parse_date(str(row.get("出貨日期", "")))
        if d:
            by_date[d].append({**row, "_row_index": i + 1})

    # 畫月曆
    year  = st.session_state.cal_year
    month = st.session_state.cal_month
    cal   = calendar.monthcalendar(year, month)

    # 星期標題
    header_cols = st.columns(7)
    for ci, wd in enumerate(WEEKDAY_ZH):
        bg = "#f0f0f0" if ci < 5 else "#fff3cd"
        header_cols[ci].markdown(
            f"<div style='background:{bg};text-align:center;"
            f"padding:6px;border-radius:4px;font-weight:bold'>{wd}</div>",
            unsafe_allow_html=True,
        )

    for week in cal:
        week_cols = st.columns(7)
        for ci, day_num in enumerate(week):
            with week_cols[ci]:
                if day_num == 0:
                    st.markdown("<div style='min-height:80px'></div>", unsafe_allow_html=True)
                    continue

                d = date(year, month, day_num)
                items_today = by_date.get(d, [])
                is_today = (d == today)
                is_weekend = ci >= 5

                day_bg = "#FFF9C4" if is_today else ("#FFF8E1" if is_weekend else "#FFFFFF")
                day_border = "2px solid #F39C12" if is_today else "1px solid #ddd"

                # 日期格
                html = (
                    f"<div style='background:{day_bg};border:{day_border};"
                    f"border-radius:6px;padding:6px;min-height:80px'>"
                    f"<div style='font-weight:bold;font-size:14px'>{day_num}</div>"
                )
                for item in items_today[:3]:
                    status = item.get("狀態", "待出貨")
                    color  = STATUS_COLOR.get(status, "#888")
                    label  = f"{item.get('客戶名稱','')} {item.get('料號','')}"
                    html += (
                        f"<div style='background:{color};color:white;"
                        f"border-radius:3px;padding:2px 4px;margin-top:2px;"
                        f"font-size:10px;overflow:hidden;white-space:nowrap;"
                        f"text-overflow:ellipsis' title='{label}'>{label}</div>"
                    )
                if len(items_today) > 3:
                    html += f"<div style='font-size:10px;color:#888'>+{len(items_today)-3} 筆</div>"
                html += "</div>"
                st.markdown(html, unsafe_allow_html=True)

    # 當月明細表
    st.divider()
    st.markdown("**當月出貨明細**")

    month_items = []
    for d, items in by_date.items():
        if d.year == year and d.month == month:
            month_items.extend(items)

    if not month_items:
        st.info("本月尚無出貨記錄")
    else:
        # 狀態篩選
        filter_status = st.multiselect(
            "篩選狀態",
            STATUS_OPTIONS,
            default=["待出貨", "延誤"],
        )
        filtered = [r for r in month_items if r.get("狀態", "") in filter_status]

        for row in sorted(filtered, key=lambda x: str(x.get("出貨日期", ""))):
            status = row.get("狀態", "待出貨")
            color  = STATUS_COLOR.get(status, "#888")
            with st.expander(
                f"{row.get('出貨日期','')}　{row.get('客戶名稱','')}　"
                f"{row.get('料號','')}　× {row.get('數量','')}",
                expanded=False,
            ):
                c1, c2, c3 = st.columns([2, 2, 2])
                with c1:
                    st.write(f"**品名：** {row.get('品名','')}")
                    st.write(f"**客戶料號：** {row.get('客戶料號','')}")
                    st.write(f"**客戶訂單：** {row.get('客戶訂單號','')}")
                with c2:
                    st.write(f"**銷貨單號：** {row.get('銷貨單號','')}")
                    st.write(f"**備註：** {row.get('備註','')}")
                with c3:
                    new_status = st.selectbox(
                        "狀態",
                        STATUS_OPTIONS,
                        index=STATUS_OPTIONS.index(status) if status in STATUS_OPTIONS else 0,
                        key=f"status_{row.get('_row_index')}",
                    )
                    new_remark = st.text_input(
                        "備註",
                        value=row.get("備註", ""),
                        key=f"remark_{row.get('_row_index')}",
                    )
                    if st.button("儲存", key=f"save_{row.get('_row_index')}"):
                        try:
                            update_schedule_status(
                                row["_row_index"], new_status, new_remark
                            )
                            st.success("已更新")
                            st.rerun()
                        except Exception as e:
                            st.error(f"更新失敗：{e}")


# ── 匯入記錄 ─────────────────────────────────────────────────────
with tab_import:
    st.subheader("匯入記錄")
    try:
        schedule = load_schedule()
        if not schedule:
            st.info("尚無資料")
        else:
            # 搜尋
            search = st.text_input("搜尋（銷貨單號 / 客戶 / 料號）")
            rows = schedule
            if search:
                rows = [
                    r for r in rows
                    if search.lower() in str(r.get("銷貨單號","")).lower()
                    or search.lower() in str(r.get("客戶名稱","")).lower()
                    or search.lower() in str(r.get("料號","")).lower()
                ]
            st.write(f"共 {len(rows)} 筆")
            st.dataframe(
                rows,
                use_container_width=True,
                hide_index=True,
                column_order=["出貨日期","銷貨單號","客戶名稱","料號","品名","數量","單位","客戶料號","客戶訂單號","狀態","備註"],
            )
    except Exception as e:
        st.error(f"載入失敗：{e}")


# ── 電子發票 ─────────────────────────────────────────────────────
with tab_b:
    st.subheader("電子發票")
    st.caption("依照 e-invoice.com.tw V1.6 格式產生上傳檔")

    orders = st.session_state.parsed_orders
    if not orders:
        st.info("請先在左側上傳並解析銷貨單")
    else:
        with st.expander("發票設定", expanded=True):
            c1, c2, c3 = st.columns(3)
            with c1:
                default_seller = next(
                    (o.get("seller_tax_id","") for o in orders if o.get("seller_tax_id")), ""
                )
                seller_id = st.text_input("賣方統編", value=default_seller)
            with c2:
                inv_prefix = st.text_input("發票字軌（2碼英文）", value="AA", max_chars=2)
            with c3:
                start_num = st.number_input("起始號碼", min_value=1, value=1)

        orders_with_items = [o for o in orders if o.get("items")]
        st.write(f"共 {len(orders_with_items)} 張發票")

        if orders_with_items:
            if st.button("產出電子發票 Excel", type="primary", use_container_width=True):
                buf = generate_invoice_excel(
                    orders_with_items,
                    seller_tax_id=seller_id,
                    invoice_prefix=inv_prefix,
                    start_number=int(start_num),
                )
                st.download_button(
                    "⬇️ 下載電子發票上傳檔.xlsx",
                    data=buf,
                    file_name="電子發票上傳.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )


# ── 出貨標籤 ─────────────────────────────────────────────────────
with tab_c:
    st.subheader("出貨標籤")

    orders = st.session_state.parsed_orders
    if not orders:
        st.info("請先在左側上傳並解析銷貨單")
    else:
        all_fields = get_all_field_names()

        # 找出本批次涉及的廠商
        customers = list({
            item.get("customer", "")
            for o in orders for item in o.get("items", [])
            if item.get("customer")
        })

        # 廠商選擇
        selected_customer = st.selectbox(
            "選擇廠商套用設定",
            options=["（不套用記憶設定）"] + customers,
        )

        # 讀取記憶設定
        saved_fields = None
        if selected_customer != "（不套用記憶設定）":
            try:
                saved_fields = load_label_config(selected_customer)
            except Exception:
                saved_fields = None

        default_selection = saved_fields if saved_fields else DEFAULT_FIELDS

        st.markdown("**選擇並排列欄位**")
        st.caption("勾選要顯示的欄位，上下拖曳調整順序（用數字輸入框調整優先序）")

        # 用數字輸入讓使用者設定順序
        field_orders = {}
        checked_fields = {}

        cols = st.columns(3)
        for i, field in enumerate(all_fields):
            with cols[i % 3]:
                is_checked = field in default_selection
                checked = st.checkbox(field, value=is_checked, key=f"chk_{field}")
                checked_fields[field] = checked
                if checked:
                    current_order = (
                        default_selection.index(field) + 1
                        if field in default_selection
                        else len(default_selection) + i + 1
                    )
                    order_val = st.number_input(
                        f"順序",
                        min_value=1, max_value=20,
                        value=current_order,
                        key=f"ord_{field}",
                        label_visibility="collapsed",
                    )
                    field_orders[field] = order_val

        # 依順序排列選取的欄位
        selected_fields = sorted(
            [f for f, checked in checked_fields.items() if checked],
            key=lambda f: field_orders.get(f, 99),
        )

        if selected_fields:
            st.markdown(f"**預覽順序：** {' → '.join(selected_fields)}")

        c1, c2 = st.columns(2)
        with c1:
            if selected_customer != "（不套用記憶設定）" and selected_fields:
                if st.button(f"💾 記住「{selected_customer}」的設定", use_container_width=True):
                    try:
                        save_label_config(selected_customer, selected_fields)
                        st.success("已儲存")
                    except Exception as e:
                        st.error(f"儲存失敗：{e}")

        with c2:
            if selected_fields:
                if st.button("產出標籤 Excel", type="primary", use_container_width=True):
                    buf = generate_labels_excel(orders, fields=selected_fields)
                    st.download_button(
                        "⬇️ 下載出貨標籤.xlsx",
                        data=buf,
                        file_name="出貨標籤.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
            else:
                st.warning("請至少勾選一個欄位")

        # 顯示所有廠商記憶設定
        with st.expander("所有廠商的已儲存設定"):
            try:
                all_configs = load_all_label_configs()
                if all_configs:
                    for cust, fields in all_configs.items():
                        st.write(f"**{cust}**：{' → '.join(fields)}")
                else:
                    st.info("尚無儲存的設定")
            except Exception as e:
                st.error(f"載入失敗：{e}")
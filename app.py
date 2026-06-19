"""
出貨自動化工具 v3
模組：
  A - 電子發票產生
  B - 標籤模板管理（上傳舊標籤 → 設定欄位 → 產出新標籤）
"""
import json
import tempfile
from io import BytesIO

import streamlit as st
import openpyxl

from rtf_parser import parse_sales_order_rtf
from module_b_invoice import generate_invoice_excel
from template_engine import (
    analyze_template, analyze_all_sheets, generate_from_template,
    template_to_json, template_from_json,
    get_field_options, FIELD_LABELS, DYNAMIC_FIELDS,
)
from sheets_db import (
    load_templates, save_template, delete_template,
    clear_cache,
)

st.set_page_config(page_title="出貨自動化工具", page_icon="📦", layout="wide")
st.title("📦 出貨自動化工具")

if "parsed_orders" not in st.session_state:
    st.session_state.parsed_orders = []
if "template_wb_bytes" not in st.session_state:
    st.session_state.template_wb_bytes = {}  # {template_key: bytes}



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

            for err in errors:
                st.error(err)

            if orders:
                st.session_state.parsed_orders = orders
                st.success(f"解析完成，共 {len(orders)} 張銷貨單")
                st.rerun()

    st.divider()
    if st.session_state.parsed_orders:
        st.success(f"已載入 {len(st.session_state.parsed_orders)} 張銷貨單")
        if st.button("清除暫存", use_container_width=True):
            st.session_state.parsed_orders = []
            st.rerun()


# ════════════════════════════════════════════════════════════════
#  主頁籤
# ════════════════════════════════════════════════════════════════
tab_label, tab_invoice = st.tabs([
    "🏷 出貨標籤",
    "🧾 電子發票",
])


# ════════════════════════════════════════════════════════════════
#  標籤模板管理
# ════════════════════════════════════════════════════════════════
with tab_label:
    st.subheader("出貨標籤")

    sub_tab_use, sub_tab_manage = st.tabs(["產出標籤", "管理模板"])

    # ── 產出標籤 ──────────────────────────────────────────────
    with sub_tab_use:
        all_orders = st.session_state.parsed_orders

        if not all_orders:
            st.info("請先在左側上傳並解析銷貨單")
        else:
            # 載入所有模板
            try:
                all_templates = load_templates()
            except Exception as e:
                st.error(f"載入模板失敗：{e}")
                all_templates = []

            if not all_templates:
                st.warning("尚無任何標籤模板，請先到「管理模板」上傳")
            else:
                col_left, col_right = st.columns([1, 1])

                with col_left:
                    # 選模板
                    template_options = [
                        f"{r['廠商名稱']} — {r['模板名稱']}"
                        for r in all_templates
                    ]
                    selected_idx = st.selectbox(
                        "選擇標籤模板",
                        options=range(len(template_options)),
                        format_func=lambda i: template_options[i],
                    )
                    selected_record = all_templates[selected_idx]
                    template_info = template_from_json(selected_record["設定JSON"])
                    st.caption(
                        f"標籤單元 {template_info.get('unit_rows')} 行　"
                        f"｜　每列 {template_info.get('units_per_row')} 個並排"
                    )

                with col_right:
                    # 選銷貨單
                    order_options = {
                        o.get("order_no", o.get("filename", f"單{i}")): o
                        for i, o in enumerate(all_orders)
                    }
                    selected_order_nos = st.multiselect(
                        "選擇要產生標籤的銷貨單",
                        options=list(order_options.keys()),
                        default=list(order_options.keys()),
                    )
                    selected_orders = [order_options[k] for k in selected_order_nos]

                # 顯示選取的品項
                selected_items = [
                    (item, order)
                    for order in selected_orders
                    for item in order.get("items", [])
                ]
                if selected_items:
                    st.dataframe(
                        [{
                            "銷貨單號": order.get("order_no",""),
                            "料號":     item.get("item_no",""),
                            "品名":     item.get("description",""),
                            "數量":     item.get("quantity",""),
                            "客戶料號": item.get("remark",""),
                        } for item, order in selected_items],
                        use_container_width=True,
                        hide_index=True,
                    )
                    st.caption(f"共 {len(selected_items)} 個品項，依數量產生對應張數標籤")
                else:
                    st.warning("請選擇至少一張銷貨單")

                # 模板 workbook（用來複製樣式）
                tmpl_key = f"{selected_record['廠商名稱']}_{selected_record['模板名稱']}"
                if tmpl_key not in st.session_state.template_wb_bytes:
                    st.warning("需要上傳原始模板 Excel 才能保留字型/框線樣式")
                    re_upload = st.file_uploader(
                        "重新上傳此廠商的模板 Excel",
                        type=["xlsx", "xls"],
                        key=f"reupload_{tmpl_key}",
                    )
                    if re_upload:
                        st.session_state.template_wb_bytes[tmpl_key] = re_upload.read()
                        st.rerun()

                if selected_items and st.button("產出標籤 Excel", type="primary", use_container_width=True):
                    with st.spinner("產出中..."):
                        try:
                            wb_bytes = st.session_state.template_wb_bytes.get(tmpl_key)
                            template_wb = (
                                openpyxl.load_workbook(BytesIO(wb_bytes))
                                if wb_bytes else openpyxl.Workbook()
                            )
                            buf = generate_from_template(template_info, selected_orders, template_wb)
                            st.download_button(
                                "⬇️ 下載標籤.xlsx",
                                data=buf,
                                file_name=f"標籤_{selected_record['廠商名稱']}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True,
                            )
                        except Exception as e:
                            st.error(f"產出失敗：{e}")
                            import traceback
                            st.code(traceback.format_exc())

    # ── 管理模板 ──────────────────────────────────────────────
    with sub_tab_manage:
        st.markdown("### 上傳新模板")
        st.caption("上傳舊的標籤 Excel，系統自動分析格式，讓你確認欄位對應後存檔")

        col1, col2 = st.columns(2)
        with col1:
            new_customer = st.text_input("廠商名稱", placeholder="例：晶晟")
        with col2:
            new_tmpl_name = st.text_input("模板名稱", placeholder="例：標準出貨標籤")

        uploaded_tmpl = st.file_uploader(
            "上傳標籤 Excel 範本",
            type=["xlsx", "xls"],
            key="new_template_upload",
        )

        if uploaded_tmpl and new_customer:
            tmpl_bytes = uploaded_tmpl.read()
            wb = openpyxl.load_workbook(BytesIO(tmpl_bytes))
            sheet_names = wb.sheetnames

            batch_mode = st.checkbox(
                "批次建立（每個工作表自動建立一個模板）",
                value=len(sheet_names) > 1,
            )

            if batch_mode:
                if st.button("批次分析所有工作表", type="primary"):
                    with st.spinner(f"分析 {len(sheet_names)} 個工作表..."):
                        results = analyze_all_sheets(wb)
                    if not results:
                        st.error("無法分析任何工作表")
                    else:
                        saved, skipped = 0, 0
                        for sname, info in results.items():
                            try:
                                save_template(new_customer, sname, template_to_json(info))
                                tmpl_key = f"{new_customer}_{sname}"
                                st.session_state.template_wb_bytes[tmpl_key] = tmpl_bytes
                                saved += 1
                            except Exception:
                                skipped += 1
                        clear_cache()
                        st.success(f"完成！成功建立 {saved} 個模板" + (f"，{skipped} 個失敗" if skipped else ""))
                        st.rerun()
            else:
                selected_sheet = st.selectbox("選擇要分析的工作表", sheet_names)
                if not new_tmpl_name:
                    st.info("請填寫模板名稱")
                elif st.button("分析模板", type="primary"):
                    with st.spinner("分析中..."):
                        template_info = analyze_template(wb, selected_sheet)
                    if not template_info:
                        st.error("無法分析此工作表，請確認內容不為空")
                    else:
                        st.success(
                            f"分析完成：{template_info['unit_rows']} 行/標籤，"
                            f"每列 {template_info['units_per_row']} 個並排"
                        )
                        st.session_state["_pending_template"] = template_info
                        st.session_state["_pending_tmpl_bytes"] = tmpl_bytes
                        st.session_state["_pending_customer"] = new_customer
                        st.session_state["_pending_tmpl_name"] = new_tmpl_name

        # 欄位對應確認
        if "_pending_template" in st.session_state:
            template_info = st.session_state["_pending_template"]
            cells = template_info.get("cells", [])

            st.divider()
            st.markdown("### 確認欄位對應")
            st.caption("檢查每個格子的欄位是否正確，可以修改")

            field_options = ["__fixed__（固定文字）"] + [
                f"{k}（{v}）" for k, v in DYNAMIC_FIELDS.items() if k != "固定文字"
            ]

            updated_cells = []
            for i, cell in enumerate(cells):
                c1, c2, c3 = st.columns([1, 3, 3])
                with c1:
                    st.text(f"R{cell['row']}C{cell['col']}")
                with c2:
                    st.text(cell['value'][:40] if cell['value'] else "")
                with c3:
                    current = cell.get("field", "__fixed__")
                    # 找目前值在選項中的位置
                    curr_display = next(
                        (opt for opt in field_options if opt.startswith(current)),
                        field_options[0]
                    )
                    chosen = st.selectbox(
                        "欄位",
                        field_options,
                        index=field_options.index(curr_display) if curr_display in field_options else 0,
                        key=f"field_map_{i}",
                        label_visibility="collapsed",
                    )
                    # 取出 key（括號前）
                    new_field = chosen.split("（")[0]
                    updated_cell = {**cell, "field": new_field}
                    updated_cells.append(updated_cell)

            if st.button("儲存模板", type="primary", use_container_width=True):
                template_info["cells"] = updated_cells
                config_json = template_to_json(template_info)
                customer = st.session_state["_pending_customer"]
                tmpl_name = st.session_state["_pending_tmpl_name"]
                tmpl_key = f"{customer}_{tmpl_name}"

                try:
                    save_template(customer, tmpl_name, config_json)
                    # 快取 workbook bytes
                    st.session_state.template_wb_bytes[tmpl_key] = st.session_state["_pending_tmpl_bytes"]
                    # 清除暫存
                    del st.session_state["_pending_template"]
                    del st.session_state["_pending_tmpl_bytes"]
                    del st.session_state["_pending_customer"]
                    del st.session_state["_pending_tmpl_name"]
                    clear_cache()
                    st.success(f"模板「{customer} — {tmpl_name}」已儲存！")
                    st.rerun()
                except Exception as e:
                    st.error(f"儲存失敗：{e}")

        # 現有模板列表
        st.divider()
        st.markdown("### 現有模板")
        try:
            all_templates = load_templates()
            if not all_templates:
                st.info("尚無模板")
            else:
                for r in all_templates:
                    with st.expander(f"{r['廠商名稱']} — {r['模板名稱']}　（更新：{r.get('最後更新','')}）"):
                        info = template_from_json(r["設定JSON"])
                        st.write(f"標籤行數：{info.get('unit_rows')}，並排數：{info.get('units_per_row')}")
                        cells = info.get("cells", [])
                        dynamic = [c for c in cells if c.get("field") != "__fixed__"]
                        st.write(f"動態欄位：{[c['field'] for c in dynamic]}")
                        if st.button("刪除此模板", key=f"del_{r['廠商名稱']}_{r['模板名稱']}"):
                            try:
                                delete_template(r["廠商名稱"], r["模板名稱"])
                                clear_cache()
                                st.success("已刪除")
                                st.rerun()
                            except Exception as e:
                                st.error(f"刪除失敗：{e}")
        except Exception as e:
            st.error(f"載入失敗：{e}")


# ════════════════════════════════════════════════════════════════
#  電子發票
# ════════════════════════════════════════════════════════════════
with tab_invoice:
    st.subheader("電子發票")
    st.caption("依照 e-invoice.com.tw V1.6 格式產生上傳檔")

    all_orders = st.session_state.parsed_orders
    if not all_orders:
        st.info("請先在左側上傳並解析銷貨單")
    else:
        # 選銷貨單
        order_map = {
            f"{o.get('order_no','')} — {o.get('customer_code','')} ({len(o.get('items',[]))} 項)": o
            for o in all_orders
        }
        selected_keys = st.multiselect(
            "選擇要開發票的銷貨單",
            options=list(order_map.keys()),
            default=list(order_map.keys()),
        )
        selected_orders = [order_map[k] for k in selected_keys if order_map[k].get("items")]

        if selected_orders:
            with st.expander("發票設定", expanded=True):
                c1, c2, c3 = st.columns(3)
                with c1:
                    default_seller = next(
                        (o.get("seller_tax_id","") for o in selected_orders if o.get("seller_tax_id")), ""
                    )
                    seller_id = st.text_input("賣方統編", value=default_seller)
                with c2:
                    inv_prefix = st.text_input("發票字軌（2碼英文）", value="AA", max_chars=2)
                with c3:
                    start_num = st.number_input("起始號碼", min_value=1, value=1)

            st.write(f"共 {len(selected_orders)} 張發票")
            if st.button("產出電子發票 Excel", type="primary", use_container_width=True):
                buf = generate_invoice_excel(
                    selected_orders,
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
        else:
            st.warning("請選擇至少一張銷貨單")
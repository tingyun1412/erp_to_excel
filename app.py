"""
出貨自動化工具 v3
模組：
  A - 電子發票產生
  B - 標籤模板管理（上傳舊標籤 → 設定欄位 → 產出新標籤）
"""
import json
import re
import tempfile
from io import BytesIO

import streamlit as st
import openpyxl
import subprocess
import sys
from pathlib import Path

chrome_dir = Path.home() / ".cache/ms-playwright"

if not chrome_dir.exists():
    subprocess.run(
        [sys.executable, "-m", "playwright", "install", "chromium"],
        check=True
    )
from rtf_parser import parse_sales_order_rtf
from module_b_invoice import generate_invoice_excel
from template_engine import (
    analyze_template, analyze_all_sheets,
    generate_from_template, generate_labels_multiorder,
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

    sub_tab_use, sub_tab_manage, sub_tab_erp = st.tabs(["產出標籤", "管理模板", "ERP 下載 PDF"])

    # ── 產出標籤 ──────────────────────────────────────────────
    with sub_tab_use:
        all_orders = st.session_state.parsed_orders

        if not all_orders:
            st.info("請先在左側上傳並解析銷貨單")
        else:
            try:
                all_templates = load_templates()
            except Exception as e:
                st.error(f"載入模板失敗：{e}")
                all_templates = []

            if not all_templates:
                st.warning("尚無任何標籤模板，請先到「管理模板」上傳")
            else:
                def _match_template(customer_name: str):
                    """找廠商名稱與客戶名稱有包含關係的模板（第一個符合的）"""
                    if not customer_name:
                        return None
                    for r in all_templates:
                        vendor = r.get("廠商名稱", "")
                        if vendor and (vendor in customer_name or customer_name in vendor):
                            return r
                    return None

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

                template_options = [
                    f"{r['廠商名稱']} — {r['模板名稱']}"
                    for r in all_templates
                ]

                # 每張銷貨單獨立選擇模板（自動比對可手動覆蓋）
                order_tmpl_map = {}
                if selected_orders:
                    st.markdown("**各銷貨單模板（自動比對，可手動覆蓋）：**")
                    for o in selected_orders:
                        order_key = o.get("order_no", o.get("filename", ""))
                        cname = o.get("customer_name", "")
                        auto_match = _match_template(cname)
                        default_idx = next(
                            (i for i, r in enumerate(all_templates) if r is auto_match), 0
                        )
                        _c1, _c2 = st.columns([5, 5])
                        with _c1:
                            _lbl = "✅ 自動" if auto_match else "⚠️ 未比對"
                            st.caption(f"{_lbl}　{order_key}　{cname or '未知客戶'}")
                        with _c2:
                            _chosen = st.selectbox(
                                "模板",
                                options=range(len(template_options)),
                                format_func=lambda i: template_options[i],
                                index=default_idx,
                                key=f"tmpl_sel_{order_key}",
                                label_visibility="collapsed",
                            )
                        order_tmpl_map[order_key] = all_templates[_chosen]

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
                            "品名":     item.get("name",""),
                            "規格":     item.get("description",""),
                            "數量":     item.get("quantity",""),
                            "客戶料號": item.get("remark",""),
                            "批號":     item.get("lot_no",""),
                        } for item, order in selected_items],
                        use_container_width=True,
                        hide_index=True,
                        height=min(420, 38 * (len(selected_items) + 1) + 10),
                        column_config={
                            "銷貨單號": st.column_config.TextColumn(width="medium"),
                            "料號":     st.column_config.TextColumn(width="large"),
                            "品名":     st.column_config.TextColumn(width="small"),
                            "規格":     st.column_config.TextColumn(width="large"),
                            "數量":     st.column_config.NumberColumn(width="small"),
                            "客戶料號": st.column_config.TextColumn(width="medium"),
                            "批號":     st.column_config.TextColumn(width="medium"),
                        },
                    )
                    st.caption(f"共 {len(selected_items)} 個品項，每張銷貨單一個工作表")
                else:
                    st.warning("請選擇至少一張銷貨單")

                if selected_orders and order_tmpl_map and st.button("產出標籤 Excel", type="primary", use_container_width=True):
                    with st.spinner("產出中..."):
                        try:
                            pairs = []
                            for o in selected_orders:
                                order_key = o.get("order_no", o.get("filename", ""))
                                rec = order_tmpl_map.get(order_key, all_templates[0])
                                tmpl_key = f"{rec['廠商名稱']}_{rec['模板名稱']}"
                                wb_bytes = st.session_state.template_wb_bytes.get(tmpl_key)
                                twb = (
                                    openpyxl.load_workbook(BytesIO(wb_bytes))
                                    if wb_bytes else openpyxl.Workbook()
                                )
                                pairs.append({
                                    "order": o,
                                    "template_info": template_from_json(rec["設定JSON"]),
                                    "template_wb": twb,
                                })

                            buf = generate_labels_multiorder(pairs)
                            st.download_button(
                                "⬇️ 下載標籤.xlsx",
                                data=buf,
                                file_name="出貨標籤.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True,
                            )
                        except Exception as e:
                            st.error(f"產出失敗：{e}")
                            import traceback
                            st.code(traceback.format_exc())

                # 若模板 Excel 尚未上傳，提供重新上傳入口
                if selected_orders and order_tmpl_map:
                    _shown_keys = set()
                    for _ok, _rec in order_tmpl_map.items():
                        _tk = f"{_rec['廠商名稱']}_{_rec['模板名稱']}"
                        if _tk in _shown_keys or _tk in st.session_state.template_wb_bytes:
                            continue
                        _shown_keys.add(_tk)
                        _re = st.file_uploader(
                            f"上傳「{_rec['廠商名稱']} — {_rec['模板名稱']}」原始 Excel（保留樣式用）",
                            type=["xlsx", "xls"],
                            key=f"reupload_{_tk}",
                        )
                        if _re:
                            st.session_state.template_wb_bytes[_tk] = _re.read()
                            st.rerun()

    # ── 管理模板 ──────────────────────────────────────────────
    with sub_tab_manage:
        st.markdown("### 上傳新模板")
        st.caption("上傳舊的標籤 Excel，系統自動分析格式，讓你確認欄位對應後存檔；廠商名稱／模板名稱留空時，直接使用工作表名稱命名")

        col1, col2 = st.columns(2)
        with col1:
            new_customer = st.text_input("廠商名稱（留空則用工作表名稱）", placeholder="例：晶晟")
        with col2:
            new_tmpl_name = st.text_input("模板名稱（留空則用工作表名稱）", placeholder="例：標準出貨標籤")

        uploaded_tmpl = st.file_uploader(
            "上傳標籤 Excel 範本",
            type=["xlsx", "xls"],
            key="new_template_upload",
        )

        if uploaded_tmpl:
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
                                vendor = new_customer.strip() or sname
                                tname = new_tmpl_name.strip() or sname
                                save_template(vendor, tname, template_to_json(info))
                                tmpl_key = f"{vendor}_{tname}"
                                st.session_state.template_wb_bytes[tmpl_key] = tmpl_bytes
                                saved += 1
                            except Exception:
                                skipped += 1
                        clear_cache()
                        st.success(f"完成！成功建立 {saved} 個模板" + (f"，{skipped} 個失敗" if skipped else ""))
                        st.rerun()
            else:
                selected_sheet = st.selectbox("選擇要分析的工作表", sheet_names)
                if st.button("分析模板", type="primary"):
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
                        st.session_state["_pending_customer"] = new_customer.strip() or selected_sheet
                        st.session_state["_pending_tmpl_name"] = new_tmpl_name.strip() or selected_sheet

        # 欄位對應確認
        if "_pending_template" in st.session_state:
            template_info = st.session_state["_pending_template"]
            cells = template_info.get("cells", [])

            st.divider()
            st.markdown("### 確認欄位對應")
            st.caption("檢查每個格子的欄位是否正確，可以修改")

            # 名稱可編輯（新增時預填，編輯時可改名）
            _ec1, _ec2 = st.columns(2)
            with _ec1:
                _edit_customer = st.text_input(
                    "廠商名稱（用於比對）",
                    value=st.session_state.get("_pending_customer", ""),
                    key="_edit_customer_input",
                )
            with _ec2:
                _edit_tmpl_name = st.text_input(
                    "模板名稱",
                    value=st.session_state.get("_pending_tmpl_name", ""),
                    key="_edit_tmpl_name_input",
                )

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
                customer = _edit_customer.strip() or st.session_state.get("_pending_customer", "")
                tmpl_name = _edit_tmpl_name.strip() or st.session_state.get("_pending_tmpl_name", "")
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
                        _btn_edit, _btn_del = st.columns(2)
                        with _btn_edit:
                            if st.button("✏️ 編輯欄位", key=f"edit_{r['廠商名稱']}_{r['模板名稱']}"):
                                st.session_state["_pending_template"] = dict(info)
                                st.session_state["_pending_tmpl_bytes"] = st.session_state.template_wb_bytes.get(
                                    f"{r['廠商名稱']}_{r['模板名稱']}", b""
                                )
                                st.session_state["_pending_customer"] = r["廠商名稱"]
                                st.session_state["_pending_tmpl_name"] = r["模板名稱"]
                                st.rerun()
                        with _btn_del:
                            if st.button("🗑 刪除此模板", key=f"del_{r['廠商名稱']}_{r['模板名稱']}"):
                                try:
                                    delete_template(r["廠商名稱"], r["模板名稱"])
                                    clear_cache()
                                    st.success("已刪除")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"刪除失敗：{e}")
        except Exception as e:
            st.error(f"載入失敗：{e}")

    # ── 廠商網站下載標籤 ─────────────────────────────────────────
    with sub_tab_erp:
        st.caption("直接從廠商網站下載標籤 PDF（需要電腦可連到 ERP 網路），可從已解析的銷貨單勾選，或直接輸入銷貨單號")

        all_orders = st.session_state.parsed_orders
        order_nos_from_parsed = [o.get("order_no", "") for o in all_orders if o.get("order_no")]

        input_mode = st.radio(
            "選擇出貨單方式",
            ["從已解析銷貨單勾選", "直接輸入銷貨單號"],
            horizontal=True,
        )

        selected_nos = []
        if input_mode == "從已解析銷貨單勾選":
            if not order_nos_from_parsed:
                st.info("尚無已解析的銷貨單，請先在左側上傳，或改用「直接輸入銷貨單號」")
            else:
                selected_nos = st.multiselect(
                    "選擇要下載的出貨單",
                    options=order_nos_from_parsed,
                    default=order_nos_from_parsed,
                )
        else:
            raw_text = st.text_area(
                "輸入銷貨單號（可多筆，用換行或逗號分隔）",
                placeholder="202606100012\n202606020007",
                height=120,
            )
            selected_nos = [
                n.strip() for n in re.split(r"[,\n，、]+", raw_text) if n.strip()
            ]
            if selected_nos:
                st.caption(f"共 {len(selected_nos)} 筆：{', '.join(selected_nos)}")

        col1, col2 = st.columns(2)
        with col1:
            erp_user = st.text_input("ERP 帳號", value="BR026", key="erp_user")
        with col2:
            erp_pass = st.text_input("ERP 密碼", value="5403", type="password", key="erp_pass")

        if selected_nos and st.button("從廠商網站下載標籤xlsx", type="primary", use_container_width=True):
            with st.spinner("連線 ERP 並下載中，請稍候..."):
                try:
                    from erp_downloader import download_label_pdfs, pack_zip
                    results, erp_errors = download_label_pdfs(selected_nos, erp_user, erp_pass)

                    success = {no: data for no, data in results.items() if data}
                    failed  = [no for no, data in results.items() if not data]

                    if success:
                        order_no, pdf_bytes = next(iter(success.items()))
                        if len(success) == 1:
                            # 單筆：先提供 PDF 下載，再嘗試轉 Excel
                            st.download_button(
                                f"⬇️ 下載 {order_no} 標籤.pdf（原始）",
                                data=pdf_bytes,
                                file_name=f"標籤_{order_no}.pdf",
                                mime="application/pdf",
                            )
                            try:
                                from pdf_to_excel import pdf_to_excel
                                excel_bytes = pdf_to_excel(pdf_bytes)
                                st.download_button(
                                    f"⬇️ 下載 {order_no} 標籤.xlsx（Excel）",
                                    data=excel_bytes,
                                    file_name=f"標籤_{order_no}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                )
                            except Exception as _xe:
                                import traceback as _tb
                                st.warning(f"PDF 轉 Excel 失敗（仍可下載 PDF）：{_xe}")
                                st.code(_tb.format_exc())
                        else:
                            zip_bytes = pack_zip(success)
                            st.download_button(
                                f"⬇️ 下載所有標籤 ({len(success)} 筆).zip",
                                data=zip_bytes,
                                file_name="ERP標籤.zip",
                                mime="application/zip",
                                use_container_width=True,
                            )

                    if failed:
                        st.warning(f"以下出貨單下載失敗（ERP 端）：{', '.join(failed)}")
                        for _no in failed:
                            if _no in erp_errors:
                                st.code(f"[{_no}] {erp_errors[_no]}")
                    if not success and not failed:
                        st.error("未取得任何 PDF，請確認 ERP 帳密與網路連線")

                except ImportError as e:
                    st.error(f"缺少套件：{e}")
                    import traceback
                    st.code(traceback.format_exc())
                except Exception as e:
                    st.error(f"ERP 下載失敗：{e}")
                    import traceback
                    st.code(traceback.format_exc())


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
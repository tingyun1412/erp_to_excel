"""
出貨自動化工具 v3
模組：
  A - 電子發票產生
  B - 標籤模板管理（上傳舊標籤 → 設定欄位 → 產出新標籤）
"""
import json
import math
import re
import tempfile
from io import BytesIO

import pandas as pd
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
    load_vendors, save_vendor, delete_vendor,
    download_template_excel,
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

    # 頁面一開啟就從 Google Drive 補載所有有 ID 的模板 Excel（不需等有訂單才跑）
    try:
        _all_tpls_preload = load_templates()
        for _tr in _all_tpls_preload:
            _tk = f"{_tr['廠商名稱']}_{_tr['模板名稱']}"
            if _tk not in st.session_state.template_wb_bytes and _tr.get("Excel檔案ID"):
                try:
                    st.session_state.template_wb_bytes[_tk] = download_template_excel(_tr["Excel檔案ID"])
                except Exception:
                    pass
    except Exception:
        pass

    sub_tab_use, sub_tab_manage, sub_tab_erp = st.tabs(["產出標籤", "管理模板", "從廠商網站下載標籤"])

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
                    """廠商名稱比對：先完全包含，再找最長公共子字串 ≥ 3 字"""
                    if not customer_name:
                        return None

                    def _lcs_len(a: str, b: str) -> int:
                        best = 0
                        for i in range(len(a)):
                            for j in range(len(b)):
                                l = 0
                                while i + l < len(a) and j + l < len(b) and a[i + l] == b[j + l]:
                                    l += 1
                                if l > best:
                                    best = l
                        return best

                    # 1. 優先：完全包含（原邏輯）
                    for r in all_templates:
                        vendor = r.get("廠商名稱", "")
                        if vendor and (vendor in customer_name or customer_name in vendor):
                            return r

                    # 2. 次要：最長公共子字串 ≥ 3 字
                    best_rec, best_len = None, 2
                    for r in all_templates:
                        vendor = r.get("廠商名稱", "")
                        if vendor:
                            ln = _lcs_len(customer_name, vendor)
                            if ln > best_len:
                                best_len, best_rec = ln, r
                    return best_rec

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

                    # ── 分裝設定 ─────────────────────────────────────────
                    with st.expander("📦 分裝設定"):
                        _pkg_enable = st.checkbox("啟用分裝", key="pkg_enable")
                        if _pkg_enable:
                            st.caption("填入每箱數量；若每箱數量 ≥ 總數量表示只有一箱，只印一張標籤")
                            _pkg_df = pd.DataFrame([
                                {
                                    "料號":   itm.get("item_no", ""),
                                    "品名":   itm.get("name", ""),
                                    "總數量": int(float(itm.get("quantity") or 0)),
                                    "每箱數量": int(float(itm.get("quantity") or 1)),
                                    "小標籤": True,
                                    "大標籤": True,
                                }
                                for itm, _ in selected_items
                            ])
                            _pkg_edited = st.data_editor(
                                _pkg_df,
                                column_config={
                                    "料號":   st.column_config.TextColumn(disabled=True, width="medium"),
                                    "品名":   st.column_config.TextColumn(disabled=True, width="small"),
                                    "總數量": st.column_config.NumberColumn(disabled=True, width="small"),
                                    "每箱數量": st.column_config.NumberColumn(min_value=1, width="small"),
                                    "小標籤": st.column_config.CheckboxColumn(width="small"),
                                    "大標籤": st.column_config.CheckboxColumn(width="small"),
                                },
                                hide_index=True,
                                use_container_width=True,
                                key="pkg_table",
                            )
                        else:
                            _pkg_edited = None

                else:
                    _pkg_enable = False
                    _pkg_edited = None
                    st.warning("請選擇至少一張銷貨單")

                # 警告：模板無動態欄位 → 標籤只有固定文字
                if selected_orders and order_tmpl_map:
                    for _rec in order_tmpl_map.values():
                        _info = template_from_json(_rec["設定JSON"])
                        _dyn = [c for c in _info.get("cells", []) if c.get("field") != "__fixed__"]
                        if not _dyn:
                            st.error(
                                f"⚠️ 模板「{_rec['廠商名稱']} — {_rec['模板名稱']}」沒有動態欄位，"
                                "標籤將只顯示固定文字（無料號、品名等）。"
                                "請到「管理模板」重新上傳並分析此模板。"
                            )
                            break

                if selected_orders and order_tmpl_map and st.button("產出標籤 Excel", type="primary", use_container_width=True):
                    with st.spinner("產出中..."):
                        try:
                            # 套用分裝：展開品項
                            def _expand_orders(orders, pkg_df):
                                result = []
                                for o in orders:
                                    new_o = dict(o)
                                    new_items = []
                                    for itm in o.get("items", []):
                                        total = float(itm.get("quantity") or 0)
                                        rows = pkg_df[pkg_df["料號"] == itm.get("item_no", "")]
                                        if rows.empty:
                                            new_items.append(dict(itm))
                                            continue
                                        row = rows.iloc[0]
                                        pkg = max(1.0, float(row["每箱數量"] or total))
                                        use_small = bool(row["小標籤"])
                                        use_large = bool(row["大標籤"])
                                        if pkg >= total or total == 0:
                                            # 只有一箱，印一張
                                            new_items.append(dict(itm))
                                        else:
                                            n = math.ceil(total / pkg)
                                            if use_small:
                                                for _ in range(n):
                                                    s = dict(itm)
                                                    s["quantity"] = int(pkg)
                                                    new_items.append(s)
                                            if use_large:
                                                new_items.append(dict(itm))
                                    new_o["items"] = new_items
                                    result.append(new_o)
                                return result

                            _gen_orders = (
                                _expand_orders(selected_orders, _pkg_edited)
                                if _pkg_enable and _pkg_edited is not None
                                else selected_orders
                            )

                            pairs = []
                            for o in _gen_orders:
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
                                save_template(vendor, tname, template_to_json(info), excel_bytes=tmpl_bytes)
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
                    # _inline 變體 (e.g. item_no_inline) → 找基底 key 對應選單
                    _cur_base = current[:-len("_inline")] if current.endswith("_inline") else current
                    # 選項格式：「__fixed__（固定文字）」或「中文名（field_key）」
                    if _cur_base == "__fixed__":
                        curr_display = next(
                            (o for o in field_options if o.startswith("__fixed__")),
                            field_options[0]
                        )
                    else:
                        curr_display = next(
                            (o for o in field_options if f"（{_cur_base}）" in o),
                            field_options[0]
                        )
                    chosen = st.selectbox(
                        "欄位",
                        field_options,
                        index=field_options.index(curr_display) if curr_display in field_options else 0,
                        key=f"field_map_{i}",
                        label_visibility="collapsed",
                    )
                    # 取出內部 field key（括號內），保留 _inline 格式
                    if chosen.startswith("__fixed__"):
                        new_field = "__fixed__"
                    else:
                        _fm = re.search(r'（([^）]+)）', chosen)
                        _base_new = _fm.group(1) if _fm else "__fixed__"
                        # 若原本 inline 且選同一欄位 → 保留 inline；否則看 cell value 有無冒號前綴
                        if current.endswith("_inline") and _base_new == _cur_base:
                            new_field = current
                        elif re.search(r'[：:]\s*$', cell.get("value", "")):
                            new_field = f"{_base_new}_inline"
                        else:
                            new_field = _base_new
                    updated_cell = {**cell, "field": new_field}
                    updated_cells.append(updated_cell)

            if st.button("儲存模板", type="primary", use_container_width=True):
                template_info["cells"] = updated_cells
                config_json = template_to_json(template_info)
                customer = _edit_customer.strip() or st.session_state.get("_pending_customer", "")
                tmpl_name = _edit_tmpl_name.strip() or st.session_state.get("_pending_tmpl_name", "")
                tmpl_key = f"{customer}_{tmpl_name}"

                try:
                    _wb_bytes = st.session_state.get("_pending_tmpl_bytes") or b""
                    save_template(customer, tmpl_name, config_json, excel_bytes=_wb_bytes or None)
                    # 快取 workbook bytes
                    st.session_state.template_wb_bytes[tmpl_key] = _wb_bytes
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
                        if not dynamic:
                            st.warning("⚠️ 無動態欄位，標籤將全為固定文字！請重新上傳 Excel 分析，或手動編輯欄位。")
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
                        # 重新分析（上傳 Excel → 覆寫欄位設定）
                        _tmpl_key = f"{r['廠商名稱']}_{r['模板名稱']}"
                        _re_file = st.file_uploader(
                            "🔄 重新分析（上傳原始 Excel 覆寫欄位）",
                            type=["xlsx", "xls"],
                            key=f"reanalyze_{_tmpl_key}",
                        )
                        if _re_file:
                            _re_bytes = _re_file.read()
                            _re_wb = openpyxl.load_workbook(BytesIO(_re_bytes))
                            _re_sname = info.get("sheet_name", _re_wb.sheetnames[0])
                            if _re_sname not in _re_wb.sheetnames:
                                _re_sname = _re_wb.sheetnames[0]
                            _re_info = analyze_template(_re_wb, _re_sname)
                            if _re_info and _re_info.get("cells"):
                                save_template(r["廠商名稱"], r["模板名稱"], template_to_json(_re_info), excel_bytes=_re_bytes)
                                st.session_state.template_wb_bytes[_tmpl_key] = _re_bytes
                                clear_cache()
                                st.success(f"重新分析完成，找到 {len([c for c in _re_info['cells'] if c['field']!='__fixed__'])} 個動態欄位")
                                st.rerun()
                            else:
                                st.error("分析失敗，此工作表無內容")
        except Exception as e:
            st.error(f"載入失敗：{e}")

    # ── 從廠商網站下載標籤 ───────────────────────────────────────
    with sub_tab_erp:
        st.caption("從廠商 ERP 網站下載標籤 PDF，可從已解析的銷貨單勾選或直接輸入單號")

        # ── 廠商帳號管理 ──────────────────────────────────────────
        _BUILTIN_VENDORS = [
            {"公司名稱": "鴻勁", "網址": "http://scm.honprec.com/hp/Index.aspx",
             "帳號": "BR026", "密碼": "5403"},
        ]

        try:
            _custom_vendors = load_vendors()
        except Exception:
            _custom_vendors = []

        # 合併：內建優先，避免重複
        _builtin_names = {v["公司名稱"] for v in _BUILTIN_VENDORS}
        _all_vendors = _BUILTIN_VENDORS + [v for v in _custom_vendors if v.get("公司名稱") not in _builtin_names]

        if "vendor_selected" not in st.session_state:
            st.session_state.vendor_selected = ""
        if "show_add_vendor" not in st.session_state:
            st.session_state.show_add_vendor = False

        st.markdown("**選擇廠商：**")
        _vcols = st.columns(len(_all_vendors) + 1)
        for _vi, _v in enumerate(_all_vendors):
            with _vcols[_vi]:
                _vtype = "primary" if st.session_state.vendor_selected == _v["公司名稱"] else "secondary"
                if st.button(_v["公司名稱"], key=f"vbtn_{_vi}", type=_vtype, use_container_width=True):
                    st.session_state.vendor_selected = _v["公司名稱"]
                    st.session_state.show_add_vendor = False
                    st.rerun()
        with _vcols[-1]:
            if st.button("＋ 新增廠商", key="add_vendor_btn", use_container_width=True):
                st.session_state.show_add_vendor = not st.session_state.show_add_vendor
                st.rerun()

        # 新增廠商表單
        if st.session_state.show_add_vendor:
            with st.form("add_vendor_form"):
                _fc1, _fc2 = st.columns(2)
                with _fc1:
                    _vname = st.text_input("公司名稱 *")
                    _vurl  = st.text_input("ERP 登入網址 *")
                with _fc2:
                    _vuser = st.text_input("帳號")
                    _vpass = st.text_input("密碼", type="password")
                _fs, _fc = st.columns(2)
                with _fs:
                    _do_save = st.form_submit_button("儲存", type="primary")
                with _fc:
                    _do_cancel = st.form_submit_button("取消")

                if _do_save:
                    if _vname and _vurl:
                        try:
                            save_vendor(_vname, _vurl, _vuser, _vpass)
                            clear_cache()
                            st.session_state.show_add_vendor = False
                            st.session_state.vendor_selected = _vname
                            st.rerun()
                        except Exception as _ve:
                            st.error(f"儲存失敗：{_ve}")
                    else:
                        st.warning("公司名稱和網址為必填")
                if _do_cancel:
                    st.session_state.show_add_vendor = False
                    st.rerun()

        # ── 已選廠商 → 下載介面 ──────────────────────────────────
        _sel_v = st.session_state.vendor_selected
        if _sel_v:
            _vrec = next((v for v in _all_vendors if v["公司名稱"] == _sel_v), None)
            if _vrec:
                # 非內建廠商才顯示刪除
                if _sel_v not in _builtin_names:
                    with st.expander(f"「{_sel_v}」設定"):
                        st.caption(f"網址：{_vrec.get('網址','')}")
                        if st.button(f"🗑 刪除「{_sel_v}」", key="del_vendor_btn"):
                            try:
                                delete_vendor(_sel_v)
                                clear_cache()
                                st.session_state.vendor_selected = ""
                                st.rerun()
                            except Exception as _de:
                                st.error(f"刪除失敗：{_de}")
                else:
                    st.info(f"已選擇：**{_sel_v}**　　{_vrec.get('網址','')}")

                _erp_all_orders = st.session_state.parsed_orders
                _order_nos_parsed = [o.get("order_no", "") for o in _erp_all_orders if o.get("order_no")]

                _input_mode = st.radio(
                    "選擇出貨單方式",
                    ["從已解析銷貨單勾選", "直接輸入銷貨單號"],
                    horizontal=True,
                    key="erp_input_mode",
                )

                _selected_nos: list[str] = []
                if _input_mode == "從已解析銷貨單勾選":
                    if not _order_nos_parsed:
                        st.info("尚無已解析的銷貨單，請先在左側上傳，或改用「直接輸入銷貨單號」")
                    else:
                        _selected_nos = st.multiselect(
                            "選擇要下載的出貨單",
                            options=_order_nos_parsed,
                            default=_order_nos_parsed,
                            key="erp_order_multiselect",
                        )
                else:
                    _raw_text = st.text_area(
                        "輸入銷貨單號（可多筆，用換行或逗號分隔）",
                        placeholder="202606100012\n202606020007",
                        height=100,
                        key="erp_order_text",
                    )
                    _selected_nos = [n.strip() for n in re.split(r"[,\n，、]+", _raw_text) if n.strip()]
                    if _selected_nos:
                        st.caption(f"共 {len(_selected_nos)} 筆")

                if _selected_nos and st.button("從廠商網站下載標籤", type="primary",
                                               use_container_width=True, key="erp_download_btn"):
                    with st.spinner("連線 ERP 並下載中，請稍候..."):
                        try:
                            from erp_downloader import download_label_pdfs, pack_zip
                            _results, _erp_errors = download_label_pdfs(
                                _selected_nos,
                                _vrec.get("帳號", ""),
                                _vrec.get("密碼", ""),
                            )
                            _success = {no: d for no, d in _results.items() if d}
                            _failed  = [no for no, d in _results.items() if not d]

                            if _success:
                                if len(_success) == 1:
                                    _ono, _pdf = next(iter(_success.items()))
                                    st.download_button(
                                        f"⬇️ 下載 {_ono} 標籤.pdf",
                                        data=_pdf,
                                        file_name=f"標籤_{_ono}.pdf",
                                        mime="application/pdf",
                                    )
                                    try:
                                        from pdf_to_excel import pdf_to_excel
                                        _xlsx = pdf_to_excel(_pdf)
                                        st.download_button(
                                            f"⬇️ 下載 {_ono} 標籤.xlsx（縮至8格高）",
                                            data=_xlsx,
                                            file_name=f"標籤_{_ono}.xlsx",
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        )
                                    except Exception as _xe:
                                        st.warning(f"PDF 轉 Excel 失敗（仍可下載 PDF）：{_xe}")
                                else:
                                    _zb = pack_zip(_success)
                                    st.download_button(
                                        f"⬇️ 下載所有標籤 ({len(_success)} 筆).zip",
                                        data=_zb,
                                        file_name="ERP標籤.zip",
                                        mime="application/zip",
                                        use_container_width=True,
                                    )

                            if _failed:
                                st.warning(f"以下出貨單下載失敗：{', '.join(_failed)}")
                                for _fn in _failed:
                                    if _fn in _erp_errors:
                                        st.code(f"[{_fn}] {_erp_errors[_fn]}")
                            if not _success and not _failed:
                                st.error("未取得任何 PDF，請確認帳密與網路連線")

                        except ImportError as _ie:
                            st.error(f"缺少套件：{_ie}")
                        except Exception as _ee:
                            import traceback as _etb
                            st.error(f"ERP 下載失敗：{_ee}")
                            st.code(_etb.format_exc())


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
            _with_inv    = [o for o in selected_orders if o.get("invoice_no")]
            _without_inv = [o for o in selected_orders if not o.get("invoice_no")]

            if _with_inv:
                st.success(f"共 **{len(_with_inv)}** 張有發票號碼，將產出發票")
            if _without_inv:
                _no_inv_nos = [o.get("order_no","(未知)") for o in _without_inv]
                st.warning(f"{len(_without_inv)} 張無發票號碼（略過）：{', '.join(_no_inv_nos)}")

            st.caption("發票人統編：**24405403**　受票人統編：從銷貨單統一編號　單價：暫定 100（待確認）")

            if _with_inv and st.button("產出電子發票 xls", type="primary", use_container_width=True):
                try:
                    buf = generate_invoice_excel(_with_inv)
                    st.download_button(
                        "⬇️ 下載電子發票上傳檔.xls",
                        data=buf,
                        file_name="電子發票上傳.xls",
                        mime="application/vnd.ms-excel",
                        use_container_width=True,
                    )
                except Exception as _inv_e:
                    import traceback as _inv_tb
                    st.error(f"產出失敗：{_inv_e}")
                    st.code(_inv_tb.format_exc())
        else:
            st.warning("請選擇至少一張銷貨單")
"""
出貨自動化工具 v3
模組：
  A - 電子發票產生
  B - 標籤模板管理（上傳舊標籤 → 設定欄位 → 產出新標籤）
"""
import json
import re
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
import openpyxl

from erp_db import lookup_orders_by_no, ErpDbError
from lscr_parser import parse_lscr_excel_wb
from module_b_invoice import (
    parse_acceptance_excel,
    acceptance_rows_to_order,
)
from template_engine import (
    analyze_template, analyze_all_sheets,
    generate_from_template, generate_labels_multiorder,
    template_to_json, template_from_json,
    get_field_options, FIELD_LABELS, DYNAMIC_FIELDS,
    write_lscr_labels,
)
from local_template_store import (
    load_templates, save_template, delete_template,
    download_template_excel,
    find_lscr_base_template_id, save_lscr_base_template, download_lscr_base_template,
)
from local_db import (
    load_vendors, save_vendor, delete_vendor,
    load_invoice_skip_list, save_invoice_skip, delete_invoice_skip,
    load_einvoice_log, save_einvoice_log, delete_einvoice_log, try_claim_order,
    load_shipping_notes, SHIPPING_NOTES_PATH,
    load_label_part_no_prefs, save_label_part_no_pref, delete_label_part_no_pref,
    LABEL_PART_NO_OWN, LABEL_PART_NO_CUSTOMER,
)
from module_d_report import (
    parse_daily_report_workbook,
    aggregate_daily_report,
    build_summary_workbook,
)

st.set_page_config(page_title="出貨自動化工具", page_icon="📦", layout="wide")
st.title("📦 出貨自動化工具")


def _tmpl_label(r: dict) -> str:
    """模板顯示名稱：廠商=模板名稱時只顯示一個，否則顯示『廠商 — 模板』"""
    v, t = r.get("廠商名稱", ""), r.get("模板名稱", "")
    return v if v == t else f"{v} — {t}"


_OLE2_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _read_uploaded_workbook(file_bytes: bytes, **kwargs):
    """
    包住 openpyxl.load_workbook，遇到舊版 .xls（OLE2 格式，不是 zip）時
    丟出清楚的中文錯誤訊息，而不是讓 BadZipFile 的原始 traceback 整個炸出來。
    舊版 .xls 沒辦法無損轉成 .xlsx（樣式/框線會跟著跑掉），所以這裡不自動轉檔，
    只提示使用者自己用 Excel 另存新檔成 .xlsx 再重新上傳。
    """
    if file_bytes[:8] == _OLE2_SIGNATURE:
        raise ValueError(
            "這個檔案是舊版 .xls 格式（Excel 97-2003），無法直接讀取。"
            "請用 Excel 開啟後「另存新檔」成 .xlsx 格式，再重新上傳。"
        )
    return openpyxl.load_workbook(BytesIO(file_bytes), **kwargs)


def _fuzzy_match_customer(customer_name: str, records: list[dict], key: str = "客戶") -> dict | None:
    """
    依客戶名稱比對清單：先完全包含（互為子字串）優先，找不到再用最長公共子字串 ≥ 3 字。
    用於出貨提醒／標籤料號偏好等「依客戶名稱對照設定」的共用比對邏輯。
    """
    if not customer_name or not records:
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

    for r in records:
        cust = r.get(key, "")
        if cust and (cust in customer_name or customer_name in cust):
            return r

    best_rec, best_len = None, 2
    for r in records:
        cust = r.get(key, "")
        if cust:
            ln = _lcs_len(customer_name, cust)
            if ln > best_len:
                best_len, best_rec = ln, r
    return best_rec


if "parsed_orders" not in st.session_state:
    st.session_state.parsed_orders = []
if "template_wb_bytes" not in st.session_state:
    st.session_state.template_wb_bytes = {}  # {template_key: bytes}
if "einv_port" not in st.session_state:
    st.session_state.einv_port = None
if "einv_pid" not in st.session_state:
    st.session_state.einv_pid = None
if "einv_logged_in" not in st.session_state:
    st.session_state.einv_logged_in = False
if "einv_dry_run" not in st.session_state:
    st.session_state.einv_dry_run = True
if "einv_pending_order" not in st.session_state:
    st.session_state.einv_pending_order = None
if "einv_monthly_order" not in st.session_state:
    st.session_state.einv_monthly_order = None



# ════════════════════════════════════════════════════════════════
#  側邊欄：查詢銷貨單（直接查 ERP 資料庫，不用再上傳 RTF）
# ════════════════════════════════════════════════════════════════
with st.sidebar:
    st.header("查詢銷貨單")
    order_nos_text = st.text_area(
        "輸入單據號碼（可多筆，一行一個或用逗號分隔）",
        height=100,
        placeholder="202606010004",
    )

    if st.button("查詢並匯入", type="primary", use_container_width=True):
        order_nos = [x for x in re.split(r"[,\s]+", order_nos_text.strip()) if x]
        if not order_nos:
            st.warning("請輸入至少一個單據號碼")
        else:
            with st.spinner("查詢中..."):
                try:
                    orders, errors = lookup_orders_by_no(order_nos)
                except ErpDbError as e:
                    orders, errors = [], [str(e)]

            for err in errors:
                st.error(err)

            if orders:
                st.session_state.parsed_orders = orders
                st.success(f"查詢完成，共 {len(orders)} 張銷貨單")
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
tab_label, tab_invoice, tab_report = st.tabs([
    "🏷 出貨標籤",
    "🧾 電子發票",
    "📊 報表彙總",
])


# ════════════════════════════════════════════════════════════════
#  標籤模板管理
# ════════════════════════════════════════════════════════════════
with tab_label:
    st.subheader("出貨標籤")

    # 頁面一開啟就從本機補載所有有檔名的模板 Excel（不需等有訂單才跑）
    _preload_missing = []   # 本機根本沒有 Excel檔案ID（之前上傳時就沒存成功）
    _preload_failed  = []   # 有 Excel檔案ID 但這次讀檔失敗（檔案被搬走/刪除）
    try:
        _all_tpls_preload = load_templates()
        for _tr in _all_tpls_preload:
            _tk = f"{_tr['廠商名稱']}_{_tr['模板名稱']}"
            if _tk in st.session_state.template_wb_bytes:
                continue
            if not _tr.get("Excel檔案ID"):
                _preload_missing.append(_tmpl_label(_tr))
                continue
            try:
                st.session_state.template_wb_bytes[_tk] = download_template_excel(_tr["Excel檔案ID"])
            except Exception:
                _preload_failed.append(_tmpl_label(_tr))
    except Exception:
        pass

    if _preload_missing or _preload_failed:
        with st.expander(f"⚠️ {len(_preload_missing) + len(_preload_failed)} 個模板需要重新上傳原始 Excel", expanded=False):
            if _preload_missing:
                st.caption("尚未存到本機（請到「管理模板」重新分析補傳一次即可永久解決）：" + "、".join(_preload_missing))
            if _preload_failed:
                st.caption("本機檔案讀取失敗（檔案可能被搬走或刪除）：" + "、".join(_preload_failed))

    sub_tab_use, sub_tab_manage, sub_tab_erp, sub_tab_lscr = st.tabs(["產出標籤", "管理模板", "從廠商網站下載標籤", "LSCR 確認單"])

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
                # 標籤格式已統一，固定使用單一預設模板；若還留有多個模板才需手動指定
                if len(all_templates) == 1:
                    default_tmpl = all_templates[0]
                else:
                    st.caption("偵測到多個標籤模板，請選擇本次要使用的預設模板：")
                    _tmpl_idx = st.selectbox(
                        "預設模板",
                        options=range(len(all_templates)),
                        format_func=lambda i: _tmpl_label(all_templates[i]),
                        label_visibility="collapsed",
                    )
                    default_tmpl = all_templates[_tmpl_idx]

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

                # 出貨提醒：依客戶名稱比對
                try:
                    _ship_notes = load_shipping_notes()
                except Exception:
                    _ship_notes = []

                def _match_shipping_note(customer_name: str):
                    return _fuzzy_match_customer(customer_name, _ship_notes)

                # 標籤料號偏好：依客戶名稱比對，決定該客戶要印本公司料號還是客戶料號
                _part_no_prefs = load_label_part_no_prefs()

                def _match_part_no_pref(customer_name: str) -> str:
                    _rec = _fuzzy_match_customer(customer_name, _part_no_prefs)
                    return _rec.get("料號來源", LABEL_PART_NO_OWN) if _rec else LABEL_PART_NO_OWN

                # 顯示選取的品項：每張銷貨單自己一段，出貨提醒當該單的標題只顯示一次，
                # 底下接該單自己的品項表（不用把出貨提醒塞進表格欄位裡重複出現）
                selected_items = [
                    (item, order)
                    for order in selected_orders
                    for item in order.get("items", [])
                ]
                if selected_items:
                    for order in selected_orders:
                        _items = order.get("items", [])
                        if not _items:
                            continue
                        _note = _match_shipping_note(order.get("customer_name", ""))
                        st.markdown(f"**{order.get('order_no','')}** — {order.get('customer_name','')}")
                        if _note:
                            _req = _note.get("出貨要求", "").strip()
                            _remark = _note.get("備註", "").strip()
                            if _req or _remark:
                                # st.dataframe 長文字會被裁切看不到，這裡只有一列資料，
                                # 改用 st.table──它會整格顯示完整內容、自動換行，不會裁切。
                                st.table(pd.DataFrame([{"出貨要求": _req, "備註": _remark}]).set_index(pd.Index([""])))
                        st.dataframe(
                            [{
                                "料號":     it.get("item_no", ""),
                                "品名":     it.get("name", ""),
                                "規格":     it.get("description", ""),
                                "數量":     it.get("quantity", ""),
                                "客戶料號": it.get("remark", ""),
                                "批號":     it.get("lot_no", ""),
                            } for it in _items],
                            use_container_width=True,
                            hide_index=True,
                            height=min(420, 38 * (len(_items) + 1) + 10),
                            column_config={
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
                            st.caption(
                                "填入每箱數量；若每箱數量 ≥ 總數量表示只有一箱，只印一張標籤。"
                                "否則固定印出 2 張小標籤（顯示每箱數量）+ 1 張大標籤（顯示總數量）並排，"
                                "不會依箱數自動算張數——實際要印幾份由印標籤的人自己決定。"
                            )
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
                if selected_orders and default_tmpl:
                    _info = template_from_json(default_tmpl["設定JSON"])
                    _dyn = [c for c in _info.get("cells", []) if c.get("field") != "__fixed__"]
                    if not _dyn:
                        st.error(
                            f"⚠️ 模板「{_tmpl_label(default_tmpl)}」沒有動態欄位，"
                            "標籤將只顯示固定文字（無料號、品名等）。"
                            "請到「管理模板」重新上傳並分析此模板。"
                        )

                if selected_orders and default_tmpl and st.button("產出標籤 Excel", type="primary", use_container_width=True):
                    with st.spinner("產出中..."):
                        try:
                            # 套用分裝：展開品項
                            def _expand_orders(orders, pkg_df):
                                result = []
                                _pkg_gid = [0]
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
                                            # 固定 2 張小標籤（每箱數量）+ 1 張大標籤（總數量）並排在同一列，
                                            # 不論每箱數量填多少都印 3 張——箱數由印標籤的人自己決定。
                                            # _pkg_group 標記讓產出時強制排在同一列，不受範本並排數限制。
                                            _pkg_gid[0] += 1
                                            gid = _pkg_gid[0]
                                            if use_small:
                                                s = dict(itm)
                                                s["quantity"] = int(pkg)
                                                s["_pkg_group"] = gid
                                                new_items.append(s)
                                                new_items.append(dict(s))
                                            if use_large:
                                                l = dict(itm)
                                                l["_pkg_group"] = gid
                                                new_items.append(l)
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
                                rec = default_tmpl
                                tmpl_key = f"{rec['廠商名稱']}_{rec['模板名稱']}"
                                wb_bytes = st.session_state.template_wb_bytes.get(tmpl_key)
                                twb = (
                                    openpyxl.load_workbook(BytesIO(wb_bytes))
                                    if wb_bytes else openpyxl.Workbook()
                                )
                                _tinfo = template_from_json(rec["設定JSON"])
                                if _match_part_no_pref(o.get("customer_name", "")) == LABEL_PART_NO_CUSTOMER:
                                    # 該客戶偏好印客戶料號：只把「料號」欄位改指到 remark，
                                    # 不動本來就顯示 remark（客戶料號）的其他欄位，避免誤改
                                    _tinfo = json.loads(json.dumps(_tinfo))  # 深複製，不動到共用模板設定
                                    for _c in _tinfo.get("cells", []):
                                        if _c.get("field") == "item_no":
                                            _c["field"] = "remark"
                                pairs.append({
                                    "order": o,
                                    "template_info": _tinfo,
                                    "template_wb": twb,
                                    "template_bytes": wb_bytes,
                                    "vendor": rec.get("廠商名稱", ""),
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
                if selected_orders and default_tmpl:
                    _tk = f"{default_tmpl['廠商名稱']}_{default_tmpl['模板名稱']}"
                    if _tk not in st.session_state.template_wb_bytes:
                        _re = st.file_uploader(
                            f"上傳「{_tmpl_label(default_tmpl)}」原始 Excel（保留樣式用）",
                            type=["xlsx", "xls"],
                            key=f"reupload_{_tk}",
                        )
                        if _re:
                            _re_raw = _re.read()
                            try:
                                _read_uploaded_workbook(_re_raw)  # 只是驗證格式，丟例外就不存進去
                            except ValueError as _ve:
                                st.error(str(_ve))
                            else:
                                st.session_state.template_wb_bytes[_tk] = _re_raw
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
            try:
                wb = _read_uploaded_workbook(tmpl_bytes)
            except ValueError as _ve:
                wb = None
                st.error(str(_ve))

        if uploaded_tmpl and wb is not None:
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
                        sync_errors = []
                        first_sync_err = None
                        for sname, info in results.items():
                            try:
                                vendor = new_customer.strip() or sname
                                tname = new_tmpl_name.strip() or sname
                                _err = save_template(vendor, tname, template_to_json(info), excel_bytes=tmpl_bytes)
                                if _err:
                                    sync_errors.append(f"{vendor} — {tname}")
                                    first_sync_err = first_sync_err or _err
                                tmpl_key = f"{vendor}_{tname}"
                                st.session_state.template_wb_bytes[tmpl_key] = tmpl_bytes
                                saved += 1
                            except Exception:
                                skipped += 1
                        st.success(f"完成！成功建立 {saved} 個模板" + (f"，{skipped} 個失敗" if skipped else ""))
                        if sync_errors:
                            st.warning(
                                "⚠️ 以下模板的原始 Excel 未能存到本機（僅存在目前分頁的暫存中）："
                                + "、".join(sync_errors)
                                + "。下次重新整理或換裝置時會需要重新上傳。"
                                f"\n\n錯誤訊息：{first_sync_err}"
                            )
                        else:
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
                        st.session_state["_pending_is_edit"] = False

        # 欄位對應確認
        if "_pending_template" in st.session_state:
            template_info = st.session_state["_pending_template"]
            cells = template_info.get("cells", [])

            st.divider()
            st.markdown("### 確認欄位對應")
            st.caption("檢查每個格子的欄位是否正確，可以修改")

            # 名稱可編輯（新增時預填，編輯時可改名）。
            # widget key 帶入原始名稱，確保切換到編輯「另一個」模板時輸入框會重置成新模板的名稱，
            # 不會沿用上一個模板編輯時殘留的文字（否則存檔時會誤判成新模板，另外新增一列）。
            _pending_key_suffix = (
                f"{st.session_state.get('_pending_customer','')}_{st.session_state.get('_pending_tmpl_name','')}"
            )
            _ec1, _ec2 = st.columns(2)
            with _ec1:
                _edit_customer = st.text_input(
                    "廠商名稱（用於比對）",
                    value=st.session_state.get("_pending_customer", ""),
                    key=f"_edit_customer_input_{_pending_key_suffix}",
                )
            with _ec2:
                _edit_tmpl_name = st.text_input(
                    "模板名稱",
                    value=st.session_state.get("_pending_tmpl_name", ""),
                    key=f"_edit_tmpl_name_input_{_pending_key_suffix}",
                )

            # 標籤列數（自動偵測，可手動修正）
            _auto_unit_rows = template_info.get("unit_rows", 1)
            _ec_rows, _ec_cols = st.columns(2)
            with _ec_rows:
                _edit_unit_rows = st.number_input(
                    f"每個標籤佔幾列（自動偵測={_auto_unit_rows}，模板有多份樣本時請填單份列數）",
                    min_value=1,
                    value=_auto_unit_rows,
                    step=1,
                    key="_edit_unit_rows_input",
                )
            with _ec_cols:
                st.metric("並排數（自動）", template_info.get("units_per_row", 1))

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
                template_info["unit_rows"] = int(_edit_unit_rows)
                config_json = template_to_json(template_info)
                customer = _edit_customer.strip() or st.session_state.get("_pending_customer", "")
                tmpl_name = _edit_tmpl_name.strip() or st.session_state.get("_pending_tmpl_name", "")
                tmpl_key = f"{customer}_{tmpl_name}"
                _is_edit = st.session_state.get("_pending_is_edit", False)
                _orig_customer = st.session_state.get("_pending_customer", "") if _is_edit else None
                _orig_tmpl_name = st.session_state.get("_pending_tmpl_name", "") if _is_edit else None

                try:
                    _wb_bytes = st.session_state.get("_pending_tmpl_bytes") or b""
                    _sync_err = save_template(
                        customer, tmpl_name, config_json, excel_bytes=_wb_bytes or None,
                        original_customer=_orig_customer, original_template_name=_orig_tmpl_name,
                    )
                    # 快取 workbook bytes
                    st.session_state.template_wb_bytes[tmpl_key] = _wb_bytes
                    # 清除暫存
                    del st.session_state["_pending_template"]
                    del st.session_state["_pending_tmpl_bytes"]
                    del st.session_state["_pending_customer"]
                    del st.session_state["_pending_tmpl_name"]
                    st.session_state.pop("_pending_is_edit", None)
                    st.success(f"模板「{customer} — {tmpl_name}」已儲存！")
                    if _sync_err:
                        st.warning(
                            f"⚠️ 原始 Excel 未能存到本機，重新整理頁面或換裝置後會需要重新上傳。\n\n"
                            f"錯誤訊息：{_sync_err}"
                        )
                    else:
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
                    with st.expander(f"{_tmpl_label(r)}　（更新：{r.get('最後更新','')}）"):
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
                                st.session_state["_pending_is_edit"] = True
                                st.rerun()
                        with _btn_del:
                            if st.button("🗑 刪除此模板", key=f"del_{r['廠商名稱']}_{r['模板名稱']}"):
                                try:
                                    delete_template(r["廠商名稱"], r["模板名稱"])
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
                            try:
                                _re_wb = _read_uploaded_workbook(_re_bytes)
                            except ValueError as _ve:
                                _re_wb = None
                                st.error(str(_ve))
                        if _re_file and _re_wb is not None:
                            _re_sname = info.get("sheet_name", _re_wb.sheetnames[0])
                            if _re_sname not in _re_wb.sheetnames:
                                _re_sname = _re_wb.sheetnames[0]
                            _re_info = analyze_template(_re_wb, _re_sname)
                            if _re_info and _re_info.get("cells"):
                                _re_err = save_template(r["廠商名稱"], r["模板名稱"], template_to_json(_re_info), excel_bytes=_re_bytes)
                                st.session_state.template_wb_bytes[_tmpl_key] = _re_bytes
                                st.success(f"重新分析完成，找到 {len([c for c in _re_info['cells'] if c['field']!='__fixed__'])} 個動態欄位")
                                if _re_err:
                                    st.warning(
                                        f"⚠️ 原始 Excel 未能存到本機，下次重新整理可能又要重傳一次。\n\n"
                                        f"錯誤訊息：{_re_err}"
                                    )
                                else:
                                    st.rerun()
                            else:
                                st.error("分析失敗，此工作表無內容")
        except Exception as e:
            st.error(f"載入失敗：{e}")

        st.divider()
        st.markdown("### 出貨提醒設定")
        st.caption(
            "依客戶名稱比對，銷貨單解析後會在「產出標籤」顯示對應的出貨要求／備註。"
            f"直接即時讀取倉管維護的共用檔案（`{SHIPPING_NOTES_PATH}` 的「出貨要求」工作表），"
            "不需要另外匯入，倉管更新那份檔案後這裡會自動反映最新內容。"
        )
        try:
            _ship_notes_cur = load_shipping_notes()
            st.dataframe(
                pd.DataFrame(_ship_notes_cur) if _ship_notes_cur else pd.DataFrame(columns=["客戶", "出貨要求", "備註"]),
                use_container_width=True,
                hide_index=True,
            )
        except Exception as _se2:
            st.error(f"讀取出貨提醒失敗：{_se2}")

        st.divider()
        st.markdown("### 標籤料號偏好設定")
        st.caption(
            "標籤模板已改成全公司共用一份，「料號」欄位預設印本公司料號；"
            "依客戶名稱比對，這裡設定過的客戶會改印客戶料號，設定一次之後每次產出標籤都會自動套用。"
        )
        _part_no_prefs_cur = load_label_part_no_prefs()
        if _part_no_prefs_cur:
            st.dataframe(
                pd.DataFrame(_part_no_prefs_cur),
                use_container_width=True,
                hide_index=True,
            )
            _pn_del_target = st.selectbox(
                "選擇要刪除的客戶（刪除後該客戶恢復印本公司料號）",
                [""] + [r.get("客戶", "") for r in _part_no_prefs_cur],
                key="part_no_del_sel",
            )
            if _pn_del_target and st.button("刪除選定客戶", key="part_no_del_btn"):
                delete_label_part_no_pref(_pn_del_target)
                st.rerun()
        else:
            st.caption("目前沒有設定，所有客戶都印本公司料號")

        _pn_c1, _pn_c2, _pn_c3 = st.columns([2, 2, 1])
        with _pn_c1:
            _pn_new_customer = st.text_input(
                "客戶名稱（比對銷貨單客戶名稱，包含即算符合）", key="part_no_new_customer")
        with _pn_c2:
            _pn_new_choice = st.selectbox(
                "料號來源", [LABEL_PART_NO_CUSTOMER, LABEL_PART_NO_OWN], key="part_no_new_choice")
        with _pn_c3:
            st.write("")
            st.write("")
            if st.button("新增／更新", key="part_no_add_btn", use_container_width=True) and _pn_new_customer.strip():
                save_label_part_no_pref(_pn_new_customer.strip(), _pn_new_choice)
                st.rerun()

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
                        st.info("尚無已查詢的銷貨單，請先在左側查詢，或改用「直接輸入銷貨單號」")
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

                            def _pdf_to_combined_b64(pdf_bytes: bytes, dpi: int = 150) -> str:
                                """把 PDF 所有頁面垂直合併成一張 PNG，回傳 base64"""
                                import fitz as _fitz
                                import base64 as _b64
                                from PIL import Image as _PILImage
                                _doc = _fitz.open(stream=pdf_bytes, filetype="pdf")
                                _imgs = [
                                    _PILImage.frombytes(
                                        "RGB",
                                        [_p.width, _p.height],
                                        _p.samples,
                                    )
                                    for _p in (_doc[i].get_pixmap(dpi=dpi) for i in range(len(_doc)))
                                ]
                                _w = max(im.width for im in _imgs)
                                _h = sum(im.height for im in _imgs)
                                _combined = _PILImage.new("RGB", (_w, _h), "white")
                                _y = 0
                                for im in _imgs:
                                    _combined.paste(im, (0, _y))
                                    _y += im.height
                                _buf = BytesIO()
                                _combined.save(_buf, format="PNG")
                                return _b64.b64encode(_buf.getvalue()).decode()

                            def _copy_button_html(b64: str, btn_id: str, label: str) -> str:
                                return f"""
<button id="{btn_id}" onclick="copyLabel_{btn_id}()" style="
    background:#0068c9;color:white;border:none;border-radius:6px;
    padding:8px 0;font-size:15px;cursor:pointer;width:100%;margin-top:4px">
    {label}
</button>
<div id="toast_{btn_id}" style="display:none;margin-top:6px;padding:8px;
    background:#21c354;color:white;border-radius:6px;text-align:center;font-size:14px">
    ✓ 已複製！可直接貼上到 Codex
</div>
<script>
async function copyLabel_{btn_id}(){{
    try{{
        const blob=await fetch('data:image/png;base64,{b64}').then(r=>r.blob());
        await navigator.clipboard.write([new ClipboardItem({{'image/png':blob}})]);
        document.getElementById('toast_{btn_id}').style.display='block';
        document.getElementById('{btn_id}').textContent='✓ 已複製';
        document.getElementById('{btn_id}').style.background='#21c354';
        setTimeout(()=>{{
            document.getElementById('toast_{btn_id}').style.display='none';
            document.getElementById('{btn_id}').textContent='{label}';
            document.getElementById('{btn_id}').style.background='#0068c9';
        }},2500);
    }}catch(e){{alert('複製失敗：'+e.message);}}
}}
</script>"""

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
                                        _b64str = _pdf_to_combined_b64(_pdf)
                                        st.components.v1.html(
                                            _copy_button_html(_b64str, "cp_single", "複製截圖"),
                                            height=100,
                                        )
                                        st.image(
                                            BytesIO(__import__('base64').b64decode(_b64str)),
                                            caption=f"標籤預覽（全頁）：{_ono}",
                                            use_container_width=True,
                                        )
                                    except Exception as _pe:
                                        st.warning(f"無法產生預覽：{_pe}")
                                else:
                                    _zb = pack_zip(_success)
                                    st.download_button(
                                        f"⬇️ 下載所有標籤 ({len(_success)} 筆).zip",
                                        data=_zb,
                                        file_name="ERP標籤.zip",
                                        mime="application/zip",
                                        use_container_width=True,
                                    )
                                    try:
                                        for _mno, _mpdf in _success.items():
                                            _mb64 = _pdf_to_combined_b64(_mpdf)
                                            _bid = f"cp_{_mno.replace('-','_')}"
                                            st.components.v1.html(
                                                _copy_button_html(_mb64, _bid, f"複製截圖（{_mno}）"),
                                                height=100,
                                            )
                                            st.image(
                                                BytesIO(__import__('base64').b64decode(_mb64)),
                                                caption=f"標籤預覽（全頁）：{_mno}",
                                                use_container_width=True,
                                            )
                                    except Exception as _mpe:
                                        st.warning(f"無法產生預覽：{_mpe}")

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


    # ── LSCR 確認單直接產出 ────────────────────────────────────────
    with sub_tab_lscr:
        st.caption("上傳 LSCR 出貨明細確認單（xlsx），自動解析明細並用內建 lable 工作表產出標籤"
                   "（若檔案只有 list 沒有 lable，會自動沿用本機的預設模板0）")

        _lscr_up = st.file_uploader(
            "上傳 LSCR 確認單 xlsx",
            type=["xlsx"],
            key="lscr_up",
        )

        if _lscr_up:
            _lscr_bytes = _lscr_up.read()
            try:
                _wb_data = openpyxl.load_workbook(BytesIO(_lscr_bytes), data_only=True)

                if "list" not in _wb_data.sheetnames:
                    st.error("此檔案缺少工作表：list")
                    _tmpl_bytes = None
                elif "lable" in _wb_data.sheetnames:
                    # 這次上傳的檔案自帶 lable：直接用它，並在本機還沒有模板0 時存一份
                    # （只存第一次，之後不再覆蓋，避免之後上傳的檔案版型跑掉時把模板0 也帶壞）
                    _tmpl_bytes = _lscr_bytes
                    try:
                        if not find_lscr_base_template_id():
                            save_lscr_base_template(_lscr_bytes)
                            st.info("已將本次的 lable 版型另存為本機預設模板0，"
                                    "之後上傳只有 list 的檔案會自動沿用。")
                    except Exception as _sav_e:
                        st.warning(f"本機模板0 儲存失敗（不影響本次產出）：{_sav_e}")
                else:
                    # 檔案只有 list 沒有 lable：沿用本機的預設模板0
                    _base_bytes = download_lscr_base_template()
                    if _base_bytes:
                        _tmpl_bytes = _base_bytes
                        st.caption("此檔案沒有 lable 工作表，已自動沿用本機預設模板0 排版")
                    else:
                        st.error("此檔案缺少工作表：lable，且本機尚無預設模板0 可沿用"
                                 "（請先上傳一份含 lable 工作表的檔案，之後才能沿用）")
                        _tmpl_bytes = None

                if _tmpl_bytes:
                    _wb_tmpl = openpyxl.load_workbook(BytesIO(_tmpl_bytes))
                    _lscr_orders = parse_lscr_excel_wb(_wb_data)
                    _lscr_raw_items = [
                        (itm, o)
                        for o in _lscr_orders
                        for itm in o.get("items", [])
                    ]

                    st.success(f"解析完成：{len(_lscr_orders)} 張訂單，共 {len(_lscr_raw_items)} 個品項")

                    # 品項預覽
                    st.dataframe(
                        [{
                            "PO NO":    o.get("order_no", ""),
                            "料號":      itm.get("item_no", ""),
                            "品名":      itm.get("name", ""),
                            "規格":      itm.get("description", ""),
                            "總數量":    f"{itm.get('_total_qty','')}{itm.get('_large_unit','PCS')}",
                            "大包裝":    f"{itm.get('_large_qty','')}{itm.get('_large_unit','PCS')}",
                            "小包裝":    f"{itm.get('_small_qty','')}{itm.get('_small_unit','PCS')}",
                            "LOT NO":   itm.get("lot_no", ""),
                        } for itm, o in _lscr_raw_items],
                        use_container_width=True,
                        hide_index=True,
                        height=min(420, 38 * (len(_lscr_raw_items) + 1) + 10),
                        column_config={
                            "PO NO":  st.column_config.TextColumn(width="medium"),
                            "料號":    st.column_config.TextColumn(width="large"),
                            "品名":    st.column_config.TextColumn(width="medium"),
                            "規格":    st.column_config.TextColumn(width="large"),
                            "總數量":  st.column_config.TextColumn(width="small"),
                            "大包裝":  st.column_config.TextColumn(width="small"),
                            "小包裝":  st.column_config.TextColumn(width="small"),
                            "LOT NO": st.column_config.TextColumn(width="medium"),
                        },
                    )

                    # 分裝設定
                    st.markdown("**標籤設定**")
                    _lscr_c1, _lscr_c2 = st.columns(2)
                    with _lscr_c1:
                        _lscr_small = st.checkbox("印小標籤（小包裝數量）", value=True, key="lscr_small")
                    with _lscr_c2:
                        _lscr_large = st.checkbox("印大標籤（總出貨數量）", value=True, key="lscr_large")

                    if st.button("產出標籤 Excel", type="primary",
                                 use_container_width=True, key="lscr_gen"):
                        with st.spinner("產出中..."):
                            try:
                                _lscr_tmpl_info = analyze_template(_wb_tmpl, "lable")
                                if not _lscr_tmpl_info or not _lscr_tmpl_info.get("cells"):
                                    st.error("無法分析 lable 工作表")
                                else:
                                    _buf = write_lscr_labels(
                                        _lscr_orders,
                                        openpyxl.load_workbook(BytesIO(_tmpl_bytes)),
                                        _lscr_tmpl_info,
                                        include_small=_lscr_small,
                                        include_large=_lscr_large,
                                        tmpl_bytes=_tmpl_bytes,
                                    )
                                    st.download_button(
                                        "⬇️ 下載標籤.xlsx",
                                        data=_buf,
                                        file_name="LSCR標籤.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        use_container_width=True,
                                    )
                            except Exception as _le:
                                import traceback as _ltb
                                st.error(f"產出失敗：{_le}")
                                st.code(_ltb.format_exc())

            except Exception as _lscr_e:
                st.error(f"解析失敗：{_lscr_e}")
                import traceback as _ltb2
                st.code(_ltb2.format_exc())


# ════════════════════════════════════════════════════════════════
#  電子發票
# ════════════════════════════════════════════════════════════════
with tab_invoice:
    st.subheader("電子發票")
    st.caption("逐張自動送到 e-invoice.com.tw 開立；月結先在下方彙整成一張再一起送出")

    with st.expander("📅 月結（驗收資訊）上傳", expanded=bool(st.session_state.einv_monthly_order)):
        st.caption("上傳從 ERP 匯出的「驗收資訊」xlsx，整批彙整成一張發票，一樣要逐項填進網站、由人工按下開立發票")

        if st.session_state.einv_monthly_order:
            _mo = st.session_state.einv_monthly_order
            _moc1, _moc2 = st.columns([4, 1])
            with _moc1:
                st.info(
                    f"待送出：**{_mo['order_no']}** — {_mo['customer_name']}"
                    f"（{len(_mo['items'])} 項，統編 {_mo['buyer_tax_id']}）"
                )
            with _moc2:
                if st.button("移除", key="acc_remove_btn", use_container_width=True):
                    st.session_state.einv_monthly_order = None
                    st.rerun()
        else:
            _acc_up = st.file_uploader("上傳驗收資訊 xlsx", type=["xlsx"], key="acc_up")

            if _acc_up:
                _acc_bytes = _acc_up.read()
                try:
                    _acc_rows = parse_acceptance_excel(_acc_bytes)
                except Exception as _ae:
                    st.error(f"解析失敗：{_ae}")
                    _acc_rows = []

                if _acc_rows:
                    _acc_total = sum(r["amount"] for r in _acc_rows)
                    st.success(f"解析完成：**{len(_acc_rows)}** 行，未稅合計 **{_acc_total:,.0f}**，"
                               f"稅額 **{round(_acc_total*0.05):,.0f}**，"
                               f"含稅 **{round(_acc_total*1.05):,.0f}**")

                    st.dataframe(
                        [{
                            "出貨單號": r["order_no"],
                            "單號":     r["line_no"],
                            "品號":     r["part_no"],
                            "品名":     r["name"],
                            "規格":     r["spec"],
                            "數量":     r["qty"],
                            "單價":     r["unit_price"],
                            "金額(未稅)": r["amount"],
                        } for r in _acc_rows],
                        use_container_width=True,
                        hide_index=True,
                        height=min(400, 38 * (len(_acc_rows) + 1) + 10),
                    )

                    _mc1, _mc2, _mc3 = st.columns(3)
                    with _mc1:
                        _acc_customer = st.text_input("客戶名稱", key="acc_customer")
                    with _mc2:
                        _acc_buyer_id = st.text_input("受票人統編", placeholder="12345678", key="acc_buyer_id")
                    with _mc3:
                        _acc_order_no = st.text_input(
                            "發票識別號（自訂，用於記錄與避免重複送出）",
                            value=f"月結_{Path(_acc_up.name).stem}", key="acc_order_no",
                        )

                    st.caption("發票人統編：**24405403**")

                    _can_add = bool(_acc_customer.strip() and _acc_buyer_id.strip() and _acc_order_no.strip())
                    if not _can_add:
                        st.warning("請填入客戶名稱、受票人統編、發票識別號")

                    if _can_add and st.button("加入送出清單", type="primary",
                                              use_container_width=True, key="acc_add_btn"):
                        st.session_state.einv_monthly_order = acceptance_rows_to_order(
                            _acc_rows,
                            order_no=_acc_order_no.strip(),
                            customer_name=_acc_customer.strip(),
                            buyer_tax_id=_acc_buyer_id.strip(),
                        )
                        st.rerun()

    all_orders = list(st.session_state.parsed_orders)
    if st.session_state.einv_monthly_order:
        all_orders.append(st.session_state.einv_monthly_order)

    if not all_orders:
        st.info("請先在左側上傳並解析銷貨單，或在上方加入月結彙整發票")
    else:
        order_map = {
            f"{o.get('order_no','')} — {o.get('customer_name') or o.get('customer_code','')} "
            f"({len(o.get('items',[]))} 項)": o
            for o in all_orders
        }
        selected_keys = st.multiselect(
            "選擇要開發票的銷貨單",
            options=list(order_map.keys()),
            default=list(order_map.keys()),
        )
        selected_orders = [order_map[k] for k in selected_keys if order_map[k].get("items")]

        # ── 跳過名單管理（例如：不開立發票的客戶、月結另外處理的客戶）──
        with st.expander("⚙️ 發票跳過名單管理"):
            _skip_list = load_invoice_skip_list()
            if _skip_list:
                st.dataframe(
                    [{"客戶名稱": r.get("客戶名稱",""), "原因": r.get("原因","")} for r in _skip_list],
                    use_container_width=True, hide_index=True,
                )
                _del_target = st.selectbox(
                    "選擇要刪除的客戶",
                    [""] + [r.get("客戶名稱","") for r in _skip_list],
                    key="inv_skip_del_sel",
                )
                if _del_target and st.button("刪除選定客戶", key="inv_skip_del_btn"):
                    delete_invoice_skip(_del_target)
                    st.rerun()
            else:
                st.caption("目前沒有跳過名單")

            _sc1, _sc2, _sc3 = st.columns([2, 1, 1])
            with _sc1:
                _new_skip_name = st.text_input(
                    "客戶名稱（比對銷貨單客戶名稱，包含即算符合）", key="inv_skip_new_name")
            with _sc2:
                _new_skip_reason = st.text_input(
                    "原因", value="不開立", key="inv_skip_new_reason")
            with _sc3:
                st.write("")
                st.write("")
                if st.button("新增", key="inv_skip_add_btn", use_container_width=True) and _new_skip_name.strip():
                    save_invoice_skip(_new_skip_name.strip(), _new_skip_reason.strip() or "不開立")
                    st.rerun()

        def _match_skip(customer: str) -> str | None:
            cn = (customer or "").strip()
            if not cn:
                return None
            for r in _skip_list:
                sn = r.get("客戶名稱", "").strip()
                if sn and (sn in cn or cn in sn):
                    return r.get("原因") or "跳過"
            return None

        _skipped = []       # (order, reason)
        _eligible = []
        for o in selected_orders:
            reason = _match_skip(o.get("customer_name") or o.get("customer_code"))
            if reason:
                _skipped.append((o, reason))
            else:
                _eligible.append(o)

        if not selected_orders:
            st.warning("請選擇至少一張銷貨單")
        else:
            if _skipped:
                _skip_lines = [
                    f"{o.get('order_no','(未知)')}／{o.get('customer_name') or o.get('customer_code','')}（{reason}）"
                    for o, reason in _skipped
                ]
                st.warning(f"{len(_skipped)} 張因客戶在跳過名單中而略過：" + "、".join(_skip_lines))

            # ── 已開立過／正在處理中的（依銷貨單號查記錄）從送出清單濾掉，避免重複送出 ──
            _einv_log = load_einvoice_log()
            _excluded_map = {
                r.get("銷貨單號"): r
                for r in _einv_log if r.get("狀態") in ("已開立", "處理中")
            }
            _already_excluded = [o for o in _eligible if o.get("order_no") in _excluded_map]
            _not_yet = [o for o in _eligible if o.get("order_no") not in _excluded_map]

            if _already_excluded:
                with st.expander(f"⏸️ {len(_already_excluded)} 張已開立或正在處理中（不會重複送出）"):
                    st.dataframe(
                        [{"銷貨單號": o.get("order_no", ""),
                          "客戶": o.get("customer_name") or o.get("customer_code", ""),
                          "狀態": _excluded_map.get(o.get("order_no"), {}).get("狀態", ""),
                          "發票號碼": _excluded_map.get(o.get("order_no"), {}).get("發票號碼", "")}
                         for o in _already_excluded],
                        use_container_width=True, hide_index=True,
                    )
                    _reset_no = st.selectbox(
                        "若該筆發票已在網站上作廢、或「處理中」卡住了，可清除記錄讓它重新送出",
                        [""] + [o.get("order_no", "") for o in _already_excluded],
                        key="einv_log_reset_sel",
                    )
                    if _reset_no and st.button("清除記錄", key="einv_log_reset_btn"):
                        delete_einvoice_log(_reset_no)
                        st.rerun()

            import einvoice_submitter as _einv

            _general_orders, _core_orders = [], []
            for o in _not_yet:
                _mtype = _einv.resolve_member_type(o.get("customer_name") or o.get("customer_code", ""))
                (_core_orders if _mtype == "core" else _general_orders).append(o)

            if _core_orders:
                _core_lines = [f"{o.get('order_no','(未知)')}／{o.get('customer_name','')}" for o in _core_orders]
                st.warning("核心會員（日月光）發票流程尚未支援自動送出，請人工於網站開立：" + "、".join(_core_lines))

            st.caption("發票人統編：**24405403**　受票人統編：從銷貨單統一編號")

            st.divider()
            st.markdown("### 登入 e-invoice.com.tw")
            st.caption("帳密已自動帶入，只需要在跳出的視窗輸入圖形驗證碼後按登入")

            _lc1, _lc2, _lc3 = st.columns(3)
            with _lc1:
                if st.button("開啟登入視窗", key="einv_login_btn", use_container_width=True):
                    _res = _einv.launch_login_browser()
                    st.session_state.einv_port = _res["port"]
                    st.session_state.einv_pid = _res["pid"]
                    st.session_state.einv_logged_in = False
                    st.rerun()
            with _lc2:
                if st.session_state.einv_port and st.button(
                    "我已完成登入，檢查連線", key="einv_check_btn", use_container_width=True
                ):
                    st.session_state.einv_logged_in = _einv.is_login_alive(st.session_state.einv_port)
                    if not st.session_state.einv_logged_in:
                        st.error("尚未偵測到登入成功，請確認已在跳出的視窗完成帳密與驗證碼")
            with _lc3:
                if st.session_state.einv_pid and st.button(
                    "結束連線", key="einv_close_btn", use_container_width=True
                ):
                    _einv.close_login_browser(st.session_state.einv_pid)
                    st.session_state.einv_port = None
                    st.session_state.einv_pid = None
                    st.session_state.einv_logged_in = False
                    st.rerun()

            if st.session_state.einv_port and not st.session_state.einv_logged_in:
                st.info("請在跳出的瀏覽器視窗輸入帳號密碼並完成圖形驗證碼登入，完成後按「我已完成登入，檢查連線」")

            if st.session_state.einv_logged_in and _general_orders:
                st.success(f"已連線，{len(_general_orders)} 張待處理")

                st.session_state.einv_dry_run = st.checkbox(
                    "測試模式（表單照樣填完，但自動按「放棄開立」，不會真的送出、也不會留下記錄）",
                    value=st.session_state.einv_dry_run, key="einv_dry_run_cb",
                )

                if st.session_state.einv_dry_run:
                    # 測試模式：一次全部跑完，每張都自動放棄開立，純粹驗證欄位填寫邏輯
                    with st.expander("📋 待處理清單", expanded=True):
                        for o in _general_orders:
                            _amt = sum((it.get("quantity", 0) or 0) * (it.get("unit_price", 0) or 0)
                                       for it in o.get("items", []))
                            st.caption(
                                f"**{o.get('order_no','')}** — {o.get('customer_name','')}"
                                f"（統編 {o.get('buyer_tax_id') or '⚠️缺'}）　未稅合計 {_amt:,.0f}"
                            )

                    if st.button("測試模式：全部跑一次", type="primary",
                                 use_container_width=True, key="einv_dryrun_btn"):
                        _progress_area = st.empty()
                        _log_lines = []

                        def _on_progress(i, total, order, result):
                            _mark = "🧪" if result.get("state") == "dry_run_cancelled" else "❌"
                            _err = f"（{result.get('error','')}）" if result.get("error") else ""
                            _log_lines.append(f"{_mark} {order.get('order_no','')}{_err}")
                            _progress_area.text("\n".join(_log_lines))

                        _einv.dry_run_batch(st.session_state.einv_port, _general_orders, on_progress=_on_progress)

                else:
                    # 正式送出：一次只處理一筆，填完停在瀏覽器畫面上，等人工親自按「開立發票」
                    st.warning("⚠️ 測試模式已關閉，接下來會是真實發票，號碼由網站配發、無法撤銷")

                    if not st.session_state.get("einv_pending_order"):
                        _next_order = _general_orders[0]
                        _amt = sum((it.get("quantity", 0) or 0) * (it.get("unit_price", 0) or 0)
                                   for it in _next_order.get("items", []))
                        st.markdown(
                            f"下一筆：**{_next_order.get('order_no','')}** — {_next_order.get('customer_name','')}"
                            f"（統編 {_next_order.get('buyer_tax_id') or '⚠️缺'}）　未稅合計 {_amt:,.0f}"
                        )
                        st.dataframe(
                            [{"品名": it.get("name", ""), "單位": it.get("unit", "PCS"),
                              "數量": it.get("quantity", 0), "單價": it.get("unit_price", 0),
                              "客戶料號": it.get("remark", "")}
                             for it in _next_order.get("items", [])],
                            use_container_width=True, hide_index=True,
                        )
                        if st.button("把這筆填進瀏覽器", type="primary",
                                     use_container_width=True, key="einv_fill_btn"):
                            if try_claim_order(_next_order.get("order_no", ""), _next_order.get("customer_name", ""),
                                                _next_order.get("buyer_tax_id", "")):
                                _fill_result = _einv.fill_one_order_via_port(
                                    st.session_state.einv_port, _next_order, dry_run=False)
                                if _fill_result.get("state") == "filled_awaiting_manual_submit":
                                    st.session_state.einv_pending_order = _next_order
                                else:
                                    save_einvoice_log(
                                        _next_order.get("order_no", ""), _next_order.get("customer_name", ""),
                                        _next_order.get("buyer_tax_id", ""), "", "失敗",
                                        _fill_result.get("error", "") or "填表失敗",
                                    )
                                    st.error(f"填表失敗：{_fill_result.get('error')}")
                            else:
                                st.error("認領失敗，可能有其他人正在處理這張銷貨單，稍後重新整理再試")
                            st.rerun()
                    else:
                        _pending = st.session_state.einv_pending_order
                        st.info(
                            f"已將「{_pending.get('order_no','')}」的資料填進瀏覽器視窗，"
                            "請切換過去確認內容後，**親自在瀏覽器裡按下「開立發票」**（或「放棄開立」）。"
                        )
                        _pc1, _pc2 = st.columns(2)
                        with _pc1:
                            if st.button("✅ 已在瀏覽器按下開立發票，查詢號碼並記錄",
                                         type="primary", use_container_width=True, key="einv_confirm_sent_btn"):
                                _no = _einv.lookup_invoice_no_via_port(
                                    st.session_state.einv_port, _pending.get("order_no", ""))
                                save_einvoice_log(
                                    _pending.get("order_no", ""), _pending.get("customer_name", ""),
                                    _pending.get("buyer_tax_id", ""), _no or "", "已開立",
                                    "" if _no else "已送出但查無發票號碼，需到查詢作業手動確認",
                                )
                                if not _no:
                                    st.warning("已記錄為已開立，但沒查到發票號碼，麻煩到「查詢作業」手動確認並回來補上")
                                st.session_state.einv_pending_order = None
                                st.rerun()
                        with _pc2:
                            if st.button("❌ 取消這筆（按了放棄開立／沒有送出）",
                                         use_container_width=True, key="einv_cancel_pending_btn"):
                                delete_einvoice_log(_pending.get("order_no", ""))
                                st.session_state.einv_pending_order = None
                                st.rerun()


# ════════════════════════════════════════════════════════════════
#  報表彙總
# ════════════════════════════════════════════════════════════════
_REPORT_FOLDER = Path(r"\\192.168.10.253\a10 210專區\生產日報表")


def _scan_report_files() -> list:
    """掃描月報表資料夾，排除 Excel 開啟時的鎖定暫存檔（~$開頭）跟範本檔（檔名含「範本」）。"""
    if not _REPORT_FOLDER.exists():
        return []
    files = [
        p for p in _REPORT_FOLDER.glob("*.xlsx")
        if not p.name.startswith("~$") and "範本" not in p.name
    ]
    return sorted(files, key=lambda p: p.name)


def _report_month_label(path) -> str:
    m = re.match(r"(\d{1,2})\s*月", path.name)
    return f"{m.group(1)}月" if m else path.stem


def _run_report_aggregation(name: str, file_bytes: bytes) -> dict:
    """跑完整套解析＋彙總，回傳結果 dict（含 error 代表失敗）給畫面顯示用。"""
    try:
        wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
        long_df = parse_daily_report_workbook(wb)
        if long_df.empty:
            return {"error": "找不到可解析的資料，請確認檔案格式（工作表需有「日期」「料號」標題列）"}
        wide_df = aggregate_daily_report(long_df)
        return {"wide": wide_df, "bytes": file_bytes, "name": name}
    except Exception as e:
        import traceback as _rtb
        return {"error": f"{e}\n\n{_rtb.format_exc()}"}


with tab_report:
    st.caption("依「日期＋料號＋站」彙總，同一天同一料號在同一站的數量會自動加總")

    _report_files = _scan_report_files()
    _month_opts = {_report_month_label(p): p for p in _report_files}

    if _report_files:
        st.success(f"從 `{_REPORT_FOLDER}` 自動讀取到 {len(_report_files)} 份月報表")
        _sel_months = st.multiselect(
            "選擇要匯出的月份（各月各自獨立產出一份彙總檔案）",
            options=list(_month_opts.keys()),
            default=list(_month_opts.keys()),
        )
    else:
        st.warning(f"讀不到資料夾 `{_REPORT_FOLDER}`（可能是網路磁碟沒連線或沒有權限），可改用下方手動上傳")
        _sel_months = []

    if "report_results" not in st.session_state:
        st.session_state["report_results"] = {}

    if _sel_months and st.button("彙總選中的月份", type="primary", key="report_gen_btn"):
        with st.spinner(f"彙總 {len(_sel_months)} 份月報表中..."):
            for _m in _sel_months:
                _p = _month_opts[_m]
                st.session_state["report_results"][_m] = _run_report_aggregation(_p.stem, _p.read_bytes())

    with st.expander("或手動上傳其他檔案（不在上面資料夾裡的）"):
        _report_file = st.file_uploader("選擇生產日報表 Excel", type=["xlsx", "xlsm"], key="report_uploader")
        if _report_file and st.button("彙總此檔案", key="report_manual_btn"):
            _name = re.sub(r"\.xlsx?$", "", _report_file.name, flags=re.IGNORECASE)
            with st.spinner("彙總中..."):
                st.session_state["report_results"][_name] = _run_report_aggregation(_name, _report_file.read())

    for _m, _result in st.session_state["report_results"].items():
        st.divider()
        st.markdown(f"### {_m}")
        if _result.get("error"):
            st.error(_result["error"])
            continue
        st.dataframe(_result["wide"], use_container_width=True, hide_index=True)
        try:
            _wb_out = openpyxl.load_workbook(BytesIO(_result["bytes"]))
            _buf = build_summary_workbook(_wb_out, _result["wide"])
            st.download_button(
                f"⬇️ 下載 {_result['name']}_已彙總.xlsx",
                data=_buf,
                file_name=f"{_result['name']}_已彙總.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"report_dl_{_m}",
            )
        except Exception as _rwe:
            import traceback as _rwtb
            st.error(f"產出總表失敗：{_rwe}")
            st.code(_rwtb.format_exc())
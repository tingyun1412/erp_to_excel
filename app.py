"""
倉庫出貨自動化工具
Streamlit 主介面

使用方式：
    streamlit run app.py
"""
import io
import json
import tempfile
from pathlib import Path

import streamlit as st

from rtf_parser import parse_sales_order_rtf
from module_a_calendar import generate_shipping_calendar
from module_b_invoice import generate_invoice_excel
from module_c_labels import generate_labels_excel, list_templates

# ── 頁面設定 ─────────────────────────────────────────────────────
st.set_page_config(
    page_title="出貨自動化工具",
    page_icon="📦",
    layout="wide",
)

st.title("📦 倉庫出貨自動化工具")
st.caption("上傳銷貨單 RTF → 一鍵產出出貨行事曆、電子發票、標籤")

# ── Session State 初始化 ─────────────────────────────────────────
if "parsed_orders" not in st.session_state:
    st.session_state.parsed_orders = []


# ════════════════════════════════════════════════════════════════
#  側邊欄：上傳銷貨單
# ════════════════════════════════════════════════════════════════
with st.sidebar:
    st.header("📂 上傳銷貨單")
    uploaded_files = st.file_uploader(
        "選擇 RTF 銷貨單（可多選）",
        type=["rtf"],
        accept_multiple_files=True,
        help="從 ERP 匯出的 RTF 格式銷貨單",
    )

    if uploaded_files:
        if st.button("🔍 解析銷貨單", use_container_width=True, type="primary"):
            orders = []
            errors = []
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
            
            st.session_state.parsed_orders = orders
            
            if errors:
                for err in errors:
                    st.error(err)
            else:
                st.success(f"✅ 成功解析 {len(orders)} 張銷貨單")

    st.divider()

    if st.session_state.parsed_orders:
        st.success(f"已載入 {len(st.session_state.parsed_orders)} 張銷貨單")
        if st.button("🗑 清除全部", use_container_width=True):
            st.session_state.parsed_orders = []
            st.rerun()


# ════════════════════════════════════════════════════════════════
#  主區域：分頁
# ════════════════════════════════════════════════════════════════
orders = st.session_state.parsed_orders

tab_preview, tab_a, tab_b, tab_c = st.tabs([
    "📋 銷貨單預覽",
    "📅 A｜出貨行事曆",
    "🧾 B｜電子發票",
    "🏷 C｜出貨標籤",
])


# ── Tab 0：預覽 ──────────────────────────────────────────────────
with tab_preview:
    if not orders:
        st.info("請先在左側上傳並解析銷貨單 RTF 檔案")
    else:
        st.subheader(f"共 {len(orders)} 張銷貨單")
        
        for i, order in enumerate(orders):
            with st.expander(
                f"📄 {order.get('filename','')}"
                f"  ｜  單號：{order.get('order_no','N/A')}"
                f"  ｜  品項數：{len(order.get('items',[]))}",
                expanded=(i == 0),
            ):
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("**基本資訊**")
                    st.write(f"銷貨單號：`{order.get('order_no','')}`")
                    st.write(f"銷貨日期：`{order.get('order_date','')}`")
                    st.write(f"客戶訂單號：`{order.get('customer_order_no','')}`")
                    st.write(f"賣方統編：`{order.get('seller_tax_id','')}`")
                    st.write(f"買方統編：`{order.get('buyer_tax_id','')}`")
                    st.write(f"聯絡人：`{order.get('contact','')}`")
                with col2:
                    st.markdown("**解析到的欄位值**")
                    raw = order.get("raw_values", [])
                    st.text("\n".join(raw[:20]) if raw else "（無）")
                
                items = order.get("items", [])
                if items:
                    st.markdown("**品項**")
                    st.dataframe(
                        [{"料號": it["item_no"],
                          "品名": it.get("description",""),
                          "數量": it.get("quantity",""),
                          "單位": it.get("unit","PC"),
                          "出貨日期": it.get("ship_date",""),
                          "客戶": it.get("customer",""),
                          "備註": it.get("remark","")}
                         for it in items],
                        use_container_width=True,
                        hide_index=True,
                    )
                else:
                    st.warning("⚠️ 未解析到品項資料，請確認 RTF 格式是否標準")
                
                # 手動補充欄位
                st.markdown("**補充資料（若自動解析不完整，可手動填入）**")
                c1, c2, c3 = st.columns(3)
                with c1:
                    new_buyer = st.text_input(
                        "買方統編", value=order.get("buyer_tax_id",""),
                        key=f"buyer_{i}"
                    )
                with c2:
                    new_seller = st.text_input(
                        "賣方統編", value=order.get("seller_tax_id",""),
                        key=f"seller_{i}"
                    )
                with c3:
                    new_date = st.text_input(
                        "出貨日期(YYYYMMDD)", value=order.get("order_date",""),
                        key=f"date_{i}"
                    )
                
                if st.button("💾 儲存補充", key=f"save_{i}"):
                    st.session_state.parsed_orders[i]["buyer_tax_id"]  = new_buyer
                    st.session_state.parsed_orders[i]["seller_tax_id"] = new_seller
                    st.session_state.parsed_orders[i]["order_date"]    = new_date
                    for item in st.session_state.parsed_orders[i]["items"]:
                        item["ship_date"] = new_date
                    st.success("已儲存")


# ── Tab A：出貨行事曆 ────────────────────────────────────────────
with tab_a:
    st.subheader("📅 產出出貨行事曆")
    st.markdown("""
    將所有銷貨單的出貨品項，依出貨日期填入月曆格式。  
    每個月份產生一個工作表，格式與原始倉庫出貨通知相同。
    """)
    
    if not orders:
        st.info("請先上傳並解析銷貨單")
    else:
        # 整理所有品項
        all_items = []
        for order in orders:
            for item in order.get("items", []):
                all_items.append({**item, "order_no": order.get("order_no","")})
        
        st.write(f"共 **{len(all_items)}** 筆出貨品項")
        
        if all_items:
            if st.button("產出出貨行事曆 Excel", type="primary", use_container_width=True):
                with st.spinner("產出中..."):
                    buf = generate_shipping_calendar(all_items)
                st.download_button(
                    label="⬇️ 下載出貨行事曆.xlsx",
                    data=buf,
                    file_name="出貨行事曆.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        else:
            st.warning("銷貨單中未解析到品項，無法產出行事曆")


# ── Tab B：電子發票 ──────────────────────────────────────────────
with tab_b:
    st.subheader("🧾 產出電子發票上傳 Excel")
    st.markdown("""
    依照 e-invoice.com.tw V1.6 格式產生可直接上傳的 Excel。  
    包含「發票主檔」和「發票明細」兩個工作表。
    """)
    
    if not orders:
        st.info("請先上傳並解析銷貨單")
    else:
        with st.expander("⚙️ 發票設定", expanded=True):
            col1, col2, col3 = st.columns(3)
            with col1:
                default_seller = next(
                    (o.get("seller_tax_id","") for o in orders if o.get("seller_tax_id")),
                    ""
                )
                seller_id = st.text_input("賣方統編（預設）", value=default_seller)
            with col2:
                inv_prefix = st.text_input("發票字軌（2碼）", value="AA", max_chars=2)
            with col3:
                start_num = st.number_input("起始號碼", min_value=1, value=1, step=1)
        
        orders_with_items = [o for o in orders if o.get("items")]
        st.write(f"共 **{len(orders_with_items)}** 張發票")
        
        if orders_with_items:
            if st.button("產出電子發票 Excel", type="primary", use_container_width=True):
                with st.spinner("產出中..."):
                    buf = generate_invoice_excel(
                        orders_with_items,
                        seller_tax_id=seller_id,
                        invoice_prefix=inv_prefix,
                        start_number=int(start_num),
                    )
                st.download_button(
                    label="⬇️ 下載電子發票上傳檔.xlsx",
                    data=buf,
                    file_name="電子發票上傳.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )


# ── Tab C：出貨標籤 ──────────────────────────────────────────────
with tab_c:
    st.subheader("🏷 產出出貨標籤 Excel")
    st.markdown("""
    每個品項產生一張標籤，可選擇不同格式。
    """)
    
    if not orders:
        st.info("請先上傳並解析銷貨單")
    else:
        templates = list_templates()
        
        with st.expander("⚙️ 標籤設定", expanded=True):
            selected_template = st.selectbox(
                "標籤格式",
                options=templates,
                index=0,
                help="選擇標籤樣式，開發者可在 module_c_labels.py 的 LABEL_TEMPLATES 新增格式",
            )
            st.caption(
                "💡 **新增標籤格式**：在 `module_c_labels.py` 的 `LABEL_TEMPLATES` 字典"
                "中加入新的 key-value（function），重啟 app 即可在上方下拉選單看到。"
            )
        
        total_items = sum(len(o.get("items",[])) for o in orders)
        st.write(f"共 **{total_items}** 張標籤")
        
        if total_items > 0:
            if st.button("產出標籤 Excel", type="primary", use_container_width=True):
                with st.spinner("產出中..."):
                    buf = generate_labels_excel(orders, template_name=selected_template)
                st.download_button(
                    label="⬇️ 下載出貨標籤.xlsx",
                    data=buf,
                    file_name="出貨標籤.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        else:
            st.warning("銷貨單中未解析到品項")

# ── Footer ───────────────────────────────────────────────────────
st.divider()
st.caption("v0.1 · 本地開發版 · 確認功能後可部署至 Streamlit Cloud 或打包成 .exe")

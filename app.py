import pandas as pd
import streamlit as st
import io

# 設定頁面資訊
st.set_page_config(page_title="會員抽獎券數據整理系統", layout="wide")

st.title("🎟️ 會員銷售業績與抽獎券資料整理系統")
st.markdown("請上傳 **會員銷售業績**、**會員管理** 以及 **抽獎券格式** 樣板檔案，並輸入指定產品編號後點擊確認，系統將自動比對並產出對應的抽獎券檔案。")

# 1. 檔案上傳區塊
col1, col2, col3 = st.columns(3)
with col1:
    uploaded_sales = st.file_uploader("1️⃣ 上傳【會員銷售業績】(.xls / .xlsx)", type=["xls", "xlsx"])
with col2:
    uploaded_members = st.file_uploader("2️⃣ 上傳【會員管理】(.xls / .xlsx)", type=["xls", "xlsx"])
with col3:
    uploaded_template = st.file_uploader("3️⃣ 上傳【抽獎券格式】樣板(.xlsx)", type=["xlsx"])

# 2. 條件設定：文字框無預設文字，並加入 placeholder 提示文字與確認按鈕
st.subheader("⚙️ 條件設定")

with st.form(key="condition_form"):
    specified_pids_input = st.text_area(
        "【條件一】請輸入指定的產品編號（多個請用逗號、半形分號或換行分隔）",
        value="",
        placeholder="請輸入產品編號",
        help="若訂單中包含此指定產品編號，將依據該筆對應的數量直接換算為抽獎券張數，且該筆金額會從訂單總金額中扣除後再計算條件二、三。"
    )
    
    # 確認按鈕
    submit_button = st.form_submit_button(label="🚀 確認並開始跑資料統整", use_container_width=True)

# 只有在三個檔案皆已上傳，且使用者點擊了確認按鈕後才執行統整
if uploaded_sales is not None and uploaded_members is not None and uploaded_template is not None:
    if submit_button:
        # 將輸入的產品編號解析成列表
        specified_pids = [
            pid.strip() for pid in specified_pids_input.replace('\n', ',').replace(';', ',').split(',') if pid.strip()
        ]

        try:
            # 通用檔案讀取函數
            def safe_read_file(uploaded_file):
                try:
                    uploaded_file.seek(0)
                    dfs = pd.read_html(uploaded_file)
                    if dfs and len(dfs) > 0:
                        return dfs[0]
                except Exception:
                    pass
                try:
                    uploaded_file.seek(0)
                    return pd.read_excel(uploaded_file, engine='openpyxl')
                except Exception:
                    pass
                try:
                    uploaded_file.seek(0)
                    return pd.read_excel(uploaded_file, engine='xlrd')
                except Exception:
                    pass
                uploaded_file.seek(0)
                content = uploaded_file.read()
                return pd.read_html(content.decode('utf-8', errors='ignore'))[0]

            # 讀取三個檔案
            df_sales = safe_read_file(uploaded_sales)
            df_members = safe_read_file(uploaded_members)
            df_template = pd.read_excel(uploaded_template, engine='openpyxl')
            
            template_columns = df_template.columns.tolist()

            # 清理銷售資料
            df_sales['會員編號'] = df_sales['會員編號'].astype(str).str.replace('="', '').str.replace('"', '').str.strip()
            df_sales['產品編號'] = df_sales['產品編號'].astype(str).str.replace('="', '').str.replace('"', '').str.strip()
            df_sales['會員名稱'] = df_sales['會員名稱'].astype(str).str.strip()
            df_sales['訂購單號'] = df_sales['訂購單號'].astype(str).str.replace('="', '').str.replace('"', '').str.strip()
            df_sales['數量'] = pd.to_numeric(df_sales['數量'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            df_sales['單價'] = pd.to_numeric(df_sales['單價'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            df_sales['金額小計'] = pd.to_numeric(df_sales['金額小計'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            df_sales['訂單總金額'] = pd.to_numeric(df_sales['訂單總金額'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

            # 清理會員管理資料
            df_members['會員編號'] = df_members['會員編號'].astype(str).str.replace('="', '').str.replace('"', '').str.strip()
            
            id_col = next((c for c in df_members.columns if '身份證' in c or '身分證' in c), None)
            phone_col = next((c for c in df_members.columns if '手機' in c or '電話' in c), None)

            member_info_map = {}
            for _, row in df_members.iterrows():
                m_id = row['會員編號']
                m_id_card = str(row[id_col]).replace('="', '').replace('"', '').strip() if id_col and pd.notna(row[id_col]) else ""
                m_phone = str(row[phone_col]).replace('="', '').replace('"', '').strip() if phone_col and pd.notna(row[phone_col]) else ""
                member_info_map[m_id] = {
                    '身份證字號': m_id_card if m_id_card != 'nan' else '',
                    '手機號碼': m_phone if m_phone != 'nan' else ''
                }

            # ---------------------------------------------------------
            # 核心邏輯計算
            # ---------------------------------------------------------
            
            # 1. 【條件一】：指定產品編號 -> 數量即為張數
            df_cond1 = df_sales[df_sales['產品編號'].isin(specified_pids)].copy()
            summary_cond1 = df_cond1.groupby(['會員編號', '會員名稱'])['數量'].sum().reset_index()
            summary_cond1.rename(columns={'數量': 'cond1_tickets'}, inplace=True)

            # 2. 計算每張訂單中屬於指定產品的金額小計總和，並自訂單總金額中扣除
            order_base = df_sales.groupby('訂購單號').agg({
                '會員編號': 'first',
                '會員名稱': 'first',
                '訂單總金額': 'first'
            }).reset_index()

            order_cond1_subtotal = df_cond1.groupby('訂購單號')['金額小計'].sum().reset_index()
            order_cond1_subtotal.rename(columns={'金額小計': 'excluded_subtotal'}, inplace=True)

            order_merged = pd.merge(order_base, order_cond1_subtotal, on='訂購單號', how='left')
            order_merged['excluded_subtotal'] = order_merged['excluded_subtotal'].fillna(0)
            order_merged['剩餘訂單金額'] = order_merged['訂單總金額'] - order_merged['excluded_subtotal']

            # 3. 依照剩餘金額跑條件二與條件三級距計算
            def calc_remaining_tickets(amt):
                if amt >= 12800:
                    return int(amt // 12800) * 2
                elif amt >= 6800:
                    return 1
                else:
                    return 0

            order_merged['other_tickets'] = order_merged['剩餘訂單金額'].apply(calc_remaining_tickets)
            summary_other = order_merged.groupby(['會員編號', '會員名稱'])['other_tickets'].sum().reset_index()

            # 4. 合併所有條件結果
            all_members = df_sales[['會員編號', '會員名稱']].drop_duplicates()
            merged = pd.merge(all_members, summary_cond1, on=['會員編號', '會員名稱'], how='left')
            merged = pd.merge(merged, summary_other, on=['會員編號', '會員名稱'], how='left')
            merged['cond1_tickets'] = merged['cond1_tickets'].fillna(0).astype(int)
            merged['other_tickets'] = merged['other_tickets'].fillna(0).astype(int)
            merged['抽獎券總共張數'] = merged['cond1_tickets'] + merged['other_tickets']

            df_summary_export = merged[['會員編號', '會員名稱', '抽獎券總共張數']].copy()
            df_summary_export = df_summary_export[df_summary_export['抽獎券總共張數'] > 0].reset_index(drop=True)

            # 5. 展開為【抽獎券格式】對應檔
            ticket_rows = []
            ticket_counter = 1

            for idx, row in df_summary_export.iterrows():
                total_cnt = row['抽獎券總共張數']
                m_id = row['會員編號']
                m_name = row['會員名稱']
                
                extra_info = member_info_map.get(m_id, {'身份證字號': '', '手機號碼': ''})

                for _ in range(total_cnt):
                    ticket_code = f"TH{ticket_counter:04d}"
                    row_dict = {col: None for col in template_columns}
                    row_dict['摸彩券編號'] = ticket_code
                    row_dict['會員編號'] = m_id
                    row_dict['會員姓名'] = m_name
                    if '身份證字號' in row_dict:
                        row_dict['身份證字號'] = extra_info['身份證字號']
                    if '手機號碼' in row_dict:
                        row_dict['手機號碼'] = extra_info['手機號碼']
                    
                    ticket_rows.append(row_dict)
                    ticket_counter += 1

            df_ticket_export = pd.DataFrame(ticket_rows)

            # ---------------------------------------------------------
            # 畫面呈現：下載區在上方，具備淺藍色放大按鈕
            # ---------------------------------------------------------
            st.success("✅ 資料比對與整理完成！")
            
            st.subheader("📥 檔案下載專區")
            
            st.markdown("""
            <style>
            div.stDownloadButton > button {
                background-color: #E0F2FE !important;
                color: #0369A1 !important;
                font-size: 18px !important;
                font-weight: bold !important;
                padding: 14px 28px !important;
                border-radius: 10px !important;
                border: 2px solid #38BDF8 !important;
                width: 100% !important;
                box-shadow: 0 4px 6px rgba(0,0,0,0.05);
                transition: 0.3s;
            }
            div.stDownloadButton > button:hover {
                background-color: #BAE6FD !important;
                border-color: #0284C7 !important;
                color: #0C4A6E !important;
            }
            </style>
            """, unsafe_allow_html=True)

            output_summary = io.BytesIO()
            with pd.ExcelWriter(output_summary, engine='openpyxl') as writer:
                df_summary_export.to_excel(writer, index=False, sheet_name='抽獎券彙總')
            output_summary.seek(0)

            output_tickets = io.BytesIO()
            with pd.ExcelWriter(output_tickets, engine='openpyxl') as writer:
                df_ticket_export.to_excel(writer, index=False, sheet_name='摸彩券明細')
            output_tickets.seek(0)

            d_col1, d_col2 = st.columns(2)
            with d_col1:
                st.download_button(
                    label="⬇️ 下載【會員抽獎券統計表.xlsx】",
                    data=output_summary,
                    file_name="會員抽獎券統計表.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with d_col2:
                st.download_button(
                    label="⬇️ 下載【抽獎券明細表 (含身份證與手機).xlsx】",
                    data=output_tickets,
                    file_name="抽獎券明細表.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            st.markdown("---")

            st.subheader("📊 會員抽獎券統計預覽")
            st.dataframe(df_summary_export.head(10), use_container_width=True)

            st.subheader("🎟️ 抽獎券展開明細預覽")
            st.dataframe(df_ticket_export.head(10), use_container_width=True)

        except Exception as e:
            st.error(f"❌ 處理過程發生錯誤：{str(e)}")
    else:
        st.info("💡 請在上方輸入指定的產品編號，並點擊下方的 **「確認並開始跑資料統整」** 按鈕開始計算。")
else:
    st.info("💡 請先上傳全部三個檔案（會員銷售業績、會員管理、抽獎券格式樣板）。")

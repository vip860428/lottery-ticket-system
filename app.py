import pandas as pd
import streamlit as st
import io

# 設定頁面資訊
st.set_page_config(page_title="會員抽獎券數據整理系統", layout="wide")

st.title("🎟️ 會員銷售業績與抽獎券資料整理系統")
st.markdown("請上傳 **會員銷售業績** Excel/HTML 檔案與 **抽獎券格式** 樣板檔案，系統將自動依據條件計算並匯出符合格式的結果檔案。")

# 1. 檔案上傳區塊
col1, col2 = st.columns(2)
with col1:
    uploaded_sales = st.file_uploader("上傳【會員銷售業績】(.xls / .xlsx)", type=["xls", "xlsx"])
with col2:
    uploaded_template = st.file_uploader("上傳【抽獎券格式】(.xlsx)", type=["xlsx"])

# 2. 手動輸入指定產品編號（條件一）
st.subheader("⚙️ 條件設定")
specified_pids_input = st.text_area(
    "【條件一】請輸入指定的產品編號（多個請用逗號、半形分號或換行分隔）",
    value="GTH417826, GTH417854",
    help="若訂單中包含此指定產品編號，將依據該筆對應的數量直接換算為抽獎券張數（數量1=1張，數量2=2張...）"
)

if uploaded_sales is not None and uploaded_template is not None:
    # 解析指定產品編號列表
    specified_pids = [
        pid.strip() for pid in specified_pids_input.replace('\n', ',').replace(';', ',').split(',') if pid.strip()
    ]

    try:
        # 讀取銷售檔案 (支援 HTML 格式的 .xls 檔與一般 .xlsx 檔)
        try:
            dfs = pd.read_html(uploaded_sales)
            df_sales = dfs[0]
        except Exception:
            uploaded_sales.seek(0)
            df_sales = pd.read_excel(uploaded_sales)

        # 讀取摸彩券格式樣板
        df_template = pd.read_excel(uploaded_template)
        template_columns = df_template.columns.tolist() # 保留欄位結構: ['摸彩券編號', '會員編號', '會員姓名', '身份證字號', '手機號碼']

        # 清理銷售資料
        df_sales['會員編號'] = df_sales['會員編號'].astype(str).str.replace('="', '').str.replace('"', '').str.strip()
        df_sales['產品編號'] = df_sales['產品編號'].astype(str).str.replace('="', '').str.replace('"', '').str.strip()
        df_sales['會員名稱'] = df_sales['會員名稱'].astype(str).str.strip()
        df_sales['數量'] = pd.to_numeric(df_sales['數量'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        df_sales['訂單總金額'] = pd.to_numeric(df_sales['訂單總金額'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

        # ---------------------------------------------------------
        # 邏輯計算
        # ---------------------------------------------------------
        
        # 【條件一】：指定產品編號 -> 數量即為張數
        df_cond1 = df_sales[df_sales['產品編號'].isin(specified_pids)].copy()
        summary_cond1 = df_cond1.groupby(['會員編號', '會員名稱'])['數量'].sum().reset_index()
        summary_cond1.rename(columns={'數量': 'cond1_tickets'}, inplace=True)

        # 【條件二與三】：排除指定產品編號後的訂單金額計算
        df_other = df_sales[~df_sales['產品編號'].isin(specified_pids)].copy()

        # 以「訂購單號」為單位彙總該筆訂單金額
        order_summary = df_other.groupby(['訂購單號', '會員編號', '會員名稱'])['訂單總金額'].first().reset_index()

        def calc_order_tickets(amt):
            if amt >= 12800:
                # 條件三: 每滿 12,800 元給 2 張 (級距遞增: 12,800->2張, 25,600->4張, 38,400->6張...)
                return int(amt // 12800) * 2
            elif amt >= 6800:
                # 條件二: 6,800 ~ 12,799 元給 1 張
                return 1
            else:
                return 0

        order_summary['tickets'] = order_summary['訂單總金額'].apply(calc_order_tickets)
        summary_other = order_summary.groupby(['會員編號', '會員名稱'])['tickets'].sum().reset_index()
        summary_other.rename(columns={'tickets': 'other_tickets'}, inplace=True)

        # 合併所有條件結果
        all_members = df_sales[['會員編號', '會員名稱']].drop_duplicates()
        merged = pd.merge(all_members, summary_cond1, on=['會員編號', '會員名稱'], how='left')
        merged = pd.merge(merged, summary_other, on=['會員編號', '會員名稱'], how='left')
        merged['cond1_tickets'] = merged['cond1_tickets'].fillna(0).astype(int)
        merged['other_tickets'] = merged['other_tickets'].fillna(0).astype(int)
        merged['抽獎券總共張數'] = merged['cond1_tickets'] + merged['other_tickets']

        # 條件四檔案 A：可供下載之彙總統計檔
        df_summary_export = merged[['會員編號', '會員名稱', '抽獎券總共張數']].copy()
        df_summary_export = df_summary_export[df_summary_export['抽獎券總共張數'] > 0].reset_index(drop=True)

        # 條件四檔案 B：展開為【抽獎券格式】對應檔
        ticket_rows = []
        ticket_counter = 1

        for idx, row in df_summary_export.iterrows():
            total_cnt = row['抽獎券總共張數']
            for _ in range(total_cnt):
                ticket_code = f"TH{ticket_counter:04d}" # 生成 TH0001, TH0002...
                # 依照樣板欄位結構填入資料（若樣板有額外欄位如身份證字號、手機號碼，預設為空）
                row_dict = {col: None for col in template_columns}
                row_dict['摸彩券編號'] = ticket_code
                row_dict['會員編號'] = row['會員編號']
                row_dict['會員姓名'] = row['會員名稱']
                ticket_rows.append(row_dict)
                ticket_counter += 1

        df_ticket_export = pd.DataFrame(ticket_rows)

        # 顯示處理結果預覽
        st.success("✅ 資料整理 completed！")
        
        st.subheader("📊 會員抽獎券統計預覽")
        st.dataframe(df_summary_export.head(10), use_container_width=True)

        st.subheader("🎟️ 抽獎券展開明細預覽")
        st.dataframe(df_ticket_export.head(10), use_container_width=True)

        # ---------------------------------------------------------
        # 提供檔案下載
        # ---------------------------------------------------------
        # 生成檔案 1: 會員抽獎券總張數彙總表.xlsx
        output_summary = io.BytesIO()
        with pd.ExcelWriter(output_summary, engine='openpyxl') as writer:
            df_summary_export.to_excel(writer, index=False, sheet_name='抽獎券彙總')
        output_summary.seek(0)

        # 生成檔案 2: 抽獎券明細表.xlsx (100% 符合樣板格式與欄位名稱)
        output_tickets = io.BytesIO()
        with pd.ExcelWriter(output_tickets, engine='openpyxl') as writer:
            df_ticket_export.to_excel(writer, index=False, sheet_name='摸彩券明細')
        output_tickets.seek(0)

        st.subheader("📥 檔案下載")
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
                label="⬇️ 下載【抽獎券明細表 (符合摸彩券格式).xlsx】",
                data=output_tickets,
                file_name="抽獎券明細表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"❌ 資料處理時發生錯誤：{str(e)}")
else:
    st.info("💡 請上傳兩個必備檔案以開始進行計算。")
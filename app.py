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

# 2. 條件設定：文字框與動態防呆按鈕
st.subheader("⚙️ 條件設定")

# 文字框輸入
specified_pids_input = st.text_area(
    "【條件一】請輸入指定的產品編號（多個請用逗號、半形分號或換行分隔）",
    value="",
    placeholder="請輸入產品編號",
    help="若訂單中包含此指定產品編號，將依據該筆對應的數量直接換算為抽獎券張數，且該筆金額會從訂單總金額中扣除後再計算條件二、三。"
)

# 檢查是否有輸入內容
has_input = bool(specified_pids_input and specified_pids_input.strip())

# 動態 CSS：依據是否有輸入內容來改變按鈕顏色與點選狀態
if not has_input:
    # 未輸入：紅色、禁用
    btn_bg = "#EF4444"
    btn_hover = "#DC2626"
    btn_cursor = "not-allowed"
    btn_opacity = "0.7"
else:
    # 已輸入：綠色、可點選
    btn_bg = "#10B981"
    btn_hover = "#059669"
    btn_cursor = "pointer"
    btn_opacity = "1.0"

st.markdown(f"""
<style>
/* 縮小按鈕並靠左，呈現動態防呆色彩 */
div.stButton > button {{
    background-color: {btn_bg} !important;
    color: white !important;
    font-size: 15px !important;
    font-weight: bold !important;
    padding: 8px 20px !important;
    border-radius: 8px !important;
    border: none !important;
    cursor: {btn_cursor} !important;
    opacity: {btn_opacity} !important;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    transition: 0.3s;
    width: auto !important;
}}
div.stButton > button:hover {{
    background-color: {btn_hover} !important;
}}
</style>
""", unsafe_allow_html=True)

# 建立精巧的確認按鈕（放在左下方）
run_clicked = st.button("🚀 確認並開始彙整", disabled=not has_input)

# 只有在三個檔案皆已上傳，且使用者點擊了已解鎖的確認按鈕後才執行統整
if uploaded_sales is not None and uploaded_members is not None and uploaded_template is not None:
    if run_clicked:
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
            df_sales['訂購單號'] = df_sales['訂購單號'].astype(str).str.replace('="',

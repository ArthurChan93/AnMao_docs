import streamlit as st
import pandas as pd
from io import BytesIO
import re

# 页面基本设置
st.set_page_config(page_title="安贸数据整合系统", layout="wide")
st.title("📁 安贸审核资料自动整合系统")
st.subheader("YAMAHA 供应商数据自动化处理平台", divider="rainbow")

# 初始化session state
if 'mc_data' not in st.session_state:
    st.session_state.mc_data = None
if 'rel_data' not in st.session_state:
    st.session_state.rel_data = None
if 'stock_data' not in st.session_state:
    st.session_state.stock_data = None

# ================== 左侧MC Info处理区 ==================
with st.container(border=True):
    left_col, _ = st.columns([3, 1])
    with left_col:
        st.subheader("⚙️ MC Info Sheet信息处理区", divider="blue")

        mc_files = st.file_uploader(
            "请上传MC Info文件（可多选）",
            type=['xls', 'xlsx', 'xlsm'],
            accept_multiple_files=True,
            key="mc_uploader"
        )

        if mc_files:
            valid_ext_mc_files = [f for f in mc_files if f.name.lower().endswith(('.xls', '.xlsx', '.xlsm'))]
            non_excel_files = [f.name for f in mc_files if f not in valid_ext_mc_files]

            if non_excel_files:
                st.warning(f"已忽略非Excel文件：{', '.join(non_excel_files)}")

            invalid_files = [f.name for f in valid_ext_mc_files if 'MC Info' not in f.name]

            if invalid_files:
                st.error(f"❌ 不合格文件名：{', '.join(invalid_files)}")
                st.stop()
            elif valid_ext_mc_files:
                mc_data = []
                for file in valid_ext_mc_files:
                    try:
                        if file.name.endswith('.xls'):
                            df = pd.read_excel(file, header=None, engine='xlrd')
                        else:
                            df = pd.read_excel(file, header=None, engine='openpyxl')

                        cd_code = df.iloc[10, 3]
                        row = 20
                        while row < len(df):
                            if pd.isna(df.iloc[row, 0]) and pd.isna(df.iloc[row, 1]) and pd.isna(df.iloc[row, 2]) and pd.isna(df.iloc[row, 3]):
                                row += 1
                                continue

                            mc_data.append({
                                "CD Code": cd_code,
                                "Machine Type": df.iloc[row, 2],
                                "S/N#": df.iloc[row, 3],
                                "File_name": file.name
                            })
                            row += 1
                    except Exception as e:
                        st.error(f"文件 {file.name} 读取失败：{str(e)}")
                        st.stop()

                st.session_state.mc_data = pd.DataFrame(mc_data)
                st.dataframe(st.session_state.mc_data, use_container_width=True)

                if not st.session_state.mc_data.empty:
                    buffer = BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        st.session_state.mc_data.to_excel(writer, sheet_name="MC Info", index=False)
                        worksheet = writer.sheets["MC Info"]
                        worksheet.freeze_panes(1, 0)
                        for col_num, col_name in enumerate(st.session_state.mc_data.columns):
                            max_len = max(st.session_state.mc_data[col_name].astype(str).str.len().max(), len(col_name)) + 2
                            worksheet.set_column(col_num, col_num, max_len)
                    
                    st.download_button(
                        "💾 下载机台信息",
                        data=buffer.getvalue(),
                        file_name="MC_Info_Data.xlsx",
                        mime="application/vnd.ms-excel"
                    )

# ================== 右侧Relocation处理区 ==================
with st.container(border=True):
    right_col, _ = st.columns([3, 1])
    with right_col:
        st.subheader("🚚 Relocation_sheet信息处理区", divider="orange")

        rel_files = st.file_uploader(
            "请上传Relocation文件（可多选）",
            type=['xls', 'xlsx', 'xlsm'],
            accept_multiple_files=True,
            key="rel_uploader"
        )

        if rel_files:
            valid_ext_rel_files = [f for f in rel_files if f.name.lower().endswith(('.xls', '.xlsx', '.xlsm'))]
            non_excel_files = [f.name for f in rel_files if f not in valid_ext_rel_files]

            if non_excel_files:
                st.warning(f"已忽略非Excel文件：{', '.join(non_excel_files)}")

            pattern = re.compile(r'relocation', re.IGNORECASE)
            invalid_files = [f.name for f in valid_ext_rel_files if not pattern.search(f.name)]

            if invalid_files:
                st.error(f"❌ 不合格文件名：{', '.join(invalid_files)}")
                st.stop()
            elif valid_ext_rel_files:
                rel_data = []
                for file in valid_ext_rel_files:
                    try:
                        if file.name.endswith('.xls'):
                            df = pd.read_excel(file, header=None, engine='xlrd')
                        else:
                            df = pd.read_excel(file, header=None, engine='openpyxl')

                        from_cd = df.iloc[24, 3]
                        to_cd = df.iloc[26, 3]
                        row = 32
                        while row < len(df):
                            if row < len(df):
                                if pd.isna(df.iloc[row, 7]) and pd.isna(df.iloc[row, 8]):
                                    break
                            else:
                                break

                            rel_data.append({
                                "From_CD Code": from_cd,
                                "To_CD Code": to_cd,
                                "Machine Type": df.iloc[row, 1],
                                "S/N#": df.iloc[row, 4],
                                "File_name": file.name
                            })
                            row += 1
                    except Exception as e:
                        st.error(f"文件 {file.name} 读取失败：{str(e)}")
                        st.stop()

                st.session_state.rel_data = pd.DataFrame(rel_data)
                st.dataframe(st.session_state.rel_data, use_container_width=True)

                if not st.session_state.rel_data.empty:
                    buffer = BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        st.session_state.rel_data.to_excel(writer, sheet_name="Relocation", index=False)
                        worksheet = writer.sheets["Relocation"]
                        worksheet.freeze_panes(1, 0)
                        for col_num, col_name in enumerate(st.session_state.rel_data.columns):
                            max_len = max(st.session_state.rel_data[col_name].astype(str).str.len().max(), len(col_name)) + 2
                            worksheet.set_column(col_num, col_num, max_len)
                    
                    st.download_button(
                        "💾 下载移机信息",
                        data=buffer.getvalue(),
                        file_name="Relocation_Data.xlsx",
                        mime="application/vnd.ms-excel"
                    )

# ================== Stock Machine处理区 ==================
with st.container(border=True):
    stock_col, _ = st.columns([3, 1])
    with stock_col:
        st.subheader("📦 Stock Machine处理区", divider="violet")
        
        stock_files = st.file_uploader(
            "请上传Stock Machine文件（可多选）",
            type=['xls', 'xlsx', 'xlsm'],
            accept_multiple_files=True,
            key="stock_uploader"
        )

        if stock_files:
            valid_stock_files = [f for f in stock_files if f.name.lower().endswith(('.xls', '.xlsx', '.xlsm'))]
            non_excel_files = [f.name for f in stock_files if f not in valid_stock_files]
            
            if non_excel_files:
                st.warning(f"已忽略非Excel文件：{', '.join(non_excel_files)}")
            
            pattern = re.compile(r'(Stock Machine|二合一)', re.IGNORECASE)
            invalid_files = [f.name for f in valid_stock_files if not pattern.search(f.name)]
            
            if invalid_files:
                st.error(f"❌ 不合格文件名：{', '.join(invalid_files)}")
                st.stop()
            else:
                normal_files = [f for f in valid_stock_files if '二合一' not in f.name]
                combined_files = [f for f in valid_stock_files if '二合一' in f.name]

                normal_data = []
                for file in normal_files:
                    try:
                        if file.name.endswith('.xls'):
                            df = pd.read_excel(file, header=None, engine='xlrd')
                        else:
                            df = pd.read_excel(file, header=None, engine='openpyxl')
                        
                        # 增强括号匹配逻辑
                        c15 = str(df.iloc[14, 2])
                        matches = re.findall(r'[（(]([^）)]+)[）)]', c15)
                        if matches:
                            cd_code = matches[-1].strip()  # 取最后一个括号内容
                        else:
                            cd_code = ''
                            st.warning(f"文件 {file.name} 的C15单元格未找到有效括号内容：{c15}")
                        
                        row = 20
                        while row < len(df):
                            if pd.isna(df.iloc[row, 8]) and pd.isna(df.iloc[row, 9]):
                                break
                            
                            b_col = df.iloc[row, 1]
                            e_col = df.iloc[row, 4]
                            
                            if pd.notna(b_col) or pd.notna(e_col):
                                normal_data.append({
                                    "CD Code": cd_code,
                                    "Machine Type": b_col,
                                    "S/N#": e_col,
                                    "File_name": file.name
                                })
                            row += 1
                    except Exception as e:
                        st.error(f"文件 {file.name} 处理失败：{str(e)}")
                        st.stop()

                combined_data = []
                for file in combined_files:
                    try:
                        if file.name.endswith('.xls'):
                            df = pd.read_excel(file, header=None, engine='xlrd')
                        else:
                            df = pd.read_excel(file, header=None, engine='openpyxl')
                        
                        cd_end_user = df.iloc[14, 3]
                        cd_distributor = df.iloc[15, 3]
                        
                        row = 21
                        while row < len(df):
                            if pd.isna(df.iloc[row, 9]):
                                break
                            
                            c_col = df.iloc[row, 2]
                            f_col = df.iloc[row, 5]
                            
                            if pd.notna(c_col) or pd.notna(f_col):
                                combined_data.append({
                                    "CD Code_End User": cd_end_user,
                                    "CD Code_Distributor": cd_distributor,
                                    "Machine Type": c_col,
                                    "S/N#": f_col,
                                    "File_name": file.name
                                })
                            row += 1
                    except Exception as e:
                        st.error(f"文件 {file.name} 处理失败：{str(e)}")
                        st.stop()

                stock_df = pd.DataFrame()
                if normal_data:
                    normal_df = pd.DataFrame(normal_data)
                    stock_df = pd.concat([stock_df, normal_df], ignore_index=True)
                if combined_data:
                    combined_df = pd.DataFrame(combined_data)
                    stock_df = pd.concat([stock_df, combined_df], ignore_index=True)
                
                st.session_state.stock_data = stock_df

                if not st.session_state.stock_data.empty:
                    st.dataframe(st.session_state.stock_data, use_container_width=True)
                    
                    buffer = BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        if normal_data:
                            normal_df.to_excel(writer, sheet_name="STOCK MACHINE SHIPPING INFO", index=False)
                            worksheet = writer.sheets["STOCK MACHINE SHIPPING INFO"]
                            worksheet.freeze_panes(1, 0)
                            for col_num, col_name in enumerate(normal_df.columns):
                                max_len = max(normal_df[col_name].astype(str).str.len().max(), len(col_name)) + 2
                                worksheet.set_column(col_num, col_num, max_len)
                        
                        if combined_data:
                            combined_df.to_excel(writer, sheet_name="二合一STOCK MACHINE SHIPPING INFO", index=False)
                            worksheet = writer.sheets["二合一STOCK MACHINE SHIPPING INFO"]
                            worksheet.freeze_panes(1, 0)
                            for col_num, col_name in enumerate(combined_df.columns):
                                max_len = max(combined_df[col_name].astype(str).str.len().max(), len(col_name)) + 2
                                worksheet.set_column(col_num, col_num, max_len)
                    
                    st.download_button(
                        "💾 下载Stock machine数据",
                        data=buffer.getvalue(),
                        file_name="Stock_Data.xlsx",
                        mime="application/vnd.ms-excel"
                    )

# ================== 整合下载区 ==================
if st.session_state.mc_data is not None or st.session_state.rel_data is not None or st.session_state.stock_data is not None:
    st.divider()
    with st.container(border=True):
        center_col, _ = st.columns([1, 3])
        with center_col:
            st.subheader("🚀 数据整合下载区", divider="green")

            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                sheets = {
                    "MC Info": st.session_state.mc_data,
                    "Relocation": st.session_state.rel_data,
                    "STOCK MACHINE SHIPPING INFO": st.session_state.stock_data[st.session_state.stock_data['File_name'].str.contains('二合一') == False] if st.session_state.stock_data is not None else pd.DataFrame(),
                    "二合一STOCK MACHINE SHIPPING INFO": st.session_state.stock_data[st.session_state.stock_data['File_name'].str.contains('二合一')] if st.session_state.stock_data is not None else pd.DataFrame()
                }

                for sheet_name, df in sheets.items():
                    if df is not None and not df.empty:
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        worksheet = writer.sheets[sheet_name]
                        worksheet.freeze_panes(1, 0)
                        for col_num, col_name in enumerate(df.columns):
                            max_len = max(df[col_name].astype(str).str.len().max(), len(col_name)) + 2
                            worksheet.set_column(col_num, col_num, max_len)

            st.download_button(
                "🌟 下载完整整合报告",
                data=buffer.getvalue(),
                file_name="Full_Consolidated_Report.xlsx",
                mime="application/vnd.ms-excel",
                use_container_width=True,
                key="full_download"
            )

# ================== 页面样式优化 ==================
st.markdown("""
<style>
[data-testid="stFileUploader"] {
    background-color: #f0f2f6;
    border-radius: 10px;
    padding: 20px;
}
.stDownloadButton button {
    border-radius: 8px !important;
    padding: 10px 24px !important;
}
.stDownloadButton [data-testid="baseButton-secondary"] {
    background: linear-gradient(45deg, #32CD32, #228B22) !important;
    color: white !important;
}
</style>
""", unsafe_allow_html=True)

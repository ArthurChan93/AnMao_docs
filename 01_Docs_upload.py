import streamlit as st
import pandas as pd
from io import BytesIO
import re
 
# 页面基本设置
st.set_page_config(page_title="安贸数据整合系统", layout="wide")
st.title("📁 YAMAHA 供应商安贸审核资料")
st.subheader("数据自动化处理系统", divider="rainbow")
 
# 初始化session state
if 'mc_data' not in st.session_state:
    st.session_state.mc_data = None
if 'rel_data' not in st.session_state:
    st.session_state.rel_data = None
 
# ================== 左侧MC Info处理区 ==================
with st.container(border=True):
    left_col, _ = st.columns([3, 1])
    with left_col:
        st.subheader(":notebook_with_decorative_cover: MC Info Sheet 信息处理区", divider="blue")
 
        # Excel上传器（允许xls/xlsx/xlsm）
        mc_files = st.file_uploader(
            "请上传MC Info excel文件（可多选）",
            type=['xls', 'xlsx', 'xlsm'],
            accept_multiple_files=True,
            key="mc_uploader"
        )
 
        if mc_files:
            # 过滤非Excel文件（二次验证）
            valid_ext_mc_files = [f for f in mc_files if f.name.lower().endswith(('.xls', '.xlsx', '.xlsm'))]
            non_excel_files = [f.name for f in mc_files if f not in valid_ext_mc_files]
 
            # 若有非Excel文件则提示（但不中断）
            if non_excel_files:
                st.warning(f"已忽略非Excel文件：{', '.join(non_excel_files)}")
 
            # 验证文件名必须含"MC Info"
            invalid_files = [f.name for f in valid_ext_mc_files if 'MC Info' not in f.name]
 
            if invalid_files:
                st.error(f"❌ 不合格文件名：{', '.join(invalid_files)}")
                st.stop()
            elif valid_ext_mc_files:  # 仅当有有效文件时处理
                mc_data = []
                for file in valid_ext_mc_files:
                    try:
                        df = pd.read_excel(file, header=None, engine='xlrd' if file.name.endswith('.xls') else 'openpyxl')
 
                        # 提取数据（修改点：D21单元格）
                        cd_code = df.iloc[10, 3]  # D11 (0-based索引)
                        row = 20  # C21起始行
 
                        while row < len(df):
                            # 检查A列至D列是否都为空
                            if pd.isna(df.iloc[row, 0]) and pd.isna(df.iloc[row, 1]) and pd.isna(df.iloc[row, 2]) and pd.isna(df.iloc[row, 3]):
                                row += 1
                                continue
 
                            mc_data.append({
                                "CD Code": cd_code,
                                "Machine Type": df.iloc[row, 2],  # C列
                                "S/N#": df.iloc[row, 3],  # D列
                                "File_name": file.name  # 新增File_name列
                            })
                            row += 1
                    except Exception as e:
                        st.error(f"文件 {file.name} 读取失败：{str(e)}")
                        st.stop()
 
                st.session_state.mc_data = pd.DataFrame(mc_data)
 
                # 显示表格
                st.dataframe(st.session_state.mc_data, use_container_width=True)
 
                # 下载按钮（带格式设置）
                if not st.session_state.mc_data.empty:
                    buffer = BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        st.session_state.mc_data.to_excel(writer,
                                                         sheet_name="MC Info",
                                                         index=False)
                        workbook = writer.book
                        worksheet = writer.sheets["MC Info"]
 
                        # 冻结首行并设置列宽
                        worksheet.freeze_panes(1, 0)  # 冻结第一行
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
        st.subheader(":green_book: Relocation_sheet 信息处理区", divider="orange")
 
        # Excel上传器（允许xls/xlsx/xlsm）
        rel_files = st.file_uploader(
            "请上传Relocation excel文件（可多选）",
            type=['xls', 'xlsx', 'xlsm'],
            accept_multiple_files=True,
            key="rel_uploader"
        )
 
        if rel_files:
            # 过滤非Excel文件（二次验证）
            valid_ext_rel_files = [f for f in rel_files if f.name.lower().endswith(('.xls', '.xlsx', '.xlsm'))]
            non_excel_files = [f.name for f in rel_files if f not in valid_ext_rel_files]
 
            if non_excel_files:
                st.warning(f"已忽略非Excel文件：{', '.join(non_excel_files)}")
 
            # 验证文件名必须含"relocation"（支持正则）
            pattern = re.compile(r'relocation', re.IGNORECASE)
            invalid_files = [f.name for f in valid_ext_rel_files if not pattern.search(f.name)]
 
            if invalid_files:
                st.error(f"❌ 不合格文件名：{', '.join(invalid_files)}")
                st.stop()
            elif valid_ext_rel_files:  # 仅当有有效文件时处理
                rel_data = []
                for file in valid_ext_rel_files:
                    try:
                        df = pd.read_excel(file, header=None, engine='xlrd' if file.name.endswith('.xls') else 'openpyxl')
 
                        # 提取固定值
                        from_cd = df.iloc[24, 3]  # D25 (0-based)
                        to_cd = df.iloc[26, 3]  # D27
 
                        row = 32  # B33起始行
                        while row < len(df):
                            # 确保当前行索引在有效范围内再检查I列和H列
                            if row < len(df):
                                if pd.isna(df.iloc[row, 7]) and pd.isna(df.iloc[row, 8]):
                                    break
                            else:
                                break
 
                            rel_data.append({
                                "From_CD Code": from_cd,
                                "To_CD Code": to_cd,
                                "Machine Type": df.iloc[row, 1],  # B列
                                "S/N#": df.iloc[row, 4],  # E列
                                "File_name": file.name  # 新增File_name列
                            })
                            row += 1
                    except Exception as e:
                        st.warning(f"文件 {file.name} 读取时出现问题，将采用备用抓取逻辑：{str(e)}")
                        try:
                            df = pd.read_excel(file, header=None, engine='xlrd' if file.name.endswith('.xls') else 'openpyxl')
                            from_cd = df.iloc[24, 3]  # D25 (0-based)
                            to_cd = df.iloc[26, 3]  # D27
                            row = 32  # B33起始行
                            while row < len(df):
                                if pd.isna(df.iloc[row, 0]):
                                    break
                                rel_data.append({
                                    "From_CD Code": from_cd,
                                    "To_CD Code": to_cd,
                                    "Machine Type": df.iloc[row, 1],  # B列
                                    "S/N#": df.iloc[row, 4],  # E列
                                    "File_name": file.name  # 新增File_name列
                                })
                                row += 1
                        except Exception as inner_e:
                            st.error(f"文件 {file.name} 采用备用抓取逻辑仍失败：{str(inner_e)}")
 
                st.session_state.rel_data = pd.DataFrame(rel_data)
 
                # 显示表格
                st.dataframe(st.session_state.rel_data, use_container_width=True)
 
                # 下载按钮（带格式设置）
                if not st.session_state.rel_data.empty:
                    buffer = BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        st.session_state.rel_data.to_excel(writer,
                                                          sheet_name="Relocation",
                                                          index=False)
                        workbook = writer.book
                        worksheet = writer.sheets["Relocation"]
 
                        # 冻结首行并设置列宽
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
 
# ================== 组合下载区 ==================
if st.session_state.mc_data is not None and st.session_state.rel_data is not None:
    st.divider()
    with st.container(border=True):
        center_col, _ = st.columns([1, 3])
        with center_col:
            st.subheader("🚀 数据整合下载区", divider="green")
 
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                st.session_state.mc_data.to_excel(writer,
                                                 sheet_name="MC Info",
                                                 index=False)
                st.session_state.rel_data.to_excel(writer,
                                                  sheet_name="Relocation",
                                                  index=False)
 
                # 设置两个sheet的格式
                for sheet_name in ["MC Info", "Relocation"]:
                    worksheet = writer.sheets[sheet_name]
                    worksheet.freeze_panes(1, 0)
                    df = st.session_state.mc_data if sheet_name == "MC Info" else st.session_state.rel_data
                    for col_num, col_name in enumerate(df.columns):
                        max_len = max(df[col_name].astype(str).str.len().max(), len(col_name)) + 2
                        worksheet.set_column(col_num, col_num, max_len)
 
            st.download_button(
                "🌟 下载完整报告",
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
/* 组合下载按钮橙色样式 */
.stDownloadButton [data-testid="baseButton-secondary"] {
    background: linear-gradient(45deg, #32CD32, #228B22) !important; /* 漸變綠色 */
    color: white !important;
}
</style>
""", unsafe_allow_html=True)

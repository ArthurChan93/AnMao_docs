import streamlit as st
import pandas as pd
from io import BytesIO
import re

# 页面基本设置
st.set_page_config(page_title="安贸数据整合系统", layout="wide")
st.title("📁 安贸审核资料自动整合系统")
st.subheader("YAMAHA 供应商数据自动化处理平台(MC sheet/ Relocation sheet/ STK MACHINE SHIPPPING info sheet)", divider="rainbow")

# 初始化session state
session_defaults = {
    'mc_data': None,
    'rel_data': None,
    'stock_data': None,
    'processed_files': set(),
    'mc_success_count': 0,
    'rel_success_count': 0,
    'stock_success_count': 0
}
for key, value in session_defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value

# ================== 统一文件处理函数 ==================
def process_mc_file(file):
    """处理MC Info文件"""
    try:
        engine = 'xlrd' if file.name.endswith('.xls') else 'openpyxl'
        df = pd.read_excel(file, header=None, engine=engine)
        
        mc_data = []
        cd_code = df.iloc[10, 3] if df.shape[0] > 10 else ''
        
        for row in range(20, min(len(df), 100)):
            if all(pd.isna(df.iloc[row, i]) for i in range(4)):
                continue
                
            mc_data.append({
                "CD Code": cd_code,
                "Machine Type": df.iloc[row, 2] if df.shape[1] > 2 else '',
                "S/N#": df.iloc[row, 3] if df.shape[1] > 3 else '',
                "File_name": file.name
            })
            
        return pd.DataFrame(mc_data)
    except Exception as e:
        st.error(f"❌ MC文件处理失败：{file.name} - {str(e)}")
        return None

def process_rel_file(file):
    """处理Relocation文件（四阶段抓取）"""
    try:
        engine = 'xlrd' if file.name.endswith('.xls') else 'openpyxl'
        df = pd.read_excel(file, header=None, engine=engine)
        
        rel_data = []
        from_cd = df.iloc[24, 3] if df.shape[0] > 24 else ''
        to_cd = df.iloc[26, 3] if df.shape[0] > 26 else ''
        
        # 模式1：原始抓取方式
        for row in range(32, min(len(df), 100)):
            if row >= len(df):
                break
            if df.shape[1] > 7 and pd.isna(df.iloc[row, 6]) and pd.isna(df.iloc[row, 7]):
                break
            rel_data.append({
                "From_CD Code": from_cd,
                "To_CD Code": to_cd,
                "Machine Type": df.iloc[row, 1] if df.shape[1] > 1 else '',
                "S/N#": df.iloc[row, 4] if df.shape[1] > 4 else '',
                "File_name": file.name
            })
        
        # 模式2：备用抓取方案
        if len(rel_data) == 0:
            for row in range(33, min(len(df), 100)):
                if row >= len(df):
                    break
                if df.shape[1] > 8 and pd.isna(df.iloc[row, 7]) and pd.isna(df.iloc[row, 8]):
                    break
                machine_type = df.iloc[row, 1] if df.shape[1] > 1 else ''
                sn = df.iloc[row, 4] if df.shape[1] > 4 else ''
                if pd.notna(machine_type) or pd.notna(sn):
                    rel_data.append({
                        "From_CD Code": from_cd,
                        "To_CD Code": to_cd,
                        "Machine Type": machine_type,
                        "S/N#": sn,
                        "File_name": file.name
                    })
        
        # 模式3：终极抓取方案
        if len(rel_data) == 0:
            # B列抓取（B33开始）
            b_data = []
            for row in range(33, min(len(df), 200)):
                if row >= len(df):
                    break
                if df.shape[1] > 8 and pd.isna(df.iloc[row, 7]) and pd.isna(df.iloc[row, 8]):
                    break
                b_value = df.iloc[row, 1] if df.shape[1] > 1 else None
                if pd.notna(b_value):
                    b_data.append({
                        "From_CD Code": from_cd,
                        "To_CD Code": to_cd,
                        "Machine Type": b_value,
                        "S/N#": "",
                        "File_name": file.name
                    })
            
            # E列抓取（E33开始）
            e_data = []
            for row in range(33, min(len(df), 200)):
                if row >= len(df):
                    break
                if df.shape[1] > 8 and pd.isna(df.iloc[row, 7]) and pd.isna(df.iloc[row, 8]):
                    break
                e_value = df.iloc[row, 4] if df.shape[1] > 4 else None
                if pd.notna(e_value):
                    e_data.append({
                        "From_CD Code": from_cd,
                        "To_CD Code": to_cd,
                        "Machine Type": "",
                        "S/N#": e_value,
                        "File_name": file.name
                    })
            
            rel_data.extend(b_data)
            rel_data.extend(e_data)
        
        # 模式4：旧版处理逻辑（从第32行开始逐行扫描）
        if len(rel_data) == 0:
            row = 32
            while row < len(df):
                if row >= len(df):
                    break
                # 检查H列(7)和I列(8)是否同时为空
                if df.shape[1] > 8 and pd.isna(df.iloc[row, 7]) and pd.isna(df.iloc[row, 8]):
                    break
                
                # 提取数据
                machine_type = df.iloc[row, 1] if df.shape[1] > 1 else ''
                sn = df.iloc[row, 4] if df.shape[1] > 4 else ''
                
                if pd.notna(machine_type) or pd.notna(sn):
                    rel_data.append({
                        "From_CD Code": from_cd,
                        "To_CD Code": to_cd,
                        "Machine Type": machine_type,
                        "S/N#": sn,
                        "File_name": file.name
                    })
                row += 1
        
        return pd.DataFrame(rel_data)
    except Exception as e:
        st.error(f"❌ Relocation文件处理失败：{file.name} - {str(e)}")
        return None

def process_stock_file(file):
    """处理Stock Machine文件"""
    try:
        engine = 'xlrd' if file.name.endswith('.xls') else 'openpyxl'
        df = pd.read_excel(file, header=None, engine=engine)
        
        if '二合一' in file.name:
            return process_combined_stock(df, file)
        return process_normal_stock(df, file)
    except Exception as e:
        st.error(f"❌ Stock文件处理失败：{file.name} - {str(e)}")
        return None

def process_normal_stock(df, file):
    """处理普通Stock文件"""
    stock_data = []
    c15 = str(df.iloc[14, 2]) if df.shape[0] > 14 and df.shape[1] > 2 else ''
    matches = re.findall(r'[（(]([^）)]+)[）)]', c15)
    cd_code = matches[-1].strip() if matches else ''
    
    for row in range(20, min(len(df), 100)):
        if df.shape[1] < 10:
            break
            
        if pd.isna(df.iloc[row, 8]) and pd.isna(df.iloc[row, 9]):
            break
            
        b_col = df.iloc[row, 1] if df.shape[1] > 1 else None
        e_col = df.iloc[row, 4] if df.shape[1] > 4 else None
        
        if pd.notna(b_col) or pd.notna(e_col):
            stock_data.append({
                "CD Code": cd_code,
                "Machine Type": b_col,
                "S/N#": e_col,
                "File_name": file.name
            })
            
    return pd.DataFrame(stock_data)

def process_combined_stock(df, file):
    """处理二合一Stock文件"""
    combined_data = []
    cd_end_user = df.iloc[14, 3] if df.shape[0] > 14 and df.shape[1] > 3 else ''
    cd_distributor = df.iloc[15, 3] if df.shape[0] > 15 and df.shape[1] > 3 else ''
    
    for row in range(21, min(len(df), 100)):
        if df.shape[1] < 6:
            break
            
        if pd.isna(df.iloc[row, 9]):
            break
            
        c_col = df.iloc[row, 2] if df.shape[1] > 2 else None
        f_col = df.iloc[row, 5] if df.shape[1] > 5 else None
        
        if pd.notna(c_col) or pd.notna(f_col):
            combined_data.append({
                "CD Code_End User": cd_end_user,
                "CD Code_Distributor": cd_distributor,
                "Machine Type": c_col,
                "S/N#": f_col,
                "File_name": file.name
            })
            
    return pd.DataFrame(combined_data)

# ================== 统一文件处理区 ==================
with st.container(border=True):
    st.subheader("📁 统一数据上传处理区", divider="rainbow")
    
    uploaded_files = st.file_uploader(
        "请上传所有相关文件（可多选）",
        type=['xls', 'xlsx', 'xlsm'],
        accept_multiple_files=True,
        key="unified_uploader"
    )

    if uploaded_files:
        file_processors = {
            'MC': {
                'pattern': re.compile(r'MC Info', re.IGNORECASE),
                'data_key': 'mc_data',
                'count_key': 'mc_success_count'
            },
            'REL': {
                'pattern': re.compile(r'relocation', re.IGNORECASE),
                'data_key': 'rel_data',
                'count_key': 'rel_success_count'
            },
            'STOCK': {
                'pattern': re.compile(r'(Stock Machine|二合一)', re.IGNORECASE),
                'data_key': 'stock_data',
                'count_key': 'stock_success_count'
            }
        }

        for file in uploaded_files:
            if file.name in st.session_state.processed_files:
                st.warning(f"⏩ 已跳过重复文件：{file.name}")
                continue

            file_type = None
            for ft, config in file_processors.items():
                if config['pattern'].search(file.name):
                    file_type = ft
                    break

            if not file_type:
                st.error(f"❌ 无法识别的文件类型：{file.name}")
                continue

            try:
                processor = globals()[f'process_{file_type.lower()}_file']
                processed_data = processor(file)
                
                if processed_data is not None and not processed_data.empty:
                    current_data = st.session_state[file_processors[file_type]['data_key']]
                    if current_data is not None:
                        st.session_state[file_processors[file_type]['data_key']] = pd.concat(
                            [current_data, processed_data], ignore_index=True)
                    else:
                        st.session_state[file_processors[file_type]['data_key']] = processed_data
                    
                    st.session_state.processed_files.add(file.name)
                    st.session_state[file_processors[file_type]['count_key']] += 1
                else:
                    st.warning(f"⚠️ 文件未包含有效数据：{file.name}")
            except Exception as e:
                st.error(f"❌ 处理文件时发生严重错误：{file.name} - {str(e)}")

        # 显示成功统计
        success_col1, success_col2, success_col3 = st.columns(3)
        with success_col1:
            st.info(f"✅ 成功处理MC文件数量：{st.session_state.mc_success_count}")
        with success_col2:
            st.info(f"✅ 成功处理Relocation文件数量：{st.session_state.rel_success_count}")
        with success_col3:
            st.info(f"✅ 成功处理Stock文件数量：{st.session_state.stock_success_count}")

        # 三栏并排显示
        col1, col2, col3 = st.columns(3)
        
        # MC数据展示
        with col1:
            if st.session_state.mc_data is not None and not st.session_state.mc_data.empty:
                st.subheader("MC Info 数据", divider="blue")
                st.dataframe(st.session_state.mc_data, use_container_width=True)
                
                buffer = BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    st.session_state.mc_data.to_excel(writer, index=False)
                    worksheet = writer.sheets["Sheet1"]
                    worksheet.freeze_panes(1, 0)
                    for col_num, col_name in enumerate(st.session_state.mc_data.columns):
                        max_len = max(st.session_state.mc_data[col_name].astype(str).str.len().max(), len(col_name)) + 2
                        worksheet.set_column(col_num, col_num, max_len)
                
                st.download_button(
                    "💾 下载MC数据",
                    data=buffer.getvalue(),
                    file_name="MC_Data.xlsx",
                    mime="application/vnd.ms-excel",
                    use_container_width=True
                )

        # Relocation数据展示
        with col2:
            if st.session_state.rel_data is not None and not st.session_state.rel_data.empty:
                st.subheader("Relocation 数据", divider="orange")
                st.dataframe(st.session_state.rel_data, use_container_width=True)
                
                buffer = BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    st.session_state.rel_data.to_excel(writer, index=False)
                    worksheet = writer.sheets["Sheet1"]
                    worksheet.freeze_panes(1, 0)
                    for col_num, col_name in enumerate(st.session_state.rel_data.columns):
                        max_len = max(st.session_state.rel_data[col_name].astype(str).str.len().max(), len(col_name)) + 2
                        worksheet.set_column(col_num, col_num, max_len)
                
                st.download_button(
                    "💾 下载Relocation数据",
                    data=buffer.getvalue(),
                    file_name="Relocation_Data.xlsx",
                    mime="application/vnd.ms-excel",
                    use_container_width=True
                )

        # Stock数据展示
        with col3:
            if st.session_state.stock_data is not None and not st.session_state.stock_data.empty:
                st.subheader("Stock 数据", divider="violet")
                st.dataframe(st.session_state.stock_data, use_container_width=True)
                
                buffer = BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    st.session_state.stock_data.to_excel(writer, index=False)
                    worksheet = writer.sheets["Sheet1"]
                    worksheet.freeze_panes(1, 0)
                    for col_num, col_name in enumerate(st.session_state.stock_data.columns):
                        max_len = max(st.session_state.stock_data[col_name].astype(str).str.len().max(), len(col_name)) + 2
                        worksheet.set_column(col_num, col_num, max_len)
                
                st.download_button(
                    "💾 下载Stock数据",
                    data=buffer.getvalue(),
                    file_name="Stock_Data.xlsx",
                    mime="application/vnd.ms-excel",
                    use_container_width=True
                )

# ================== 整合下载区 ==================
if any(st.session_state[key] is not None for key in ['mc_data', 'rel_data', 'stock_data']):
    st.divider()
    with st.container(border=True):
        st.subheader("🚀 数据整合下载区", divider="green")

        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            sheets = {
                "MC Info": st.session_state.mc_data,
                "Relocation": st.session_state.rel_data,
                "STOCK MACHINE SHIPPING INFO": st.session_state.stock_data[
                    st.session_state.stock_data['File_name'].str.contains('二合一') == False] 
                    if st.session_state.stock_data is not None else pd.DataFrame(),
                "二合一STOCK MACHINE SHIPPING INFO": st.session_state.stock_data[
                    st.session_state.stock_data['File_name'].str.contains('二合一')] 
                    if st.session_state.stock_data is not None else pd.DataFrame()
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
            key="unique_orange_btn"  # 唯一标识符
        )

# ================== 页面样式优化 ==================
st.markdown("""
<style>
/* 精准定位完整整合报告按钮 */
div[data-testid="stDownloadButton"] button[data-testid="baseButton-unique_orange_btn"] {
    background: #FFA500 !important;  /* 纯橙色背景 */
    border: 2px solid #FF8C00 !important;  /* 深橙色边框 */
    color: #000000 !important;  /* 黑色文字 */
    font-weight: bold;
}

/* 悬停状态 */
div[data-testid="stDownloadButton"] button[data-testid="baseButton-unique_orange_btn"]:hover {
    background: #FF8C00 !important;
    border-color: #FF6B00 !important;
}

/* 按下状态 */
div[data-testid="stDownloadButton"] button[data-testid="baseButton-unique_orange_btn"]:active {
    background: #FF6B00 !important;
    border-color: #FF4500 !important;
}

/* 其他按钮保持绿色 */
.stDownloadButton button:not([data-testid="baseButton-unique_orange_btn"]) {
    background: linear-gradient(45deg, #32CD32, #228B22) !important;
    color: white !important;
}
</style>
""", unsafe_allow_html=True)

#####################
#我想設計一個精美的python的streamlit app解決，要有title和subheader，app用簡單中文

#我有以下3個部分想進行:

#1. 在左邊建立一個excel的uploader，容許用家upload多於一個xlsx或xlsm格式的excel，當用家upload完成並確認後，先檢查用家upload的excel是否檔案名字內都有"MC Info"字眼，例如"MC Info ABC"便是合格，"Info ABC"則是不合格的檔名。如果任一檔案名稱不合格的話就提醒一下用家並讓他重新upload，如果所有檔案名稱都合格的話就進行以下動作:

#把已upload的excel各自的以下條件位置的內容全部抽出並以表格顯示出來，列名分別為”CD Code”和”Machine Type”和”S/N#”，sheet名為”MC Info”:

#-D1格子內容
#-C列內容(C21開始一直向下，直到對應的B列沒有內容為止的C列內容)
#-D列內容(D21開始一直向下，直到對應的B列沒有內容為止的D列內容)

#最後在subheader旁加一個download button容易用家download已整合的資料表格。

#2. 在右邊建立另一個excel的uploader，容許用家upload多於一個xlsx或xlsm格式的excel，當用家upload完成並確認後，先檢查用家upload的excel是否檔案名字內都有"relocation"字眼，例如"MC Info Relocation_sheet"或者”HuaYun-★Relocation_sheet-Y53636”便是合格，"HuaYun-★Y53636"則是不合格的檔名。如果任一檔案名稱不合格的話就提醒一下用家並讓他重新upload，如果所有檔案名稱都合格的話就進行以下動作:

#把已upload的excel各自的以下條件位置的內容全部抽出並以表格顯示出來，列名分別為”From_CD Code”和”To_CD Code”和”Machine Type”和”S/N#”，sheet名為”Relocation”:

#-D25格子內容
#-D27格子內容
#1. B列內容(B33開始一直向下的B列內容，直到對應的I列和H列首次同時沒有內容為止，只要對應的I列或H列任一有內容都可，只是當對應的I列和H列首次同時沒有內容才停止)
#2. E列內容(E33開始一直向下的E列內容，直到對應的I列和H列首次同時沒有內容為止，只要對應的I列或H列任一有內容都可，只是當對應的I列和H列首次同時沒有內容才停止)


#在抓取"relocation"檔案資料時，非固定內容的抓取位置保持現有的做法作為優先做法，如果找不到可提取資料，就試用以下第二方案抓取：
#1. B列內容(B34開始一直向下的B列內容，直到對應的I列和H列首次同時沒有內容為止，只要對應的I列或H列任一有內容都可，只是當對應的I列和H列首次同時沒有內容才停止)
#2. E列內容(E34開始一直向下的E列內容，直到對應的I列和H列首次同時沒有內容為止，只要對應的I列或H列任一有內容都可，只是當對應的I列和H列首次同時沒有內容才停止)
#在讀取"relocation"檔案時，其中一個檔案如果出現读取失敗ERROR: index 8 is out of bounds for axis 0 with size 8，除了要解決讀取問題，在已正常讀取亦要通知用家已正常讀取。

#在抓取"relocation"檔案資料時，非固定內容的抓取位置的做法，如果第一和第二個方案都找不到可提取資料，就試用以下第三方案抓取：
#1. B列內容(B34開始一直向下的B列內容，直到對應的I列和H列首次同時沒有內容為止，只要對應的I列或H列任一有內容都可，只是當對應的I列和H列首次同時沒有內容才停止)
#2. E列內容(E34開始一直向下的E列內容，直到對應的I列和H列首次同時沒有內容為止，只要對應的I列或H列任一有內容都可，只是當對應的I列和H列首次同時沒有內容才停止)




#最後在subheader旁加一個download button容易用家download已整合的資料表格。


#3. 當以上兩個動作都已完成，就在不影響以上兩個表格的情況下，在置中位置加一個名為"Combine"的download button讓用家download已整合的1和2部分的excel workbook，分兩張sheet放在同一個workbook就可以

#在處理任何sheet時，如發現用家同一個檔案重複上載了，該重複的檔就只提取一次數據，不用重複提取，並通知用家該檔重複上傳了


#- 由於現在已經有第1和第2部分的兩組file uploader，這兩組都不用再改。我現想再另外建立第3部分的excel uploader，同樣容許用家upload多於一個xlsx或xlsm式xls格式的excel，當用家upload完成並確認後，先檢查用家upload的excel是否檔案名字內是否有"Stock Machine"或"二合一"字眼，例如"二合一 STOCK MACHINE SHIPPPING INFORMATION-雷特"或者”ESE HK-Stock machine shipping information（SuZhou Bako)”便是合格的檔名，"HuaYun-★Y53636"則是不合格的檔名。如果任一檔案名稱不合格的話就提醒一下用家並讓他重新upload，當用家upload的檔案裡如果夾集excel以外的文件如word或pdf檔等等時，就像之前第1和第2部分的兩組file uploader那樣無視該些非excel的文檔即可，繼續操作。如果所有檔案名稱都合格的話就進行以下動作:

#把已upload的excel先分為兩部分，第一部分是檔名裡沒有"二合一"字眼的，例如: "ESE HK-Stock machine shipping information（SuZhou Bako)"; 第二部分是檔名裡有"二合一"字眼的，例如: "二合一 STOCK MACHINE SHIPPPING INFORMATION-雷特";。

#第一部分的數據提取，就是以下條件位置的內容全部抽出並以表格顯示出來，列名分別為"CD Code"和 "Machine Type"和"S/N#"，sheet名為"STOCK MACHINE SHIPPING INFO":

#-C15格子內，找出有括號內的內容，例如: C15是"SuZhou Bako Optoelectronics Co.,Ltd.（#05910）No.2 Xinfeng East Road, Yangshe Town, Zhangjiagang City, Suzhou City"，那麼就只需提取#05910就可以
#-B列內容(B21開始一直向下，直到對應的I列和J列同時沒有內容為止的B列內容)
#-E列內容(E21開始一直向下，直到對應的I列和J列同時沒有內容為止的E列內容)

#第一部分的表格可獨立供用家download

#第二部分的數據提取，就是以下條件位置的內容全部抽出並以表格顯示出來，列名分別為"CD Code_End User"和"CD Code_Distributor"和"Machine Type"和"S/N#"，sheet名為"二合一STOCK MACHINE SHIPPING INFO":

#-D15格子內容
#-D16格子內容
#-C列內容(C22開始一直向下，直到對應的J列沒有內容為止的C列內容)
#-F列內容(F22開始一直向下，直到對應的J列沒有內容為止的F列內容)

#第二部分的表格可獨立供用家download

#另外:
#-每一個用家download的表格內的每一張sheet中，不用顯示A列至D列都沒有內容的行數
##-每一個用家download的表格內的每一張sheet，都在E列加一新列，列名為File_name，這列每一行的內容就是對應行數的數據的來源excel檔案名稱，假設該行的數據來源檔案的名稱是"MC Info Sheet－ Keweixin"，就顯示MC Info Sheet－ Keweixin

#當以上動作都已完成後，就在不影響以上所有表格的情況下，本來在置中位置的整合download button改為讓用家download已整合三個excel uploader資料表的excel workbook，分四張sheet放在同一個workbook就可以，用家download的表格內的每一張sheet的第一行都要freeze掉，並且A至D列都要對整齊。

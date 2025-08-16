import streamlit as st
import pandas as pd
import os
import re
from collections import Counter

# --- 这是我们之前打造的“究极体”核心分析逻辑，现在把它封装成一个函数 ---
def analyze_excel_file(uploaded_file):
    """
    接收一个上传的Excel文件，并返回分析结果字符串。
    """
    # --- 楼栋房型代码规则 ---
    jinling_room_types = [
        'DETN', 'DKN', 'DKS', 'DQN', 'DQS', 'DSKN', 'DSTN', 'DTN',
        'EKN', 'EKS', 'ESN', 'ESS', 'ETN', 'ETS', 'FSB', 'FSC', 'FSN',
        'STN', 'STS', 'SKN', 'RSN', 'SQS', 'SQN'
    ]
    yatai_room_types = [
        'JDEN', 'JDKN', 'JDKS', 'JEKN', 'JESN', 'JESS', 'JETN', 'JETS',
        'JKN', 'JLKN', 'JTN', 'JTS', 'VCKD', 'VCKN'
    ]
    # --- 规则结束 ---

    unknown_codes_collection = Counter()
    file_base_name = os.path.splitext(uploaded_file.name)[0]

    try:
        # 1. 智能解析
        df_raw = pd.read_excel(uploaded_file, header=None, dtype=str)
        all_bookings = []
        current_group_name = "未知团队"
        current_market_code = "无"
        column_map = {}
        header_row_index = -1

        for index, row in df_raw.iterrows():
            row_str = ' '.join(cell.strip() for cell in row.dropna())
            if '团体/单位/旅行社/订房中心：' in row_str:
                match = re.search(r'团体/单位/旅行社/订房中心：([^,]+)', row_str)
                if match: current_group_name = match.group(1).strip()
                column_map, header_row_index, current_market_code = {}, -1, "无"
                continue
            if '市场码：' in row_str:
                match = re.search(r'市场码：(\w+)', row_str)
                if match: current_market_code = match.group(1).strip()
            if '房号' in row_str and '姓名' in row_str and '人数' in row_str:
                header_row_index = index
                for i, col in enumerate(row):
                    if pd.notna(col): column_map[re.sub(r'\s+', '', str(col))] = i
                continue
            if header_row_index != -1 and index > header_row_index and not row.dropna().empty:
                all_bookings.append({'团队名称': current_group_name, '市场码': current_market_code, 'data': row})

        if not all_bookings: return f"【{file_base_name}】处理失败：未能在文件中解析出任何有效的数据行。"

        # 2. 转换为DataFrame
        processed_rows = []
        for item in all_bookings:
            row_data = item['data']
            processed_row = {'团队名称': item['团队名称'], '市场码': item['市场码']}
            for col_name, col_index in column_map.items():
                processed_row[col_name] = row_data.get(col_index)
            processed_rows.append(processed_row)
        df = pd.DataFrame(processed_rows)

        # 3. 执行过滤和统计
        df['状态'] = df['状态'].astype(str).str.strip()
        df_active = df[df['状态'] == 'R'].copy()
        df_counted = df_active[~df_active['团队名称'].str.contains('WA|FIT', case=False, na=False)].copy()
        
        df_counted['房数'] = pd.to_numeric(df_counted['房数'], errors='coerce').fillna(0)
        df_counted['人数'] = pd.to_numeric(df_counted['人数'], errors='coerce').fillna(0)
        df_counted['房类'] = df_counted['房类'].astype(str).str.strip()

        total_rooms = int(df_counted['房数'].sum())
        total_guests = int(df_counted['人数'].sum())

        def assign_building(room_type):
            if room_type in yatai_room_types: return '亚太楼'
            elif room_type in jinling_room_types: return '金陵楼'
            else:
                if room_type and room_type.lower() != 'nan': unknown_codes_collection.update([room_type])
                return '其他楼'
        df_counted['准确楼栋'] = df_counted['房类'].apply(assign_building)

        con_df = df_counted[df_counted['团队名称'].str.contains('CON', case=False, na=False)].copy()
        con_group_count = int(con_df['团队名称'].nunique())
        total_con_rooms = int(con_df['房数'].sum())
        con_jinling_rooms = int(con_df[con_df['准确楼栋'] == '金陵楼']['房数'].sum())
        con_yatai_rooms = int(con_df[con_df['准确楼栋'] == '亚太楼']['房数'].sum())
        con_other_rooms = int(con_df[con_df['准确楼栋'] == '其他楼']['房数'].sum())

        gto_df = df_counted[df_counted['市场码'] == 'GTO'].copy()
        gto_group_count = int(gto_df['团队名称'].nunique())
        total_gto_rooms = int(gto_df['房数'].sum())
        total_gto_guests = int(gto_df['人数'].sum())
        gto_jinling_rooms = int(gto_df[gto_df['准确楼栋'] == '金陵楼']['房数'].sum())
        gto_yatai_rooms = int(gto_df[gto_df['准确楼栋'] == '亚太楼']['房数'].sum())
        gto_other_rooms = int(gto_df[gto_df['准确楼栋'] == '其他楼']['房数'].sum())

        # 4. 生成报告
        summary_parts = [f"【{file_base_name}】: 有效总房数 {total_rooms} 间 (共 {total_guests} 人)"]
        if con_group_count > 0:
            con_report = f"CON团队房({con_group_count}个团队, 共{total_con_rooms}间)分布: 金陵楼 {con_jinling_rooms} 间, 亚太楼 {con_yatai_rooms} 间"
            if con_other_rooms > 0: con_report += f", 其他楼 {con_other_rooms} 间"
            summary_parts.append(f"，其中{con_report}.")
        else:
            summary_parts.append("，(无CON团队房).")
        
        if total_gto_rooms > 0:
            gto_report = f"旅行社(GTO)房({gto_group_count}个团队, {total_gto_rooms}间, 共{total_gto_guests}人)分布: 金陵楼 {gto_jinling_rooms} 间, 亚太楼 {gto_yatai_rooms} 间"
            if gto_other_rooms > 0: gto_report += f", 其他楼 {gto_other_rooms} 间"
            summary_parts.append(f" | {gto_report}.")
        else:
            summary_parts.append(" | (无GTO旅行社房).")

        final_report = "".join(summary_parts)

        if unknown_codes_collection:
            final_report += "\n\n--- 侦测到的未知房型代码 ---"
            for code, count in unknown_codes_collection.items():
                final_report += f"\n代码: '{code}' (出现了 {count} 次)"
        
        return final_report

    except Exception as e:
        return f"【{file_base_name}】处理失败，错误: {e}"


# --- Streamlit 网页界面 ---

st.set_page_config(page_title="酒店团队报表智能分析工具", layout="wide")

st.title("酒店团队报表智能分析工具 🤖")
st.markdown("##### Bro，把你的Excel报表扔进来，剩下的交给我！")

uploaded_files = st.file_uploader(
    "上传您的Excel报表 (可一次上传多个)", 
    type=['xlsx'], 
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("开始分析！"):
        with st.spinner('引擎启动，正在玩命分析中...'):
            st.write("---")
            st.subheader("分析结果报告:")
            for file in uploaded_files:
                # 对每个上传的文件，调用我们的核心分析函数
                result_string = analyze_excel_file(file)
                st.text(result_string) # st.text 可以很好地显示多行文本
        st.success("分析完成！")
        st.balloons()

st.sidebar.header("关于这个工具")
st.sidebar.info(
    "这是一个智能报表解析工具，能自动处理格式不统一、包含多个业务板块的复杂Excel文件。"
    "它由Python和Streamlit强力驱动，旨在将您从繁琐的手工统计中解放出来。"
)

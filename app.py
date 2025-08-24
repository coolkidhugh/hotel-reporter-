import pandas as pd
import os
import re
from collections import Counter

def analyze_reports_ultimate():
    """
    智能解析并动态定位列，对包含多个团队的Excel报告进行详细统计。
    (究极体：v3 - 根据报告类型动态筛选R/I/O状态)
    """
    reports_folder_path = 'reports'

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
    final_summary_lines = []

    print("--- 启动【究极体】分析引擎 (v3-动态状态版) ---")

    if not os.path.isdir(reports_folder_path):
        print(f"错误：未找到名为 '{reports_folder_path}' 的文件夹。")
        print("请在脚本所在的目录中创建一个 'reports' 文件夹，并将Excel文件放入其中。")
        return

    files_to_analyze = [f for f in os.listdir(reports_folder_path) if f.endswith('.xlsx') and not f.startswith('~$')]
    
    if not files_to_analyze:
        print(f"在 '{reports_folder_path}' 文件夹中没有找到任何有效的 .xlsx 文件。")
        return

    for filename in files_to_analyze:
        file_base_name = os.path.splitext(filename)[0]
        try:
            df_raw = pd.read_excel(os.path.join(reports_folder_path, filename), header=None, dtype=str)
            all_bookings = []
            current_group_name = "未知团队"
            current_market_code = "无"
            column_map = {}
            header_row_index = -1

            for index, row in df_raw.iterrows():
                row_str = ' '.join(str(cell).strip() for cell in row.dropna() if str(cell).strip())
                if not row_str:
                    continue

                if '团体名称:' in row_str:
                    match = re.search(r'团体名称:\s*(.*?)(?:\s*市场码：|$)', row_str)
                    if match:
                        current_group_name = match.group(1).strip()
                    else:
                        current_group_name = "未知团队(解析失败)"
                        
                    column_map, header_row_index, current_market_code = {}, -1, "无"
                    
                    market_match = re.search(r'市场码：(\w+)', row_str)
                    if market_match:
                        current_market_code = market_match.group(1).strip()
                    continue
                
                if '团体/单位/旅行社/订房中心：' in row_str:
                    desc_match = re.search(r'团体/单位/旅行社/订房中心：(.*)', row_str)
                    if desc_match and desc_match.group(1):
                        current_group_name += " " + desc_match.group(1).strip()
                    continue

                if '市场码：' in row_str and not '团体名称:' in row_str:
                    match = re.search(r'市场码：(\w+)', row_str)
                    if match:
                        current_market_code = match.group(1).strip()
                    continue

                if '房号' in row_str and '姓名' in row_str and '人数' in row_str:
                    header_row_index = index
                    for i, col in enumerate(row):
                        if pd.notna(col):
                            column_map[re.sub(r'\s+', '', str(col))] = i
                    continue

                if header_row_index != -1 and index > header_row_index and not row.dropna().empty:
                    if '小计' not in row_str:
                         all_bookings.append({'团队名称': current_group_name, '市场码': current_market_code, 'data': row})
            
            if not all_bookings:
                # 文件中可能确实没有数据，这不是一个错误，只是一个空的报告
                print(f"注意：文件【{file_base_name}】中未解析到有效预订数据行。")
                final_summary_lines.append(f"【{file_base_name}】: 有效总房数 0 间 (共 0 人)，(无CON/会议团队房). | (无GTO旅行社房).")
                continue # 继续处理下一个文件

            processed_rows = []
            for item in all_bookings:
                row_data = item['data']
                processed_row = {'团队名称': item['团队名称'], '市场码': item['市场码']}
                for col_name, col_index in column_map.items():
                    processed_row[col_name] = row_data.get(col_index)
                processed_rows.append(processed_row)
            df = pd.DataFrame(processed_rows)

            df['状态'] = df['状态'].astype(str).str.strip()
            
            # --- V V V 核心修改：根据文件名动态设置需要统计的状态 V V V ---
            if '在住' in file_base_name:
                valid_statuses = ['R', 'I']
            elif '离店' in file_base_name:
                valid_statuses = ['I', 'R', 'O']
            else:
                valid_statuses = ['R']
            
            print(f"--- 报告【{file_base_name}】将统计状态为 {valid_statuses} 的房间 ---")
            
            # 使用 .isin() 方法来筛选一个列表内的所有状态
            df_active = df[df['状态'].isin(valid_statuses)].copy()
            # --- ^ ^ ^ 核心修改结束 ^ ^ ^ ---

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
                    if room_type and room_type.lower() != 'nan':
                        unknown_codes_collection.update([room_type])
                    return '其他楼'
            df_counted['准确楼栋'] = df_counted['房类'].apply(assign_building)

            con_df = df_counted[df_counted['团队名称'].str.contains('CON|会议', case=False, na=False)].copy()
            
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

            summary_parts = [f"【{file_base_name}】: 有效总房数 {total_rooms} 间 (共 {total_guests} 人)"]

            if con_group_count > 0:
                con_report = f"CON/会议团队房({con_group_count}个团队, 共{total_con_rooms}间)分布: 金陵楼 {con_jinling_rooms} 间, 亚太楼 {con_yatai_rooms} 间"
                if con_other_rooms > 0: con_report += f", 其他楼 {con_other_rooms} 间"
                summary_parts.append(f"，其中{con_report}.")
            else:
                summary_parts.append("，(无CON/会议团队房).")

            if total_gto_rooms > 0:
                gto_report = f"旅行社(GTO)房({gto_group_count}个团队, {total_gto_rooms}间, 共{total_gto_guests}人)分布: 金陵楼 {gto_jinling_rooms} 间, 亚太楼 {gto_yatai_rooms} 间"
                if gto_other_rooms > 0: gto_report += f", 其他楼 {gto_other_rooms} 间"
                summary_parts.append(f" | {gto_report}.")
            else:
                summary_parts.append(" | (无GTO旅行社房).")

            final_summary_lines.append("".join(summary_parts))

        except Exception as e:
            final_summary_lines.append(f"【{file_base_name}】处理失败，错误: {e}")

    print("\n" + "\n".join(final_summary_lines))
    if unknown_codes_collection:
        print("\n--- 侦测到的未知房型代码 (请检查是否需要更新规则) ---")
        for code, count in unknown_codes_collection.items():
            print(f"代码: '{code}' (出现了 {count} 次)")
    print("\n--- 总结生成完毕 ---")


if __name__ == "__main__":
    analyze_reports_ultimate()

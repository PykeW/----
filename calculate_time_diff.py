import pandas as pd
from datetime import datetime

def parse_time(time_str):
    if pd.isna(time_str):
        return None
    # 处理两种可能的时间格式
    try:
        # 尝试解析已有的时间格式
        return pd.to_datetime(time_str)
    except:
        try:
            # 尝试解析其他可能的格式
            time_str = str(time_str).replace('GMT+8', '').strip()
            if '年' in time_str:
                return pd.to_datetime(time_str, format='%Y年%m月%d日 %H:%M:%S')
            else:
                return pd.to_datetime(time_str)
        except:
            return None

def calculate_days(start_time, end_time, title=""):
    if pd.isna(start_time) or pd.isna(end_time):
        print(f"  {title}: 无法计算时间差 - 开始时间: {start_time}, 结束时间: {end_time}")
        return None
    diff = end_time - start_time
    days = round(diff.total_seconds() / (24 * 3600), 2)
    print(f"  {title}: {days} ({start_time} -> {end_time})")
    return days

def process_excel_data():
    try:
        # 读取Excel文件的原始数据表
        excel_file = 'git发布到内测完成用时.xlsx'
        print(f"\n正在读取Excel文件: {excel_file}")
        df_source = pd.read_excel(excel_file, sheet_name='原始数据')
        
        print("\n原始数据表的列名:")
        print(df_source.columns.tolist())
        
        # 检查必要的列是否存在
        required_columns = ['Title', 'URL', 'State', 'Assignee', '紧急程度', '录入', '开始', '待验收', '关闭']
        missing_columns = [col for col in required_columns if col not in df_source.columns]
        if missing_columns:
            print(f"\n警告: 以下列不存在: {missing_columns}")
            return
        
        # 解析时间
        entry_times = df_source['录入'].apply(parse_time)
        start_times = df_source['开始'].apply(parse_time)
        validation_times = df_source['待验收'].apply(parse_time)
        close_times = df_source['关闭'].apply(parse_time)
        
        # 计算各阶段时间差（天）
        entry_to_start = []
        start_to_validation = []
        validation_to_close = []
        
        print("\n开始计算时间差:")
        for i, title in enumerate(df_source['Title']):
            print(f"\n处理任务 {i+1}: {title}")
            entry_to_start.append(calculate_days(entry_times[i], start_times[i], "录入-开始"))
            start_to_validation.append(calculate_days(start_times[i], validation_times[i], "开始-待验收"))
            validation_to_close.append(calculate_days(validation_times[i], close_times[i], "待验收-关闭"))
        
        # 创建新的DataFrame用于用时计算表
        calculation_df = pd.DataFrame({
            'Title': df_source['Title'],
            'URL': df_source['URL'],
            'State': df_source['State'],
            'Assignee': df_source['Assignee'],
            '紧急程度': df_source['紧急程度'],
            '录入-开始': [f"{diff}" if diff is not None else None for diff in entry_to_start],
            '开始-待验收': [f"{diff}" if diff is not None else None for diff in start_to_validation],
            '待验收-关闭': [f"{diff}" if diff is not None else None for diff in validation_to_close]
        })
        
        print("\n数据处理完成，准备写入'用时计算'表")
        
        # 将数据写入Excel文件的"用时计算"表
        with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            calculation_df.to_excel(writer, sheet_name='用时计算', index=False)
        
        print(f"\nExcel文件已更新: {excel_file}")
        print("-" * 50)
        print("数据已写入'用时计算'表")
        print(f"共处理 {len(calculation_df)} 条记录")
        
    except Exception as e:
        print(f"\n错误: {str(e)}")
        print("\n详细错误信息:")
        import traceback
        print(traceback.format_exc())

if __name__ == "__main__":
    process_excel_data() 
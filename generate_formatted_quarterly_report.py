#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成格式化的季度工时统计报告
按照要求的表格格式输出
"""

import pandas as pd
import chardet

def load_data(file_path):
    """加载CSV数据"""
    # 检测编码
    with open(file_path, 'rb') as f:
        raw_data = f.read()
        result = chardet.detect(raw_data)
        encoding = result['encoding']
    
    # 读取数据
    df = pd.read_csv(file_path, encoding=encoding)
    return df

def get_quarter(week):
    """根据周次获取季度"""
    if 1 <= week <= 13:
        return 1
    elif 14 <= week <= 26:
        return 2
    elif 27 <= week <= 39:
        return 3
    elif 40 <= week <= 52:
        return 4
    else:
        return None

def merge_departments(df):
    """合并T1和T1电子元件部门"""
    df_copy = df.copy()
    df_copy['订单项目.归属中心'] = df_copy['订单项目.归属中心'].replace('T1电子元件', 'T1')
    return df_copy

def generate_formatted_quarterly_report(df):
    """生成格式化的季度统计报告"""
    print("正在生成格式化季度工时统计报告...")
    
    # 合并部门
    df_merged = merge_departments(df)
    
    # 过滤掉空值
    df_clean = df_merged.dropna(subset=['订单项目.归属中心', '订单项目.立项项目', '订单项目.本周投入天数（最低半天）', '周报人'])
    
    # 添加季度列
    df_clean = df_clean.copy()
    df_clean['季度'] = df_clean['周次'].apply(get_quarter)
    
    # 过滤掉无效季度
    df_clean = df_clean[df_clean['季度'].notna()]
    
    # 按部门、项目、季度分组，计算总人天和人员详情
    result_list = []
    
    # 获取所有部门、项目、季度的组合
    groups = df_clean.groupby(['订单项目.归属中心', '订单项目.立项项目', '季度'])
    
    for (dept, project, quarter), group in groups:
        # 计算总人天
        total_days = group['订单项目.本周投入天数（最低半天）'].sum()
        
        # 获取人员详情
        person_details = group.groupby('周报人')['订单项目.本周投入天数（最低半天）'].sum().sort_values(ascending=False)
        
        # 构建人员和人天字符串
        person_list = []
        person_days_list = []
        
        for person, days in person_details.items():
            person_list.append(person)
            person_days_list.append(f"{days:.1f}")
        
        result_list.append({
            '订单项目.归属中心': dept,
            '订单项目.立项项目': project,
            '季度': int(quarter),
            '总人天': f"{total_days:.1f}",
            '人员': '; '.join(person_list),
            '人天': '; '.join(person_days_list)
        })
    
    # 转换为DataFrame并排序
    result_df = pd.DataFrame(result_list)
    result_df = result_df.sort_values(['订单项目.归属中心', '订单项目.立项项目', '季度'])
    
    # 保存为CSV文件
    output_file = '格式化季度工时统计报告.csv'
    result_df.to_csv(output_file, index=False, encoding='utf-8-sig')
    print(f"✅ 格式化季度统计报告已保存到: {output_file}")
    
    # 打印格式化表格
    print("\n" + "=" * 150)
    print("📊 格式化季度工时统计报告")
    print("=" * 150)
    
    print(f"{'部门':<25} {'项目':<50} {'季度':<6} {'总人天':<8} {'人员':<30} {'人天':<30}")
    print("-" * 150)
    
    for _, row in result_df.iterrows():
        dept = row['订单项目.归属中心']
        project = row['订单项目.立项项目']
        quarter = f"第{row['季度']}季度"
        total_days = row['总人天']
        persons = row['人员']
        person_days = row['人天']
        
        # 处理长文本截断
        if len(project) > 48:
            project = project[:45] + "..."
        if len(persons) > 28:
            persons = persons[:25] + "..."
        if len(person_days) > 28:
            person_days = person_days[:25] + "..."
        
        print(f"{dept:<25} {project:<50} {quarter:<6} {total_days:<8} {persons:<30} {person_days:<30}")
    
    print("-" * 150)
    
    # 生成汇总统计
    print(f"\n📈 季度汇总统计:")
    quarter_summary = df_clean.groupby('季度').agg({
        '订单项目.本周投入天数（最低半天）': 'sum',
        '订单项目.立项项目': 'nunique',
        '周报人': 'nunique'
    }).reset_index()
    
    for _, row in quarter_summary.iterrows():
        quarter = int(row['季度'])
        total_days = row['订单项目.本周投入天数（最低半天）']
        projects = row['订单项目.立项项目']
        people = row['周报人']
        print(f"   第{quarter}季度: 总工时{total_days:.1f}天 | 项目{projects}个 | 参与人员{people}人")
    
    # 部门汇总
    print(f"\n🏢 部门季度汇总:")
    dept_quarter_summary = df_clean.groupby(['订单项目.归属中心', '季度'])['订单项目.本周投入天数（最低半天）'].sum().reset_index()
    
    for dept in df_clean['订单项目.归属中心'].unique():
        dept_data = dept_quarter_summary[dept_quarter_summary['订单项目.归属中心'] == dept]
        print(f"\n   【{dept}】:")
        for _, row in dept_data.iterrows():
            quarter = int(row['季度'])
            total_days = row['订单项目.本周投入天数（最低半天）']
            print(f"     第{quarter}季度: {total_days:.1f}天")
    
    return result_df

def main():
    """主函数"""
    file_path = '2025年1-6.csv'
    
    try:
        # 加载数据
        print("正在加载数据...")
        df = load_data(file_path)
        
        # 生成格式化季度报告
        result = generate_formatted_quarterly_report(df)
        
        print("\n" + "=" * 80)
        print("✅ 格式化季度报告生成完成！")
        print("=" * 80)
        
    except Exception as e:
        print(f"❌ 生成报告过程中出现错误: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()

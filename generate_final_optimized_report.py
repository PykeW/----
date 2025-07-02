#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成最终优化格式的季度工时统计报告
从原始周报CSV文件直接生成最终优化格式的季度统计报告
包含项目总人天列，所有单元格填入具体数值，便于Excel手动合并
"""

import pandas as pd
import chardet

def load_raw_data(file_path):
    """加载原始周报CSV数据"""
    try:
        # 尝试不同编码读取
        for encoding in ['utf-8-sig', 'utf-8', 'gbk', 'gb2312']:
            try:
                df = pd.read_csv(file_path, encoding=encoding)
                print(f"成功使用 {encoding} 编码读取文件")
                return df
            except:
                continue

        # 如果都失败，使用chardet检测
        with open(file_path, 'rb') as f:
            raw_data = f.read()
            result = chardet.detect(raw_data)
            encoding = result['encoding']

        df = pd.read_csv(file_path, encoding=encoding)
        print(f"使用检测到的编码 {encoding} 读取文件")
        return df

    except Exception as e:
        print(f"读取文件失败: {e}")
        return None

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

def process_raw_data_to_quarterly(df):
    """将原始数据处理为季度格式"""
    print("正在处理原始数据...")

    # 合并部门
    df_merged = merge_departments(df)

    # 过滤掉空值
    df_clean = df_merged.dropna(subset=['订单项目.归属中心', '订单项目.立项项目', '订单项目.本周投入天数（最低半天）', '周报人'])

    # 添加季度列
    df_clean = df_clean.copy()
    df_clean['季度'] = df_clean['周次'].apply(get_quarter)

    # 过滤掉无效季度
    df_clean = df_clean[df_clean['季度'].notna()]

    # 按部门、项目、季度、人员分组统计
    quarterly_stats = df_clean.groupby(['订单项目.归属中心', '订单项目.立项项目', '季度', '周报人'])['订单项目.本周投入天数（最低半天）'].sum().reset_index()
    quarterly_stats.columns = ['订单项目.归属中心', '订单项目.立项项目', '季度', '人员', '人天']

    # 计算每个项目每个季度的总人天
    project_quarter_total = df_clean.groupby(['订单项目.归属中心', '订单项目.立项项目', '季度'])['订单项目.本周投入天数（最低半天）'].sum().reset_index()
    project_quarter_total.columns = ['订单项目.归属中心', '订单项目.立项项目', '季度', '季度总人天']

    # 合并数据
    result = quarterly_stats.merge(project_quarter_total, on=['订单项目.归属中心', '订单项目.立项项目', '季度'])

    # 排序
    result = result.sort_values(['订单项目.归属中心', '订单项目.立项项目', '季度', '人天'], ascending=[True, True, True, False])

    return result



def generate_final_optimized_report(quarterly_df):
    """生成最终优化格式的报告"""
    print("正在生成最终优化格式的季度工时统计报告...")

    # 计算每个项目的总人天（跨所有季度）- 需要去重相同的季度总人天
    project_quarter_unique = quarterly_df[['订单项目.归属中心', '订单项目.立项项目', '季度', '季度总人天']].drop_duplicates()
    project_totals = project_quarter_unique.groupby(['订单项目.归属中心', '订单项目.立项项目'])['季度总人天'].sum().reset_index()
    project_totals_dict = {}
    for _, row in project_totals.iterrows():
        key = (row['订单项目.归属中心'], row['订单项目.立项项目'])
        project_totals_dict[key] = row['季度总人天']

    # 创建结果列表
    result_rows = []

    # 处理每一行数据，确保所有单元格都有值
    for _, row in quarterly_df.iterrows():
        dept = row['订单项目.归属中心']
        project = row['订单项目.立项项目']
        quarter = int(row['季度'])
        quarter_total = row['季度总人天']
        person = row['人员']
        person_days = row['人天']

        # 获取项目总人天
        project_key = (dept, project)
        project_total = project_totals_dict.get(project_key, 0)

        # 格式化数据
        quarter_display = f"第{quarter}季度"
        project_total_str = f"{project_total:.1f}"
        quarter_total_str = f"{quarter_total:.1f}"
        person_days_str = f"{person_days:.1f}"

        result_rows.append({
            '订单项目.归属中心': dept,
            '订单项目.立项项目': project,
            '项目总人天': project_total_str,
            '季度': quarter_display,
            '季度总人天': quarter_total_str,
            '人员': person,
            '人天': person_days_str
        })

    # 转换为DataFrame
    result_df = pd.DataFrame(result_rows)

    return result_df

def save_final_report(df, output_file):
    """保存最终优化格式的报告"""
    # 保存为CSV文件，使用UTF-8-BOM编码
    df.to_csv(output_file, index=False, encoding='utf-8-sig')
    print(f"✅ 最终优化格式报告已保存到: {output_file}")
    
    # 打印预览
    print("\n" + "=" * 140)
    print("📊 最终优化格式季度工时统计报告预览")
    print("=" * 140)
    
    print(f"{'部门':<20} {'项目':<35} {'项目总人天':<10} {'季度':<10} {'季度总人天':<10} {'人员':<10} {'人天':<8}")
    print("-" * 140)
    
    for _, row in df.head(30).iterrows():  # 只显示前30行作为预览
        dept = row['订单项目.归属中心']
        project = row['订单项目.立项项目']
        project_total = row['项目总人天']
        quarter = row['季度']
        quarter_total = row['季度总人天']
        person = row['人员']
        person_days = row['人天']
        
        # 处理长文本截断
        if len(str(project)) > 33:
            project = str(project)[:30] + "..."
        
        print(f"{dept:<20} {project:<35} {project_total:<10} {quarter:<10} {quarter_total:<10} {person:<10} {person_days:<8}")
    
    if len(df) > 30:
        print(f"... (还有 {len(df) - 30} 行数据)")
    
    print("-" * 140)
    print(f"总计: {len(df)} 行数据")

def generate_statistics(df):
    """生成统计信息"""
    print(f"\n📈 最终报告统计信息:")
    
    # 统计唯一项目数
    unique_projects = df[['订单项目.归属中心', '订单项目.立项项目']].drop_duplicates()
    project_count = len(unique_projects)
    
    # 统计人员记录数
    person_count = len(df)
    
    # 统计季度
    quarters = df['季度'].unique()
    quarters = [q for q in quarters if q != '' and pd.notna(q)]
    
    # 统计部门
    departments = df['订单项目.归属中心'].unique()
    dept_count = len(departments)
    
    print(f"   唯一项目数: {project_count}")
    print(f"   人员记录数: {person_count}")
    print(f"   涉及部门: {dept_count}个")
    print(f"   涉及季度: {sorted(quarters)}")
    
    # 部门项目统计
    print(f"\n🏢 各部门项目数:")
    dept_project_counts = unique_projects['订单项目.归属中心'].value_counts()
    for dept, count in dept_project_counts.items():
        print(f"   {dept}: {count}个项目")
    
    # 项目总人天统计
    print(f"\n📊 项目总人天TOP10:")
    project_totals = df[['订单项目.立项项目', '项目总人天']].drop_duplicates()
    project_totals['项目总人天_数值'] = project_totals['项目总人天'].astype(float)
    top_projects = project_totals.nlargest(10, '项目总人天_数值')
    
    for _, row in top_projects.iterrows():
        project = row['订单项目.立项项目']
        total = row['项目总人天']
        if len(project) > 40:
            project = project[:37] + "..."
        print(f"   {project:<40} {total}天")

def validate_data(df):
    """验证数据完整性"""
    print(f"\n🔍 数据验证:")
    
    # 检查空值
    empty_cells = 0
    for col in df.columns:
        empty_count = (df[col] == '').sum() + df[col].isna().sum()
        if empty_count > 0:
            print(f"   ⚠️  {col}: {empty_count}个空值")
            empty_cells += empty_count
    
    if empty_cells == 0:
        print("   ✅ 所有单元格都已填入数值")
    else:
        print(f"   ❌ 总计{empty_cells}个空单元格")
    
    # 检查数据类型
    print(f"\n📋 数据格式检查:")
    print(f"   项目总人天格式: {'✅ 正确' if all('.' in str(x) for x in df['项目总人天']) else '❌ 错误'}")
    print(f"   季度总人天格式: {'✅ 正确' if all('.' in str(x) for x in df['季度总人天'] if x != '') else '❌ 错误'}")
    print(f"   人天格式: {'✅ 正确' if all('.' in str(x) for x in df['人天'] if x != '') else '❌ 错误'}")

def main():
    """主函数"""
    input_file = '2025年1-6.csv'  # 原始周报CSV文件
    output_file = '最终优化格式季度工时统计报告.csv'

    try:
        # 加载原始数据
        print("正在加载原始周报数据...")
        raw_df = load_raw_data(input_file)

        if raw_df is None:
            print("❌ 无法加载数据文件")
            return

        print(f"成功加载 {len(raw_df)} 行原始数据")
        print(f"原始列名: {list(raw_df.columns)}")

        # 处理为季度格式
        quarterly_df = process_raw_data_to_quarterly(raw_df)
        print(f"处理后得到 {len(quarterly_df)} 行季度数据")

        # 生成最终优化报告
        final_df = generate_final_optimized_report(quarterly_df)

        # 保存报告
        save_final_report(final_df, output_file)

        # 验证数据
        validate_data(final_df)

        # 生成统计信息
        generate_statistics(final_df)

        print("\n" + "=" * 80)
        print("✅ 最终优化格式季度报告生成完成！")
        print("=" * 80)
        print(f"📁 输出文件: {output_file}")
        print("💡 提示: 所有单元格都已填入具体数值，便于Excel手动合并")

    except Exception as e:
        print(f"❌ 生成报告过程中出现错误: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
周报数据分析脚本
分析每个部门每个项目的用时和每个人在每个项目的用时
"""

import pandas as pd
import numpy as np
from collections import defaultdict
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

def analyze_data(df):
    """分析数据"""
    print("=" * 80)
    print("周报数据分析报告")
    print("=" * 80)
    
    # 基本信息
    print(f"\n📊 基本信息:")
    print(f"   总记录数: {len(df)}")
    print(f"   总列数: {len(df.columns)}")
    print(f"   数据时间范围: 第{df['周次'].min()}周 - 第{df['周次'].max()}周")
    
    # 列名信息
    print(f"\n📋 数据字段:")
    for i, col in enumerate(df.columns, 1):
        print(f"   {i}. {col}")
    
    # 数据类型
    print(f"\n🔍 数据类型:")
    for col, dtype in df.dtypes.items():
        print(f"   {col}: {dtype}")
    
    # 空值检查
    print(f"\n❌ 空值统计:")
    null_counts = df.isnull().sum()
    for col, count in null_counts.items():
        if count > 0:
            print(f"   {col}: {count} ({count/len(df)*100:.1f}%)")
    
    if null_counts.sum() == 0:
        print("   ✅ 没有发现空值")
    
    # 人员统计
    print(f"\n👥 人员统计:")
    unique_people = df['周报人'].unique()
    print(f"   总人数: {len(unique_people)}")
    print(f"   人员列表: {', '.join(unique_people)}")
    
    # 部门统计
    print(f"\n🏢 部门统计:")
    dept_counts = df['订单项目.归属中心'].value_counts()
    print(f"   总部门数: {len(dept_counts)}")
    for dept, count in dept_counts.items():
        print(f"   {dept}: {count} 条记录")
    
    # 项目统计
    print(f"\n📁 项目统计:")
    project_counts = df['订单项目.立项项目'].value_counts()
    print(f"   总项目数: {len(project_counts)}")
    print(f"   前10个项目:")
    for project, count in project_counts.head(10).items():
        print(f"   {project}: {count} 条记录")
    
    # 工时统计
    print(f"\n⏰ 工时统计:")
    time_col = '订单项目.本周投入天数（最低半天）'
    total_time = df[time_col].sum()
    avg_time = df[time_col].mean()
    print(f"   总工时: {total_time:.1f} 天")
    print(f"   平均工时: {avg_time:.2f} 天/记录")
    print(f"   最大工时: {df[time_col].max():.1f} 天")
    print(f"   最小工时: {df[time_col].min():.1f} 天")
    
    return df

def analyze_department_project_time(df):
    """分析每个部门每个项目的用时"""
    print("\n" + "=" * 80)
    print("📊 部门项目用时分析")
    print("=" * 80)

    # 过滤掉空值
    df_clean = df.dropna(subset=['订单项目.归属中心', '订单项目.立项项目', '订单项目.本周投入天数（最低半天）'])

    # 按部门和项目分组统计
    dept_project_time = df_clean.groupby(['订单项目.归属中心', '订单项目.立项项目'])['订单项目.本周投入天数（最低半天）'].agg(['sum', 'count', 'mean']).round(2)
    dept_project_time.columns = ['总工时', '记录数', '平均工时']

    print("\n🏢 各部门各项目用时统计:")
    for dept in df_clean['订单项目.归属中心'].unique():
        dept_data = dept_project_time.loc[dept]
        dept_total = dept_data['总工时'].sum()
        print(f"\n【{dept}】 (总计: {dept_total:.1f}天)")

        # 按总工时排序
        dept_data_sorted = dept_data.sort_values('总工时', ascending=False)
        for project, row in dept_data_sorted.iterrows():
            print(f"  ├─ {project}")
            print(f"  │   总工时: {row['总工时']:.1f}天 | 记录数: {row['记录数']} | 平均: {row['平均工时']:.2f}天/记录")
    
    # 部门总工时排名
    print(f"\n🏆 部门总工时排名:")
    dept_total_time = df.groupby('订单项目.归属中心')['订单项目.本周投入天数（最低半天）'].sum().sort_values(ascending=False)
    for i, (dept, total_time) in enumerate(dept_total_time.items(), 1):
        print(f"   {i}. {dept}: {total_time:.1f}天")
    
    return dept_project_time

def analyze_person_project_time(df):
    """分析每个人在每个项目的用时"""
    print("\n" + "=" * 80)
    print("👤 个人项目用时分析")
    print("=" * 80)

    # 过滤掉空值
    df_clean = df.dropna(subset=['周报人', '订单项目.立项项目', '订单项目.本周投入天数（最低半天）'])

    # 按人员和项目分组统计
    person_project_time = df_clean.groupby(['周报人', '订单项目.立项项目'])['订单项目.本周投入天数（最低半天）'].agg(['sum', 'count', 'mean']).round(2)
    person_project_time.columns = ['总工时', '记录数', '平均工时']

    print("\n👥 各人员各项目用时统计:")
    for person in df_clean['周报人'].unique():
        person_data = person_project_time.loc[person]
        person_total = person_data['总工时'].sum()
        project_count = len(person_data)
        print(f"\n【{person}】 (总计: {person_total:.1f}天, {project_count}个项目)")

        # 按总工时排序
        person_data_sorted = person_data.sort_values('总工时', ascending=False)
        for project, row in person_data_sorted.iterrows():
            print(f"  ├─ {project}")
            print(f"  │   总工时: {row['总工时']:.1f}天 | 记录数: {row['记录数']} | 平均: {row['平均工时']:.2f}天/记录")
    
    # 个人总工时排名
    print(f"\n🏆 个人总工时排名:")
    person_total_time = df.groupby('周报人')['订单项目.本周投入天数（最低半天）'].sum().sort_values(ascending=False)
    for i, (person, total_time) in enumerate(person_total_time.items(), 1):
        print(f"   {i}. {person}: {total_time:.1f}天")
    
    return person_project_time

def analyze_weekly_trends(df):
    """分析周度趋势"""
    print("\n" + "=" * 80)
    print("📈 周度趋势分析")
    print("=" * 80)
    
    # 按周次统计
    weekly_stats = df.groupby('周次').agg({
        '订单项目.本周投入天数（最低半天）': ['sum', 'count', 'mean'],
        '周报人': 'nunique',
        '订单项目.立项项目': 'nunique'
    }).round(2)
    
    weekly_stats.columns = ['总工时', '记录数', '平均工时', '参与人数', '项目数']
    
    print("\n📅 各周统计:")
    for week, row in weekly_stats.iterrows():
        print(f"第{week}周: 总工时{row['总工时']:.1f}天 | 记录{row['记录数']}条 | 人员{row['参与人数']}人 | 项目{row['项目数']}个")
    
    return weekly_stats

def generate_summary_report(df, dept_project_time, person_project_time):
    """生成汇总报告"""
    print("\n" + "=" * 80)
    print("📋 汇总报告")
    print("=" * 80)
    
    # 关键指标
    total_time = df['订单项目.本周投入天数（最低半天）'].sum()
    total_people = df['周报人'].nunique()
    total_projects = df['订单项目.立项项目'].nunique()
    total_depts = df['订单项目.归属中心'].nunique()
    
    print(f"\n🎯 关键指标:")
    print(f"   总工时: {total_time:.1f} 天")
    print(f"   参与人员: {total_people} 人")
    print(f"   涉及项目: {total_projects} 个")
    print(f"   涉及部门: {total_depts} 个")
    print(f"   人均工时: {total_time/total_people:.1f} 天/人")
    
    # 最活跃的项目
    most_active_project = df['订单项目.立项项目'].value_counts().index[0]
    most_active_project_time = df[df['订单项目.立项项目'] == most_active_project]['订单项目.本周投入天数（最低半天）'].sum()
    
    print(f"\n🔥 最活跃项目:")
    print(f"   项目名称: {most_active_project}")
    print(f"   总工时: {most_active_project_time:.1f} 天")
    print(f"   记录数: {df['订单项目.立项项目'].value_counts().iloc[0]} 条")
    
    # 最忙碌的人员
    busiest_person = df.groupby('周报人')['订单项目.本周投入天数（最低半天）'].sum().idxmax()
    busiest_person_time = df.groupby('周报人')['订单项目.本周投入天数（最低半天）'].sum().max()
    
    print(f"\n💪 最忙碌人员:")
    print(f"   人员姓名: {busiest_person}")
    print(f"   总工时: {busiest_person_time:.1f} 天")
    
    # 最忙碌的部门
    busiest_dept = df.groupby('订单项目.归属中心')['订单项目.本周投入天数（最低半天）'].sum().idxmax()
    busiest_dept_time = df.groupby('订单项目.归属中心')['订单项目.本周投入天数（最低半天）'].sum().max()
    
    print(f"\n🏢 最忙碌部门:")
    print(f"   部门名称: {busiest_dept}")
    print(f"   总工时: {busiest_dept_time:.1f} 天")

def main():
    """主函数"""
    file_path = '2025年1-6.csv'
    
    try:
        # 加载数据
        print("正在加载数据...")
        df = load_data(file_path)
        
        # 基本分析
        df = analyze_data(df)
        
        # 部门项目用时分析
        dept_project_time = analyze_department_project_time(df)
        
        # 个人项目用时分析
        person_project_time = analyze_person_project_time(df)
        
        # 周度趋势分析
        weekly_stats = analyze_weekly_trends(df)
        
        # 生成汇总报告
        generate_summary_report(df, dept_project_time, person_project_time)
        
        print("\n" + "=" * 80)
        print("✅ 分析完成！")
        print("=" * 80)
        
    except Exception as e:
        print(f"❌ 分析过程中出现错误: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()

#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import re
from collections import defaultdict
import json

def read_csv_with_encoding(file_path):
    """尝试不同编码读取CSV文件"""
    encodings = ['utf-8-sig', 'utf-8', 'gbk', 'gb2312', 'gb18030']
    
    for encoding in encodings:
        try:
            print(f"尝试使用编码: {encoding}")
            df = pd.read_csv(file_path, encoding=encoding)
            print(f"成功使用编码 {encoding} 读取文件")
            print(f"数据形状: {df.shape}")
            print(f"列名: {list(df.columns)}")
            return df, encoding
        except Exception as e:
            print(f"编码 {encoding} 失败: {e}")
            continue
    
    raise Exception("无法使用任何编码读取文件")

def analyze_work_content(content):
    """分析工作内容，分类为开发、维护、调机等"""
    if pd.isna(content) or content == '':
        return {'type': 'unknown', 'details': []}
    
    content = str(content).lower()
    
    # 调机相关关键词
    tuning_keywords = [
        '调机', '设备调试', '机器调试', '系统调机', '调试设备', 
        '设备调整', '机台调试', '调试机器', '设备维护', '机器维护',
        '设备安装', '机器安装', '设备配置', '机器配置'
    ]
    
    # 开发相关关键词
    development_keywords = [
        '开发', '实现', '编写', '设计', '创建', '构建', '新增', '添加',
        '功能', '需求', '特性', '模块', '接口', 'api', '算法', '逻辑',
        '界面', 'ui', '前端', '后端', '数据库', '系统', '平台'
    ]
    
    # Bug修复关键词
    bug_keywords = [
        'bug', '修复', '修改', '解决', '问题', '错误', '异常', '故障',
        '优化', '改进', '完善', '调整', '更新', '升级'
    ]
    
    # 检查是否为调机工作
    for keyword in tuning_keywords:
        if keyword in content:
            return {'type': 'tuning', 'details': [content]}
    
    # 检查是否为开发工作
    dev_found = any(keyword in content for keyword in development_keywords)
    bug_found = any(keyword in content for keyword in bug_keywords)
    
    if dev_found and bug_found:
        return {'type': 'mixed', 'details': [content]}
    elif dev_found:
        return {'type': 'development', 'details': [content]}
    elif bug_found:
        return {'type': 'maintenance', 'details': [content]}
    else:
        return {'type': 'other', 'details': [content]}

def extract_requirements_and_bugs(content):
    """提取具体需求和bug修复内容"""
    if pd.isna(content) or content == '':
        return {'requirements': [], 'bugs': []}
    
    content = str(content)
    
    # 分割内容（按数字编号或分号分割）
    items = re.split(r'[；;]\s*|\d+[\.、]\s*', content)
    items = [item.strip() for item in items if item.strip()]
    
    requirements = []
    bugs = []
    
    for item in items:
        item_lower = item.lower()
        
        # Bug相关关键词
        if any(keyword in item_lower for keyword in ['bug', '修复', '修改', '解决', '问题', '错误', '异常', '故障']):
            bugs.append(item)
        # 需求/功能相关关键词
        elif any(keyword in item_lower for keyword in ['实现', '开发', '新增', '添加', '功能', '需求', '特性']):
            requirements.append(item)
        else:
            # 默认归类为需求
            requirements.append(item)
    
    return {'requirements': requirements, 'bugs': bugs}

def analyze_projects(df):
    """按项目分组分析工作内容"""

    # 按项目分组
    project_analysis = {}

    for project in df['订单项目.立项项目'].unique():
        if pd.isna(project):
            continue

        project_data = df[df['订单项目.立项项目'] == project]

        # 统计基本信息
        total_days = project_data['订单项目.本周投入天数（最低半天）'].sum()
        people_count = project_data['周报人'].nunique()

        # 分析工作内容
        all_requirements = []
        all_bugs = []
        tuning_work = []
        development_work = []
        maintenance_work = []
        other_work = []

        for _, row in project_data.iterrows():
            content = row['订单项目.本周进度及问题反馈']
            person = row['周报人']
            days = row['订单项目.本周投入天数（最低半天）']

            # 分析工作类型
            work_type = analyze_work_content(content)

            # 提取需求和bug
            req_bug = extract_requirements_and_bugs(content)

            work_item = {
                'person': person,
                'days': days,
                'content': content,
                'requirements': req_bug['requirements'],
                'bugs': req_bug['bugs']
            }

            if work_type['type'] == 'tuning':
                tuning_work.append(work_item)
            elif work_type['type'] == 'development':
                development_work.append(work_item)
                all_requirements.extend(req_bug['requirements'])
            elif work_type['type'] == 'maintenance':
                maintenance_work.append(work_item)
                all_bugs.extend(req_bug['bugs'])
            elif work_type['type'] == 'mixed':
                development_work.append(work_item)
                all_requirements.extend(req_bug['requirements'])
                all_bugs.extend(req_bug['bugs'])
            else:
                other_work.append(work_item)

        # 计算筛选后的工作量（排除调机）
        filtered_days = sum([item['days'] for item in development_work + maintenance_work + other_work])
        tuning_days = sum([item['days'] for item in tuning_work])

        project_analysis[project] = {
            'total_days': total_days,
            'filtered_days': filtered_days,
            'tuning_days': tuning_days,
            'people_count': people_count,
            'requirements': list(set(all_requirements)),  # 去重
            'bugs': list(set(all_bugs)),  # 去重
            'development_work': development_work,
            'maintenance_work': maintenance_work,
            'tuning_work': tuning_work,
            'other_work': other_work
        }

    return project_analysis

def generate_report(project_analysis):
    """生成分析报告"""

    print("\n" + "="*80)
    print("项目数据语义分析报告")
    print("="*80)

    # 总体统计
    total_projects = len(project_analysis)
    total_original_days = sum([p['total_days'] for p in project_analysis.values()])
    total_filtered_days = sum([p['filtered_days'] for p in project_analysis.values()])
    total_tuning_days = sum([p['tuning_days'] for p in project_analysis.values()])

    print(f"\n【总体统计】")
    print(f"项目总数: {total_projects}")
    print(f"原始总工作量: {total_original_days:.1f} 人天")
    print(f"筛选后工作量: {total_filtered_days:.1f} 人天 (排除调机工作)")
    print(f"调机工作量: {total_tuning_days:.1f} 人天")
    print(f"调机工作占比: {(total_tuning_days/total_original_days*100):.1f}%")

    # 按项目详细分析
    print(f"\n【按项目详细分析】")

    for project, analysis in project_analysis.items():
        if analysis['filtered_days'] == 0:  # 跳过纯调机项目
            continue

        print(f"\n项目: {project}")
        print(f"  参与人数: {analysis['people_count']} 人")
        print(f"  总工作量: {analysis['total_days']:.1f} 人天")
        print(f"  核心开发工作量: {analysis['filtered_days']:.1f} 人天")

        if analysis['tuning_days'] > 0:
            print(f"  调机工作量: {analysis['tuning_days']:.1f} 人天")

        # 具体需求列表
        if analysis['requirements']:
            print(f"  【具体需求列表】({len(analysis['requirements'])}项):")
            for i, req in enumerate(analysis['requirements'][:10], 1):  # 最多显示10项
                print(f"    {i}. {req}")
            if len(analysis['requirements']) > 10:
                print(f"    ... 还有{len(analysis['requirements'])-10}项需求")

        # Bug修复列表
        if analysis['bugs']:
            print(f"  【Bug修复列表】({len(analysis['bugs'])}项):")
            for i, bug in enumerate(analysis['bugs'][:10], 1):  # 最多显示10项
                print(f"    {i}. {bug}")
            if len(analysis['bugs']) > 10:
                print(f"    ... 还有{len(analysis['bugs'])-10}项Bug修复")

    # 工作性质分布统计
    print(f"\n【工作性质分布统计】")
    dev_days = sum([len(p['development_work']) * (sum([w['days'] for w in p['development_work']])/max(len(p['development_work']),1)) for p in project_analysis.values()])
    maint_days = sum([len(p['maintenance_work']) * (sum([w['days'] for w in p['maintenance_work']])/max(len(p['maintenance_work']),1)) for p in project_analysis.values()])

    print(f"开发性工作: {dev_days:.1f} 人天")
    print(f"维护性工作: {maint_days:.1f} 人天")
    print(f"调机工作: {total_tuning_days:.1f} 人天")

def main():
    file_path = "2025年1-6.csv"

    try:
        # 读取CSV文件
        df, encoding = read_csv_with_encoding(file_path)

        print(f"文件基本信息:")
        print(f"总行数: {len(df)}")
        print(f"列数: {len(df.columns)}")
        print(f"使用编码: {encoding}")

        # 进行项目分析
        project_analysis = analyze_projects(df)

        # 生成报告
        generate_report(project_analysis)

        # 保存详细分析结果
        with open('project_analysis_result.json', 'w', encoding='utf-8') as f:
            # 转换为可序列化的格式
            serializable_analysis = {}
            for project, analysis in project_analysis.items():
                serializable_analysis[project] = {
                    'total_days': float(analysis['total_days']),
                    'filtered_days': float(analysis['filtered_days']),
                    'tuning_days': float(analysis['tuning_days']),
                    'people_count': int(analysis['people_count']),
                    'requirements': analysis['requirements'],
                    'bugs': analysis['bugs'],
                    'development_work_count': len(analysis['development_work']),
                    'maintenance_work_count': len(analysis['maintenance_work']),
                    'tuning_work_count': len(analysis['tuning_work'])
                }
            json.dump(serializable_analysis, f, ensure_ascii=False, indent=2)

        print(f"\n详细分析结果已保存到 project_analysis_result.json")

        return df, project_analysis

    except Exception as e:
        print(f"处理文件时出错: {e}")
        return None, None

if __name__ == "__main__":
    df = main()

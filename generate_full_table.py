#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import pandas as pd
from collections import defaultdict

def generate_full_table():
    """生成完整的详细工作内容对照表"""
    
    # 读取分析结果
    with open('detailed_record_analysis.json', 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # 读取原始CSV获取部门信息
    df = pd.read_csv("2025年1-6.csv", encoding='gbk')
    
    # 创建索引到部门的映射
    index_to_dept = {}
    for idx, row in df.iterrows():
        dept = row['订单项目.归属中心']
        index_to_dept[idx + 1] = dept
    
    # 按部门分组收集数据
    departments = {
        'T1': {'requirements': [], 'bugs': []},
        'T1电子元件': {'requirements': [], 'bugs': []},
        'T2': {'requirements': [], 'bugs': []},
        'T3': {'requirements': [], 'bugs': []},
        'T4': {'requirements': [], 'bugs': []},
        '软件': {'requirements': [], 'bugs': []}
    }
    
    for record in data:
        index = record['index']
        dept = index_to_dept.get(index, '未知部门')
        
        if pd.isna(dept):
            dept = '未知部门'
        else:
            dept = str(dept)
        
        # 映射部门名称
        if 'T1' in dept and '电子元件' in dept:
            dept_key = 'T1电子元件'
        elif 'T1' in dept:
            dept_key = 'T1'
        elif 'T2' in dept:
            dept_key = 'T2'
        elif 'T3' in dept:
            dept_key = 'T3'
        elif 'T4' in dept:
            dept_key = 'T4'
        elif '软件' in dept:
            dept_key = '软件'
        else:
            continue
        
        content = record['content']
        work_type = record['analysis']['type']
        
        # 根据分析结果分类
        if work_type in ['software_development', 'system_integration']:
            departments[dept_key]['requirements'].append(content)
        elif work_type in ['software_maintenance']:
            departments[dept_key]['bugs'].append(content)
        else:
            # 对于其他类型，根据内容判断
            if any(keyword in content.lower() for keyword in ['开发', '实现', '创建', '新增', '添加', '功能', '设计']):
                departments[dept_key]['requirements'].append(content)
            elif any(keyword in content.lower() for keyword in ['修复', '解决', '问题', 'bug', '错误', '优化']):
                departments[dept_key]['bugs'].append(content)
            else:
                departments[dept_key]['requirements'].append(content)
    
    return departments

def write_full_table_to_file(departments):
    """将完整表格写入文件"""
    
    # 找出最大行数
    max_rows = 0
    active_depts = []
    
    for dept, data in departments.items():
        if data['requirements'] or data['bugs']:
            active_depts.append(dept)
            max_rows = max(max_rows, len(data['requirements']), len(data['bugs']))
    
    # 生成表格内容
    with open('完整工作内容对照表.md', 'w', encoding='utf-8') as f:
        f.write("# 完整工作内容对照表\n\n")
        
        # 写入统计信息
        f.write("## 统计概览\n")
        f.write("| 部门 | 需求数量 | Bug修复数量 | 总工作项 |\n")
        f.write("|------|----------|-------------|----------|\n")
        
        total_reqs = 0
        total_bugs = 0
        
        for dept in active_depts:
            req_count = len(departments[dept]['requirements'])
            bug_count = len(departments[dept]['bugs'])
            total_count = req_count + bug_count
            total_reqs += req_count
            total_bugs += bug_count
            f.write(f"| {dept} | {req_count}项 | {bug_count}项 | {total_count}项 |\n")
        
        f.write(f"| **总计** | **{total_reqs}项** | **{total_bugs}项** | **{total_reqs + total_bugs}项** |\n\n")
        
        # 写入表头
        f.write("## 详细工作内容对照表\n\n")
        header_parts = []
        for dept in active_depts:
            header_parts.append(f"{dept}需求")
            header_parts.append(f"{dept}bug")
        
        f.write("| " + " | ".join(header_parts) + " |\n")
        f.write("|" + "|".join(["--------"] * len(header_parts)) + "|\n")
        
        # 写入数据行
        for row_idx in range(min(max_rows, 50)):  # 限制最多50行
            row_data = []
            for dept in active_depts:
                # 需求列
                if row_idx < len(departments[dept]['requirements']):
                    req_content = departments[dept]['requirements'][row_idx]
                    # 清理内容，去除换行符和多余空格
                    req_content = req_content.replace('\n', ' ').replace('\r', ' ').strip()
                    # 截断过长的内容
                    if len(req_content) > 80:
                        req_content = req_content[:77] + "..."
                    row_data.append(req_content)
                else:
                    row_data.append("")
                
                # Bug列
                if row_idx < len(departments[dept]['bugs']):
                    bug_content = departments[dept]['bugs'][row_idx]
                    # 清理内容
                    bug_content = bug_content.replace('\n', ' ').replace('\r', ' ').strip()
                    # 截断过长的内容
                    if len(bug_content) > 80:
                        bug_content = bug_content[:77] + "..."
                    row_data.append(bug_content)
                else:
                    row_data.append("")
            
            f.write("| " + " | ".join(row_data) + " |\n")
        
        # 如果还有更多数据，添加说明
        if max_rows > 50:
            f.write(f"\n*注：表格仅显示前50行数据，完整数据共{max_rows}行*\n")

def main():
    print("正在生成完整的工作内容对照表...")
    departments = generate_full_table()
    write_full_table_to_file(departments)
    print("完整表格已生成到文件：完整工作内容对照表.md")
    
    # 显示统计信息
    print("\n统计信息:")
    for dept, data in departments.items():
        if data['requirements'] or data['bugs']:
            print(f"{dept}: 需求{len(data['requirements'])}项, Bug修复{len(data['bugs'])}项")

if __name__ == "__main__":
    main()

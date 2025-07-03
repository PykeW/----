#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import pandas as pd
from collections import defaultdict
import re

def extract_detailed_requirements_bugs():
    """提取详细的需求和Bug修复内容"""
    
    # 读取分析结果
    with open('detailed_record_analysis.json', 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # 按部门分组
    departments = {
        'T1': [],
        'T1电子元件': [],
        'T2': [],
        'T3': [],
        'T4': [],
        '费用中心-软件': []
    }
    
    # 读取原始CSV获取部门信息
    df = pd.read_csv("2025年1-6.csv", encoding='gbk')
    
    # 创建索引到部门的映射
    index_to_dept = {}
    for idx, row in df.iterrows():
        dept = row['订单项目.归属中心']
        index_to_dept[idx + 1] = dept  # 索引从1开始
    
    # 分类收集需求和Bug
    for record in data:
        index = record['index']
        dept = index_to_dept.get(index, '未知部门')

        # 处理NaN值
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
            dept_key = '费用中心-软件'
        else:
            dept_key = '其他'
        
        if dept_key not in departments:
            departments[dept_key] = []
        
        # 分析工作内容
        content = record['content']
        work_type = record['analysis']['type']
        
        # 提取具体的需求和Bug
        items = extract_work_items(content)
        
        for item in items:
            if is_requirement(item):
                departments[dept_key].append({
                    'type': 'requirement',
                    'content': item,
                    'person': record['person'],
                    'days': record['days'],
                    'work_type': work_type
                })
            elif is_bug_fix(item):
                departments[dept_key].append({
                    'type': 'bug',
                    'content': item,
                    'person': record['person'],
                    'days': record['days'],
                    'work_type': work_type
                })
    
    return departments

def extract_work_items(content):
    """从工作内容中提取具体的工作项"""
    if not content or pd.isna(content):
        return []
    
    content = str(content)
    
    # 按数字编号、分号、换行符分割
    items = re.split(r'[；;]\s*|\d+[\.、]\s*|\n', content)
    items = [item.strip() for item in items if item.strip() and len(item.strip()) > 5]
    
    # 如果没有分割出多个项目，返回原内容
    if len(items) <= 1:
        return [content.strip()]
    
    return items

def is_requirement(item):
    """判断是否为需求"""
    item_lower = item.lower()
    
    requirement_indicators = [
        '开发', '实现', '创建', '新增', '添加', '构建', '设计', '完成',
        '制作', '生成', '建立', '搭建', '编写', '功能', '模块', '组件',
        '接口', '系统', '平台', '界面', '算法', '流程'
    ]
    
    return any(indicator in item_lower for indicator in requirement_indicators)

def is_bug_fix(item):
    """判断是否为Bug修复"""
    item_lower = item.lower()
    
    bug_indicators = [
        'bug', '修复', '修改', '解决', '问题', '错误', '异常', '故障',
        '优化', '改进', '完善', '调整', '更新', '升级'
    ]
    
    return any(indicator in item_lower for indicator in bug_indicators)

def format_output(departments):
    """格式化输出"""
    
    print("按部门分类的需求和Bug修复工作详细列表")
    print("=" * 100)
    
    for dept, items in departments.items():
        if not items:
            continue
            
        requirements = [item for item in items if item['type'] == 'requirement']
        bugs = [item for item in items if item['type'] == 'bug']
        
        if not requirements and not bugs:
            continue
        
        print(f"\n【{dept}】")
        print(f"需求数量: {len(requirements)} 项")
        print(f"Bug修复数量: {len(bugs)} 项")
        print("-" * 80)
        
        # 输出需求
        if requirements:
            print("需求列表:")
            for i, req in enumerate(requirements[:20], 1):  # 限制显示数量
                print(f"{i:2d}. {req['content']}")
            if len(requirements) > 20:
                print(f"    ... 还有{len(requirements)-20}项需求")
        
        print()
        
        # 输出Bug修复
        if bugs:
            print("Bug修复列表:")
            for i, bug in enumerate(bugs[:20], 1):  # 限制显示数量
                print(f"{i:2d}. {bug['content']}")
            if len(bugs) > 20:
                print(f"    ... 还有{len(bugs)-20}项Bug修复")
        
        print("\n" + "=" * 100)

def generate_table_format(departments):
    """生成表格格式输出"""
    
    print("\n\n表格格式输出（按您要求的格式）:")
    print("=" * 120)
    
    # 找出有数据的部门
    active_depts = []
    for dept, items in departments.items():
        if items:
            requirements = [item for item in items if item['type'] == 'requirement']
            bugs = [item for item in items if item['type'] == 'bug']
            if requirements or bugs:
                active_depts.append((dept, requirements, bugs))
    
    if not active_depts:
        print("没有找到有效的需求和Bug数据")
        return
    
    # 生成表头
    header_parts = []
    for dept, reqs, bugs in active_depts:
        dept_short = dept.replace('费用中心-', '').replace('电子元件', '元件')
        header_parts.append(f"{dept_short}需求")
        header_parts.append(f"{dept_short}bug")
    
    print("\t".join(header_parts))
    
    # 找出最大行数
    max_rows = 0
    for dept, reqs, bugs in active_depts:
        max_rows = max(max_rows, len(reqs), len(bugs))
    
    # 生成数据行
    for row_idx in range(min(max_rows, 10)):  # 限制最多10行
        row_data = []
        for dept, reqs, bugs in active_depts:
            # 需求列
            if row_idx < len(reqs):
                req_content = reqs[row_idx]['content']
                # 截断过长的内容
                if len(req_content) > 50:
                    req_content = req_content[:47] + "..."
                row_data.append(req_content)
            else:
                row_data.append("")
            
            # Bug列
            if row_idx < len(bugs):
                bug_content = bugs[row_idx]['content']
                # 截断过长的内容
                if len(bug_content) > 50:
                    bug_content = bug_content[:47] + "..."
                row_data.append(bug_content)
            else:
                row_data.append("")
        
        print("\t".join(row_data))

def main():
    departments = extract_detailed_requirements_bugs()
    format_output(departments)
    generate_table_format(departments)

if __name__ == "__main__":
    main()

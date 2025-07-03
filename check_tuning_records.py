#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json

def check_tuning_records():
    """检查设备调机记录"""
    with open('detailed_record_analysis.json', 'r', encoding='utf-8') as f:
        data = json.load(f)

    # 查找设备调机的记录
    tuning_records = [record for record in data if record['analysis']['type'] == 'equipment_tuning']
    print(f'设备调机记录数: {len(tuning_records)}')
    print()
    
    for i, record in enumerate(tuning_records, 1):
        print(f'记录 {i}:')
        print(f'  人员: {record["person"]}')
        print(f'  项目: {record["project"]}')
        print(f'  工作量: {record["days"]} 人天')
        print(f'  工作内容: {record["content"]}')
        print(f'  分析理由: {record["analysis"]["analysis_reason"]}')
        print()

    # 查看一些软件开发记录
    dev_records = [record for record in data if record['analysis']['type'] == 'software_development'][:5]
    print(f'\n软件开发记录示例 (前5条):')
    for i, record in enumerate(dev_records, 1):
        print(f'记录 {i}:')
        print(f'  人员: {record["person"]}')
        print(f'  工作内容: {record["content"][:50]}...')
        print(f'  子类型: {record["analysis"]["subtype"]}')
        print(f'  技术领域: {record["analysis"]["technical_area"]}')
        print()

if __name__ == "__main__":
    check_tuning_records()

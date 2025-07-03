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

def analyze_work_content_semantic(content, person, project, days):
    """基于语义理解的深度工作内容分析"""
    if pd.isna(content) or content == '':
        return {
            'type': 'unknown',
            'subtype': 'empty_content',
            'technical_area': 'unknown',
            'work_nature': 'unknown',
            'analysis_reason': '工作内容为空',
            'confidence': 0.0
        }

    content_str = str(content)
    content_lower = content_str.lower()

    # 分析结果结构
    analysis_result = {
        'type': 'unknown',
        'subtype': 'unknown',
        'technical_area': 'unknown',
        'work_nature': 'unknown',
        'analysis_reason': '',
        'confidence': 0.0
    }

    # 1. 设备调机和硬件调试分析
    if _is_equipment_tuning(content_str):
        analysis_result.update({
            'type': 'equipment_tuning',
            'subtype': _get_tuning_subtype(content_str),
            'technical_area': 'hardware_equipment',
            'work_nature': 'equipment_operation',
            'analysis_reason': '涉及设备调试、机器调整、硬件配置等非软件开发工作',
            'confidence': 0.9
        })
        return analysis_result

    # 2. 软件开发工作分析
    if _is_software_development(content_str):
        subtype, tech_area, reason = _analyze_development_details(content_str)
        analysis_result.update({
            'type': 'software_development',
            'subtype': subtype,
            'technical_area': tech_area,
            'work_nature': 'development',
            'analysis_reason': reason,
            'confidence': 0.85
        })
        return analysis_result

    # 3. Bug修复和维护工作分析
    if _is_maintenance_work(content_str):
        subtype, tech_area, reason = _analyze_maintenance_details(content_str)
        analysis_result.update({
            'type': 'software_maintenance',
            'subtype': subtype,
            'technical_area': tech_area,
            'work_nature': 'maintenance',
            'analysis_reason': reason,
            'confidence': 0.8
        })
        return analysis_result

    # 4. 系统集成和配置工作
    if _is_system_integration(content_str):
        subtype, tech_area, reason = _analyze_integration_details(content_str)
        analysis_result.update({
            'type': 'system_integration',
            'subtype': subtype,
            'technical_area': tech_area,
            'work_nature': 'integration',
            'analysis_reason': reason,
            'confidence': 0.75
        })
        return analysis_result

    # 5. 学习和研究工作
    if _is_learning_research(content_str):
        analysis_result.update({
            'type': 'learning_research',
            'subtype': 'knowledge_acquisition',
            'technical_area': _infer_technical_area(content_str),
            'work_nature': 'learning',
            'analysis_reason': '涉及学习、研究、熟悉新技术或系统',
            'confidence': 0.7
        })
        return analysis_result

    # 6. 默认分类 - 基于项目和内容推断
    tech_area = _infer_technical_area(content_str)
    analysis_result.update({
        'type': 'other_work',
        'subtype': 'unclassified',
        'technical_area': tech_area,
        'work_nature': 'other',
        'analysis_reason': f'无法明确分类的工作内容，推断技术领域为{tech_area}',
        'confidence': 0.3
    })

    return analysis_result

def _is_equipment_tuning(content):
    """判断是否为设备调机工作"""
    content_lower = content.lower()

    # 明确的设备调机指标
    tuning_indicators = [
        '调机', '设备调试', '机器调试', '系统调机', '调试设备',
        '设备调整', '机台调试', '调试机器', '设备维护', '机器维护',
        '设备安装', '机器安装', '设备配置', '机器配置', '硬件调试',
        '现场调试', '设备标定', '机器标定', '设备校准', '机器校准'
    ]

    # 设备相关但可能是软件工作的词汇需要更仔细判断
    equipment_contexts = [
        '设备通信', '设备接口', '设备控制', '设备监控', '设备状态',
        '机器控制', '机器监控', '机器状态', 'plc', '传感器', '电机'
    ]

    # 直接匹配明确的调机指标
    for indicator in tuning_indicators:
        if indicator in content_lower:
            return True

    # 对于设备相关词汇，需要结合上下文判断
    equipment_count = sum(1 for term in equipment_contexts if term in content_lower)
    software_indicators = ['开发', '编写', '实现', '代码', '程序', '软件', '界面', 'ui', 'api', '算法']
    software_count = sum(1 for term in software_indicators if term in content_lower)

    # 如果设备词汇多但软件词汇少，可能是调机工作
    if equipment_count >= 2 and software_count == 0:
        return True

    return False

def _get_tuning_subtype(content):
    """获取调机工作的子类型"""
    content_lower = content.lower()

    if any(term in content_lower for term in ['安装', '配置', '部署']):
        return 'installation_configuration'
    elif any(term in content_lower for term in ['标定', '校准', '校正']):
        return 'calibration'
    elif any(term in content_lower for term in ['维护', '保养', '检修']):
        return 'maintenance'
    elif any(term in content_lower for term in ['调试', '调整', '优化']):
        return 'debugging_tuning'
    else:
        return 'general_tuning'

def _is_software_development(content):
    """判断是否为软件开发工作"""
    content_lower = content.lower()

    # 明确的开发指标
    development_indicators = [
        '开发', '实现', '编写', '创建', '构建', '设计', '新增', '添加',
        '完成', '制作', '生成', '建立', '搭建'
    ]

    # 技术开发相关词汇
    technical_indicators = [
        '功能', '模块', '组件', '接口', 'api', '算法', '逻辑', '流程',
        '界面', 'ui', '前端', '后端', '数据库', '系统', '平台', '框架',
        '代码', '程序', '脚本', '方法', '类', '函数'
    ]

    # 检查是否包含开发动词
    has_development_verb = any(verb in content_lower for verb in development_indicators)

    # 检查是否包含技术名词
    has_technical_noun = any(noun in content_lower for noun in technical_indicators)

    # 排除纯粹的学习或了解
    learning_only = any(term in content_lower for term in ['学习', '了解', '熟悉', '研究']) and \
                   not any(term in content_lower for term in development_indicators)

    return has_development_verb and has_technical_noun and not learning_only

def _analyze_development_details(content):
    """分析开发工作的详细信息"""
    content_lower = content.lower()

    # 前端开发
    if any(term in content_lower for term in ['界面', 'ui', '前端', '页面', '组件', 'react', 'vue', '按钮', '表格', '图表']):
        return 'frontend_development', 'frontend', '前端界面开发，包括UI组件、页面交互等'

    # 后端开发
    elif any(term in content_lower for term in ['后端', '服务', 'api', '接口', '数据库', '服务器', '数据处理']):
        return 'backend_development', 'backend', '后端服务开发，包括API接口、数据处理等'

    # 算法开发
    elif any(term in content_lower for term in ['算法', '模型', '识别', '检测', '分析', '计算', '处理', '匹配']):
        return 'algorithm_development', 'algorithm', '算法开发，包括图像处理、模式识别等'

    # 系统开发
    elif any(term in content_lower for term in ['系统', '平台', '框架', '架构', '流程', '逻辑']):
        return 'system_development', 'system', '系统平台开发，包括架构设计、流程实现等'

    # 工具开发
    elif any(term in content_lower for term in ['工具', '脚本', '程序', '软件', '应用']):
        return 'tool_development', 'tools', '工具软件开发，包括脚本程序、应用软件等'

    else:
        return 'general_development', 'general', '通用软件开发工作'

def _is_maintenance_work(content):
    """判断是否为维护工作"""
    content_lower = content.lower()

    # 明确的维护指标
    maintenance_indicators = [
        'bug', '修复', '修改', '解决', '问题', '错误', '异常', '故障',
        '优化', '改进', '完善', '调整', '更新', '升级', '重构'
    ]

    return any(indicator in content_lower for indicator in maintenance_indicators)

def _analyze_maintenance_details(content):
    """分析维护工作的详细信息"""
    content_lower = content.lower()

    # Bug修复
    if any(term in content_lower for term in ['bug', '错误', '异常', '故障', '问题', '修复']):
        return 'bug_fixing', _infer_technical_area(content), 'Bug修复工作，解决软件缺陷和异常'

    # 性能优化
    elif any(term in content_lower for term in ['优化', '性能', '速度', '效率', '内存', '响应']):
        return 'performance_optimization', _infer_technical_area(content), '性能优化工作，提升系统效率和响应速度'

    # 功能改进
    elif any(term in content_lower for term in ['改进', '完善', '增强', '提升']):
        return 'feature_improvement', _infer_technical_area(content), '功能改进工作，完善现有功能'

    # 代码重构
    elif any(term in content_lower for term in ['重构', '重写', '调整', '整理']):
        return 'code_refactoring', _infer_technical_area(content), '代码重构工作，改善代码结构和质量'

    # 版本更新
    elif any(term in content_lower for term in ['更新', '升级', '版本']):
        return 'version_update', _infer_technical_area(content), '版本更新工作，升级软件或依赖'

    else:
        return 'general_maintenance', _infer_technical_area(content), '通用维护工作'

def _is_system_integration(content):
    """判断是否为系统集成工作"""
    content_lower = content.lower()

    integration_indicators = [
        '集成', '对接', '连接', '通信', '协议', '接口', '配置',
        '部署', '安装', '环境', '测试', '验证', '联调'
    ]

    return any(indicator in content_lower for indicator in integration_indicators)

def _analyze_integration_details(content):
    """分析系统集成工作的详细信息"""
    content_lower = content.lower()

    # 第三方集成
    if any(term in content_lower for term in ['第三方', '外部', 'api', '接口', '对接']):
        return 'third_party_integration', 'integration', '第三方系统集成，包括API对接、外部服务集成'

    # 设备集成
    elif any(term in content_lower for term in ['设备', '硬件', 'plc', '传感器', '相机', '控制器']):
        return 'device_integration', 'hardware_integration', '设备硬件集成，包括PLC、传感器等设备对接'

    # 数据库集成
    elif any(term in content_lower for term in ['数据库', '数据', '存储', 'mysql', 'sql']):
        return 'database_integration', 'database', '数据库集成，包括数据存储、查询等'

    # 网络通信
    elif any(term in content_lower for term in ['通信', '网络', 'tcp', 'http', '协议']):
        return 'network_integration', 'network', '网络通信集成，包括通信协议、网络配置'

    else:
        return 'general_integration', 'integration', '通用系统集成工作'

def _is_learning_research(content):
    """判断是否为学习研究工作"""
    content_lower = content.lower()

    learning_indicators = [
        '学习', '了解', '熟悉', '研究', '调研', '分析', '探索',
        '掌握', '理解', '认识', '知识', '技术', '方案'
    ]

    # 必须包含学习词汇，且不包含明确的开发动作
    has_learning = any(indicator in content_lower for indicator in learning_indicators)
    has_development = any(term in content_lower for term in ['开发', '实现', '编写', '创建', '构建'])

    return has_learning and not has_development

def _infer_technical_area(content):
    """推断技术领域"""
    content_lower = content.lower()

    # 前端技术
    if any(term in content_lower for term in ['前端', 'ui', '界面', '页面', 'react', 'vue', '组件']):
        return 'frontend'

    # 后端技术
    elif any(term in content_lower for term in ['后端', '服务', 'api', '数据库', '服务器']):
        return 'backend'

    # 算法/AI
    elif any(term in content_lower for term in ['算法', '模型', 'ai', '识别', '检测', '视觉', '图像']):
        return 'algorithm_ai'

    # 设备控制
    elif any(term in content_lower for term in ['设备', '控制', 'plc', '运动', '电机', '传感器']):
        return 'device_control'

    # 系统架构
    elif any(term in content_lower for term in ['系统', '架构', '平台', '框架']):
        return 'system_architecture'

    # 数据处理
    elif any(term in content_lower for term in ['数据', '统计', '分析', '报表', '导出']):
        return 'data_processing'

    # 测试
    elif any(term in content_lower for term in ['测试', '验证', '检验', '调试']):
        return 'testing'

    else:
        return 'general'

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

def analyze_all_records(df):
    """逐条分析所有工作记录"""

    print("开始逐条分析工作记录...")

    all_analyses = []
    work_type_stats = defaultdict(lambda: {'count': 0, 'total_days': 0.0, 'records': []})

    for index, row in df.iterrows():
        content = row['订单项目.本周进度及问题反馈']
        person = row['周报人']
        project = row['订单项目.立项项目']
        days = row['订单项目.本周投入天数（最低半天）']
        week = row['周次']

        # 进行语义分析
        analysis = analyze_work_content_semantic(content, person, project, days)

        # 构建完整的记录分析
        record_analysis = {
            'index': index + 1,
            'person': person,
            'project': project,
            'week': week,
            'days': days,
            'content': content,
            'analysis': analysis
        }

        all_analyses.append(record_analysis)

        # 统计各类型工作量
        work_type = analysis['type']
        work_type_stats[work_type]['count'] += 1
        if not pd.isna(days):
            work_type_stats[work_type]['total_days'] += days
        work_type_stats[work_type]['records'].append(record_analysis)

        # 显示进度
        if (index + 1) % 50 == 0:
            print(f"已分析 {index + 1}/{len(df)} 条记录...")

    print(f"分析完成！共分析 {len(df)} 条记录")

    return all_analyses, dict(work_type_stats)

def analyze_projects_detailed(df, all_analyses):
    """基于详细分析结果进行项目分组统计"""

    project_analysis = {}

    for project in df['订单项目.立项项目'].unique():
        if pd.isna(project):
            continue

        # 获取该项目的所有分析记录
        project_records = [analysis for analysis in all_analyses
                          if analysis['project'] == project]

        if not project_records:
            continue

        # 统计基本信息
        total_days = sum(record['days'] for record in project_records)
        people_count = len(set(record['person'] for record in project_records))

        # 按工作类型分组
        type_groups = defaultdict(list)
        for record in project_records:
            work_type = record['analysis']['type']
            type_groups[work_type].append(record)

        # 计算各类型工作量
        type_stats = {}
        for work_type, records in type_groups.items():
            type_stats[work_type] = {
                'count': len(records),
                'total_days': sum(record['days'] for record in records),
                'records': records
            }

        # 计算筛选后的工作量（排除设备调机）
        core_work_days = sum(
            stats['total_days'] for work_type, stats in type_stats.items()
            if work_type != 'equipment_tuning'
        )

        equipment_tuning_days = type_stats.get('equipment_tuning', {}).get('total_days', 0)

        project_analysis[project] = {
            'total_days': total_days,
            'core_work_days': core_work_days,
            'equipment_tuning_days': equipment_tuning_days,
            'people_count': people_count,
            'record_count': len(project_records),
            'type_stats': type_stats,
            'all_records': project_records
        }

    return project_analysis

def generate_detailed_report(all_analyses, work_type_stats, project_analysis):
    """生成详细的语义分析报告"""

    print("\n" + "="*100)
    print("基于语义理解的项目数据深度分析报告")
    print("="*100)

    # 总体统计
    total_records = len(all_analyses)
    total_days = sum(analysis['days'] for analysis in all_analyses if not pd.isna(analysis['days']))

    print(f"\n【总体统计】")
    print(f"工作记录总数: {total_records} 条")
    print(f"总工作量: {total_days:.1f} 人天")
    print(f"项目总数: {len(project_analysis)} 个")

    # 工作类型分布统计
    print(f"\n【工作类型分布统计】")
    print(f"{'工作类型':<25} {'记录数':<8} {'工作量(人天)':<12} {'占比':<8}")
    print("-" * 60)

    # 定义工作类型的中文名称
    type_names = {
        'software_development': '软件开发',
        'software_maintenance': '软件维护',
        'system_integration': '系统集成',
        'equipment_tuning': '设备调机',
        'learning_research': '学习研究',
        'other_work': '其他工作'
    }

    for work_type, stats in sorted(work_type_stats.items(), key=lambda x: x[1]['total_days'], reverse=True):
        type_name = type_names.get(work_type, work_type)
        count = stats['count']
        days = stats['total_days']
        percentage = (days / total_days) * 100
        print(f"{type_name:<25} {count:<8} {days:<12.1f} {percentage:<8.1f}%")

    # 核心开发工作统计（排除设备调机）
    core_work_days = total_days - work_type_stats.get('equipment_tuning', {}).get('total_days', 0)
    equipment_days = work_type_stats.get('equipment_tuning', {}).get('total_days', 0)

    print(f"\n【核心开发工作统计】")
    print(f"核心开发工作量: {core_work_days:.1f} 人天 ({(core_work_days/total_days*100):.1f}%)")
    print(f"设备调机工作量: {equipment_days:.1f} 人天 ({(equipment_days/total_days*100):.1f}%)")

    # 技术领域分布
    print(f"\n【技术领域分布】")
    tech_area_stats = defaultdict(lambda: {'count': 0, 'days': 0.0})

    for analysis in all_analyses:
        tech_area = analysis['analysis']['technical_area']
        tech_area_stats[tech_area]['count'] += 1
        tech_area_stats[tech_area]['days'] += analysis['days']

    tech_area_names = {
        'frontend': '前端开发',
        'backend': '后端开发',
        'algorithm_ai': '算法/AI',
        'device_control': '设备控制',
        'system_architecture': '系统架构',
        'data_processing': '数据处理',
        'testing': '测试验证',
        'hardware_equipment': '硬件设备',
        'integration': '系统集成',
        'general': '通用技术',
        'unknown': '未分类'
    }

    print(f"{'技术领域':<20} {'记录数':<8} {'工作量(人天)':<12} {'占比':<8}")
    print("-" * 55)

    for tech_area, stats in sorted(tech_area_stats.items(), key=lambda x: x[1]['days'], reverse=True):
        if stats['days'] > 0:  # 只显示有工作量的领域
            area_name = tech_area_names.get(tech_area, tech_area)
            count = stats['count']
            days = stats['days']
            percentage = (days / total_days) * 100
            print(f"{area_name:<20} {count:<8} {days:<12.1f} {percentage:<8.1f}%")

def print_detailed_records(all_analyses, limit=20):
    """打印详细的记录分析结果"""

    print(f"\n【详细记录分析】(显示前{limit}条)")
    print("="*120)

    for i, analysis in enumerate(all_analyses[:limit]):
        record = analysis
        work_analysis = record['analysis']

        print(f"\n记录 #{record['index']}")
        print(f"人员: {record['person']}")
        print(f"项目: {record['project']}")
        print(f"工作量: {record['days']} 人天")
        print(f"工作内容: {record['content'][:100]}{'...' if len(record['content']) > 100 else ''}")
        print(f"分析结果:")
        print(f"  - 工作类型: {work_analysis['type']}")
        print(f"  - 子类型: {work_analysis['subtype']}")
        print(f"  - 技术领域: {work_analysis['technical_area']}")
        print(f"  - 工作性质: {work_analysis['work_nature']}")
        print(f"  - 分析理由: {work_analysis['analysis_reason']}")
        print(f"  - 置信度: {work_analysis['confidence']:.2f}")
        print("-" * 120)

def main():
    file_path = "2025年1-6.csv"

    try:
        # 读取CSV文件
        df, encoding = read_csv_with_encoding(file_path)

        print(f"文件基本信息:")
        print(f"总行数: {len(df)}")
        print(f"列数: {len(df.columns)}")
        print(f"使用编码: {encoding}")

        # 逐条分析所有记录
        all_analyses, work_type_stats = analyze_all_records(df)

        # 基于详细分析进行项目分组
        project_analysis = analyze_projects_detailed(df, all_analyses)

        # 生成详细报告
        generate_detailed_report(all_analyses, work_type_stats, project_analysis)

        # 打印部分详细记录
        print_detailed_records(all_analyses, limit=10)

        # 保存详细分析结果
        save_analysis_results(all_analyses, work_type_stats, project_analysis)

        print(f"\n分析完成！详细结果已保存到相关文件中。")

        return df, all_analyses, work_type_stats, project_analysis

    except Exception as e:
        print(f"处理文件时出错: {e}")
        import traceback
        traceback.print_exc()
        return None, None, None, None

def save_analysis_results(all_analyses, work_type_stats, project_analysis):
    """保存分析结果到文件"""

    # 保存逐条分析结果
    with open('detailed_record_analysis.json', 'w', encoding='utf-8') as f:
        serializable_analyses = []
        for analysis in all_analyses:
            serializable_analysis = {
                'index': analysis['index'],
                'person': analysis['person'],
                'project': analysis['project'],
                'week': analysis['week'],
                'days': float(analysis['days']),
                'content': analysis['content'],
                'analysis': {
                    'type': analysis['analysis']['type'],
                    'subtype': analysis['analysis']['subtype'],
                    'technical_area': analysis['analysis']['technical_area'],
                    'work_nature': analysis['analysis']['work_nature'],
                    'analysis_reason': analysis['analysis']['analysis_reason'],
                    'confidence': float(analysis['analysis']['confidence'])
                }
            }
            serializable_analyses.append(serializable_analysis)

        json.dump(serializable_analyses, f, ensure_ascii=False, indent=2)

    # 保存工作类型统计
    with open('work_type_statistics.json', 'w', encoding='utf-8') as f:
        serializable_stats = {}
        for work_type, stats in work_type_stats.items():
            serializable_stats[work_type] = {
                'count': stats['count'],
                'total_days': float(stats['total_days']),
                'average_days': float(stats['total_days'] / stats['count']) if stats['count'] > 0 else 0
            }
        json.dump(serializable_stats, f, ensure_ascii=False, indent=2)

    # 保存项目分析结果
    with open('project_detailed_analysis.json', 'w', encoding='utf-8') as f:
        serializable_projects = {}
        for project, analysis in project_analysis.items():
            serializable_projects[project] = {
                'total_days': float(analysis['total_days']),
                'core_work_days': float(analysis['core_work_days']),
                'equipment_tuning_days': float(analysis['equipment_tuning_days']),
                'people_count': analysis['people_count'],
                'record_count': analysis['record_count'],
                'type_distribution': {
                    work_type: {
                        'count': stats['count'],
                        'total_days': float(stats['total_days'])
                    }
                    for work_type, stats in analysis['type_stats'].items()
                }
            }
        json.dump(serializable_projects, f, ensure_ascii=False, indent=2)

    print(f"已保存以下分析结果文件:")
    print(f"- detailed_record_analysis.json: 逐条记录分析结果")
    print(f"- work_type_statistics.json: 工作类型统计")
    print(f"- project_detailed_analysis.json: 项目详细分析")

if __name__ == "__main__":
    df = main()

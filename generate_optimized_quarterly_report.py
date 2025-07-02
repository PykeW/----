#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成优化格式的季度工时统计报告
实现单元格合并效果和人员单独显示
"""

import pandas as pd
import chardet

def load_quarterly_data(file_path):
    """加载季度统计数据"""
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

def parse_personnel_data(row):
    """解析人员和人天数据"""
    personnel_str = str(row['人员'])
    person_days_str = str(row['人天'])
    
    # 处理空值或NaN
    if personnel_str in ['nan', 'None', ''] or person_days_str in ['nan', 'None', '']:
        return []
    
    # 分割人员和人天数据
    personnel_list = [p.strip() for p in personnel_str.split(';')]
    person_days_list = [float(d.strip()) for d in person_days_str.split(';')]
    
    # 确保两个列表长度一致
    if len(personnel_list) != len(person_days_list):
        print(f"警告: 人员和人天数据长度不匹配 - {personnel_str} vs {person_days_str}")
        min_len = min(len(personnel_list), len(person_days_list))
        personnel_list = personnel_list[:min_len]
        person_days_list = person_days_list[:min_len]
    
    # 创建人员-人天对，并按人天数降序排序
    person_data = list(zip(personnel_list, person_days_list))
    person_data.sort(key=lambda x: x[1], reverse=True)
    
    return person_data

def generate_optimized_report(df):
    """生成优化格式的报告"""
    print("正在生成优化格式的季度工时统计报告...")
    
    # 创建结果列表
    result_rows = []
    
    # 按部门、项目、季度排序
    df_sorted = df.sort_values(['订单项目.归属中心', '订单项目.立项项目', '季度'])
    
    current_dept = ""
    current_project = ""
    current_quarter = ""
    
    for _, row in df_sorted.iterrows():
        dept = row['订单项目.归属中心']
        project = row['订单项目.立项项目']
        quarter = int(row['季度'])
        total_days = float(row['总人天'])
        
        # 解析人员数据
        person_data = parse_personnel_data(row)
        
        if not person_data:
            # 如果没有人员数据，仍然添加一行
            result_rows.append({
                '订单项目.归属中心': dept if dept != current_dept else "",
                '订单项目.立项项目': project if project != current_project or dept != current_dept else "",
                '季度': quarter if quarter != current_quarter or project != current_project or dept != current_dept else "",
                '总人天': f"{total_days:.1f}" if quarter != current_quarter or project != current_project or dept != current_dept else "",
                '人员': "",
                '人天': ""
            })
        else:
            # 为每个人员创建一行
            for i, (person, person_days) in enumerate(person_data):
                # 第一行显示完整信息，后续行只显示人员信息
                if i == 0:
                    result_rows.append({
                        '订单项目.归属中心': dept if dept != current_dept else "",
                        '订单项目.立项项目': project if project != current_project or dept != current_dept else "",
                        '季度': quarter if quarter != current_quarter or project != current_project or dept != current_dept else "",
                        '总人天': f"{total_days:.1f}" if quarter != current_quarter or project != current_project or dept != current_dept else "",
                        '人员': person,
                        '人天': f"{person_days:.1f}"
                    })
                else:
                    result_rows.append({
                        '订单项目.归属中心': "",
                        '订单项目.立项项目': "",
                        '季度': "",
                        '总人天': "",
                        '人员': person,
                        '人天': f"{person_days:.1f}"
                    })
        
        current_dept = dept
        current_project = project
        current_quarter = quarter
    
    # 转换为DataFrame
    result_df = pd.DataFrame(result_rows)
    
    return result_df

def save_optimized_report(df, output_file):
    """保存优化格式的报告"""
    # 保存为CSV文件，使用UTF-8-BOM编码
    df.to_csv(output_file, index=False, encoding='utf-8-sig')
    print(f"✅ 优化格式报告已保存到: {output_file}")
    
    # 打印预览
    print("\n" + "=" * 120)
    print("📊 优化格式季度工时统计报告预览")
    print("=" * 120)
    
    print(f"{'部门':<25} {'项目':<45} {'季度':<8} {'总人天':<10} {'人员':<12} {'人天':<8}")
    print("-" * 120)
    
    for _, row in df.head(30).iterrows():  # 只显示前30行作为预览
        dept = row['订单项目.归属中心']
        project = row['订单项目.立项项目']
        quarter = row['季度']
        total_days = row['总人天']
        person = row['人员']
        person_days = row['人天']
        
        # 处理长文本截断
        if len(str(project)) > 43:
            project = str(project)[:40] + "..."
        
        # 格式化显示
        quarter_display = f"第{quarter}季度" if quarter else ""
        
        print(f"{dept:<25} {project:<45} {quarter_display:<8} {total_days:<10} {person:<12} {person_days:<8}")
    
    if len(df) > 30:
        print(f"... (还有 {len(df) - 30} 行数据)")
    
    print("-" * 120)
    print(f"总计: {len(df)} 行数据")

def generate_statistics(df):
    """生成统计信息"""
    print(f"\n📈 优化报告统计信息:")
    
    # 统计非空部门行数（即项目数）
    dept_rows = df[df['订单项目.归属中心'] != '']
    project_count = len(dept_rows)
    
    # 统计人员行数
    person_rows = df[df['人员'] != '']
    person_count = len(person_rows)
    
    # 统计季度
    quarter_rows = df[df['季度'] != '']
    quarters = quarter_rows['季度'].unique() if len(quarter_rows) > 0 else []
    
    print(f"   项目记录数: {project_count}")
    print(f"   人员记录数: {person_count}")
    print(f"   涉及季度: {sorted([int(q) for q in quarters if str(q) != ''])}")
    
    # 部门统计
    if len(dept_rows) > 0:
        dept_counts = dept_rows['订单项目.归属中心'].value_counts()
        print(f"\n🏢 各部门项目数:")
        for dept, count in dept_counts.items():
            print(f"   {dept}: {count}个项目")

def main():
    """主函数"""
    input_file = '季度工时统计报告.csv'
    output_file = '优化格式季度工时统计报告.csv'
    
    try:
        # 加载数据
        print("正在加载季度统计数据...")
        df = load_quarterly_data(input_file)
        
        if df is None:
            print("❌ 无法加载数据文件")
            return
        
        print(f"成功加载 {len(df)} 行数据")
        print(f"列名: {list(df.columns)}")
        
        # 生成优化报告
        optimized_df = generate_optimized_report(df)
        
        # 保存报告
        save_optimized_report(optimized_df, output_file)
        
        # 生成统计信息
        generate_statistics(optimized_df)
        
        print("\n" + "=" * 80)
        print("✅ 优化格式季度报告生成完成！")
        print("=" * 80)
        print(f"📁 输出文件: {output_file}")
        print("💡 提示: 在Excel中打开时，空白单元格表示与上方单元格合并")
        
    except Exception as e:
        print(f"❌ 生成报告过程中出现错误: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试NDT_result_mode1.py的功能
"""

import pandas as pd

def test_column_mapping():
    """测试列映射功能"""
    print("=== 测试列映射功能 ===")
    
    # 读取Excel文件
    df = pd.read_excel('生成器/Excel/2_生成器结果.xlsx')
    print(f"Excel文件读取成功，共{len(df)}行数据")
    
    # 建立列名映射
    column_mapping = {}
    for col in df.columns:
        col_str = str(col).strip()
        if '完成日期' in col_str:
            column_mapping['完成日期'] = col
        elif '委托单编号' in col_str:
            column_mapping['委托单编号'] = col
        elif '检件编号' in col_str:
            column_mapping['检件编号'] = col
        elif '焊口编号' in col_str:
            column_mapping['焊口编号'] = col
        elif '规格' in col_str and '底片' not in col_str:
            column_mapping['规格'] = col
        elif '材质' in col_str:
            column_mapping['材质'] = col
        elif '合格级别' in col_str:
            column_mapping['合格级别'] = col
        elif '底片规格' in col_str or col_str == '底片规格/张数':
            column_mapping['底片规格/张数'] = col
        elif col_str == '张数':
            column_mapping['张数'] = col
        elif '合格张数' in col_str:
            column_mapping['合格张数'] = col
        elif '单元名称' in col_str:
            column_mapping['单元名称'] = col
    
    print("找到的列映射:")
    for key, value in column_mapping.items():
        print(f"  {key}: {value}")
    
    # 测试第一个委托单编号的数据
    order_column = column_mapping['委托单编号']
    first_order = df[order_column].iloc[0]
    test_data = df[df[order_column] == first_order]
    
    print(f"\n=== 测试委托单编号: {first_order} ===")
    print(f"该组数据行数: {len(test_data)}")
    
    # 测试数据提取
    print("\n=== 数据提取测试 ===")
    
    # 检件编号
    if '检件编号' in column_mapping:
        pipe_numbers = []
        for idx, row in test_data.iterrows():
            pipe_num = row[column_mapping['检件编号']]
            pipe_numbers.append(str(pipe_num) if pd.notna(pipe_num) else "")
        print(f"检件编号: {pipe_numbers}")
    
    # 焊口编号
    if '焊口编号' in column_mapping:
        weld_numbers = []
        for idx, row in test_data.iterrows():
            weld_num = row[column_mapping['焊口编号']]
            weld_numbers.append(str(weld_num) if pd.notna(weld_num) else "")
        print(f"焊口编号: {weld_numbers}")
    
    # 材质
    if '材质' in column_mapping:
        materials = []
        for idx, row in test_data.iterrows():
            material = row[column_mapping['材质']]
            materials.append(str(material) if pd.notna(material) else "")
        print(f"材质: {materials}")
    
    # 规格
    if '规格' in column_mapping:
        specifications = []
        for idx, row in test_data.iterrows():
            spec = row[column_mapping['规格']]
            specifications.append(str(spec) if pd.notna(spec) else "")
        print(f"规格: {specifications}")
    
    # 底片规格/张数
    if '底片规格/张数' in column_mapping:
        film_specs = []
        for idx, row in test_data.iterrows():
            film_spec = row[column_mapping['底片规格/张数']]
            film_specs.append(str(film_spec) if pd.notna(film_spec) else "")
        print(f"底片规格/张数: {film_specs}")
    
    # 合格张数
    if '合格张数' in column_mapping:
        qualified_counts = []
        for idx, row in test_data.iterrows():
            qualified = row[column_mapping['合格张数']]
            qualified_counts.append(str(qualified) if pd.notna(qualified) else "0")
        print(f"合格张数: {qualified_counts}")
    
    # 不合格张数计算
    if '张数' in column_mapping and '合格张数' in column_mapping:
        unqualified_counts = []
        for idx, row in test_data.iterrows():
            total_count = 0
            qualified_count = 0
            
            if '张数' in column_mapping:
                total = row[column_mapping['张数']]
                if pd.notna(total):
                    try:
                        total_count = int(float(total))
                    except:
                        total_count = 0
            
            if '合格张数' in column_mapping:
                qualified = row[column_mapping['合格张数']]
                if pd.notna(qualified):
                    try:
                        qualified_count = int(float(qualified))
                    except:
                        qualified_count = 0
            
            unqualified_count = max(0, total_count - qualified_count)
            unqualified_counts.append(str(unqualified_count))
        
        print(f"不合格张数: {unqualified_counts}")
    
    # 测试单值提取
    print("\n=== 单值提取测试 ===")
    
    # 合格级别
    if '合格级别' in column_mapping:
        qual_values = test_data[column_mapping['合格级别']].dropna()
        if not qual_values.empty:
            qualification_level = str(qual_values.iloc[0])
            print(f"合格级别: {qualification_level}")
    
    # 单元名称
    if '单元名称' in column_mapping:
        unit_values = test_data[column_mapping['单元名称']].dropna()
        if not unit_values.empty:
            unit_name = str(unit_values.iloc[0])
            print(f"单元名称: {unit_name}")
    
    # 完成日期
    if '完成日期' in column_mapping:
        completion_dates = test_data[column_mapping['完成日期']].dropna()
        if not completion_dates.empty:
            try:
                completion_dates_converted = pd.to_datetime(completion_dates, errors='coerce')
                latest_completion_date = completion_dates_converted.max()
                
                if pd.notna(latest_completion_date):
                    year = latest_completion_date.year
                    month = latest_completion_date.month
                    day = latest_completion_date.day
                    print(f"最晚完成日期: {year}年{month}月{day}日")
            except Exception as e:
                print(f"日期转换错误: {e}")

if __name__ == "__main__":
    test_column_mapping()

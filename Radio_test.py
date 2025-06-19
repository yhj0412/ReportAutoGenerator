import pandas as pd
import os
import sys
import argparse
from docx import Document
from datetime import datetime
import re

def find_column_with_keyword(df, keyword):
    """查找包含指定关键字的列"""
    matching_cols = [col for col in df.columns if keyword.lower() in col.lower()]
    return matching_cols[0] if matching_cols else None

def get_output_filename(word_template_path, order_number, ray_type):
    """根据Word模板路径、委托单编号和射线类型生成输出文件名
    
    Args:
        word_template_path: Word模板文档路径
        order_number: 委托单编号
        ray_type: 射线类型（"X射线"或"γ射线"）
    
    Returns:
        str: 输出文件名
    """
    # 获取模板文件名（不含路径和扩展名）
    template_name = os.path.splitext(os.path.basename(word_template_path))[0]
    # 射线类型标识
    ray_mark = "γ" if ray_type == "γ射线" else "X"
    # 生成输出文件名
    return f"{template_name}_{order_number}_{ray_mark}_生成结果.docx"

def process_excel_to_word(excel_path, word_template_path, output_path=None, 
                       project_name=None, entrusting_unit=None, 
                       operation_guide_number=None, contracting_unit=None, 
                       equipment_model=None):
    """将Excel数据填入Word文档
    
    Args:
        excel_path: Excel表格路径
        word_template_path: Word模板文档路径
        output_path: 输出Word文档路径（如果为None，将自动生成）
        project_name: 工程名称，用于替换文档中的"工程名称值"
        entrusting_unit: 委托单位，用于替换文档中的"委托单位值"
        operation_guide_number: 操作指导书编号，用于替换文档中的"操作指导书编号值"
        contracting_unit: 承包单位，用于替换文档中的"承包单位值"
        equipment_model: 设备型号，用于替换文档中的"设备型号值"
    
    Returns:
        bool: 处理是否成功
    """
    # 检查文件是否存在
    if not os.path.exists(excel_path):
        print(f"错误: Excel文件不存在: {excel_path}")
        return False
        
    if not os.path.exists(word_template_path):
        print(f"错误: Word模板文件不存在: {word_template_path}")
        return False
    
    # 创建输出目录
    if output_path is None:
        output_dir = os.path.join("生成器", "输出报告", "4_射线检测记录")
    else:
        output_dir = output_path
        
    if not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
            print(f"创建输出目录: {output_dir}")
        except Exception as e:
            print(f"错误: 无法创建输出目录: {e}")
            return False
    
    # 读取Excel数据
    print(f"正在读取Excel文件: {excel_path}")
    try:
        df = pd.read_excel(excel_path)
        print(f"成功读取Excel文件，共有{len(df)}行数据")
    except Exception as e:
        print(f"错误: 无法读取Excel文件: {e}")
        return False
    
    # X射线参数表
    xray_params_path = os.path.join("生成器", "Excel", "4_生成器X射线指导书模版.xlsx")
    try:
        if os.path.exists(xray_params_path):
            xray_params_df = pd.read_excel(xray_params_path, sheet_name="曝光参数")
            print(f"成功读取X射线参数表，共有{len(xray_params_df)}行数据")
            # 打印X射线参数表的列名
            print(f"X射线参数表列名: {list(xray_params_df.columns)}")
        else:
            print(f"警告: X射线参数表不存在: {xray_params_path}，将不进行X射线参数匹配")
            xray_params_df = None
    except Exception as e:
        print(f"错误: 无法读取X射线参数表: {e}")
        xray_params_df = None
    
    # γ射线参数表
    gamma_params_path = os.path.join("生成器", "Excel", "4_生成器γ射线指导书模版.xlsx")
    try:
        if os.path.exists(gamma_params_path):
            gamma_params_df = pd.read_excel(gamma_params_path, sheet_name="曝光参数")
            print(f"成功读取γ射线参数表，共有{len(gamma_params_df)}行数据")
            # 打印γ射线参数表的列名
            print(f"γ射线参数表列名: {list(gamma_params_df.columns)}")
        else:
            print(f"警告: γ射线参数表不存在: {gamma_params_path}，将不进行γ射线参数匹配")
            gamma_params_df = None
    except Exception as e:
        print(f"错误: 无法读取γ射线参数表: {e}")
        gamma_params_df = None
    
    # 打印所有列名，帮助调试
    print(f"Excel表格列名: {list(df.columns)}")
    
    # 定义需要查找的列关键字
    column_keywords = {
        '完成日期': '完成日期',
        '委托单编号': '委托单编号',
        '检件编号': '检件编号',
        '焊口编号': '焊口编号',
        '焊工号': '焊工号',
        '合格级别': '合格级别',
        '检测比例': '检测比例',
        'γ射线': 'γ射线',
        '焊接方法': '焊接方法',
        '检测时机': '检测时机',
        '规格': '规格'
    }
    
    # 查找每个关键字对应的实际列名
    column_mapping = {}
    missing_columns = []
    
    for key, keyword in column_keywords.items():
        col_name = find_column_with_keyword(df, keyword)
        if col_name:
            column_mapping[key] = col_name
            print(f"找到列: '{key}' -> '{col_name}'")
        else:
            missing_columns.append(key)
            
    if missing_columns:
        print(f"警告: 未找到以下列: {', '.join(missing_columns)}")
        # 尝试使用列位置
        possible_columns = {
            '完成日期': 'B', 
            '委托单编号': 'C', 
            '检件编号': 'D', 
            '焊口编号': 'E', 
            '焊工号': 'F',
            '合格级别': 'I',
            '检测比例': 'J',
            'γ射线': 'P',
            '焊接方法': 'R',
            '检测时机': 'V',
            '规格': 'G'
        }
        
        for key in missing_columns:
            col_letter = possible_columns.get(key)
            if col_letter:
                col_idx = ord(col_letter) - ord('A')
                if col_idx < len(df.columns):
                    column_mapping[key] = df.columns[col_idx]
                    print(f"使用列位置找到: '{key}' -> '{df.columns[col_idx]}'")
    
    # 如果缺少关键列，则无法继续处理
    if '委托单编号' not in column_mapping:
        print("错误: 无法找到委托单编号列")
        return False
    
    if 'γ射线' not in column_mapping:
        print("警告: 无法找到γ射线列，将默认所有记录为X射线")
        # 增加一个全为空的列作为γ射线列
        df['γ射线'] = None
        column_mapping['γ射线'] = 'γ射线'
    
    # 根据委托单编号和射线类型分组数据
    groups = []
    order_numbers = df[column_mapping['委托单编号']].dropna().unique()
    
    for order_number in order_numbers:
        # 获取该委托单编号的所有数据
        order_df = df[df[column_mapping['委托单编号']] == order_number]
        
        # 获取该委托单编号下的所有射线类型
        ray_types = order_df[column_mapping['γ射线']].dropna().unique()
        
        # 如果没有明确的γ射线值，则视为X射线
        if len(ray_types) == 0:
            groups.append({
                'order_number': order_number,
                'ray_type': 'X射线',  # X射线表示为处理逻辑，但射源种类值会设置为空
                'data': order_df
            })
            print(f"委托单编号 {order_number} 没有明确的射线类型，处理为X射线（射源种类值为空）")
        else:
            # 有γ射线值的处理为γ射线
            for ray_type in ray_types:
                ray_df = order_df[order_df[column_mapping['γ射线']] == ray_type]
                groups.append({
                    'order_number': order_number,
                    'ray_type': 'γ射线',
                    'data': ray_df
                })
                print(f"委托单编号 {order_number} 的射线类型 γ射线 有 {len(ray_df)} 条记录")
            
            # 没有γ射线值的处理为X射线
            x_ray_df = order_df[order_df[column_mapping['γ射线']].isna()]
            if len(x_ray_df) > 0:
                groups.append({
                    'order_number': order_number,
                    'ray_type': 'X射线',  # X射线表示为处理逻辑，但射源种类值会设置为空
                    'data': x_ray_df
                })
                print(f"委托单编号 {order_number} 的射线类型 X射线 有 {len(x_ray_df)} 条记录")
    
    print(f"共有 {len(groups)} 个组合需要生成报告")
    
    # 处理每个分组
    success_count = 0
    error_count = 0
    
    for group in groups:
        order_number = group['order_number']
        ray_type = group['ray_type']
        group_df = group['data']
        
        print(f"\n{'='*50}")
        print(f"处理委托单编号: {order_number}, 射线类型: {ray_type}")
        print(f"{'='*50}")
        
        try:
            # 为该分组生成输出文件名
            output_filename = get_output_filename(word_template_path, order_number, ray_type)
            report_output_path = os.path.join(output_dir, output_filename)
            print(f"输出文件路径: {report_output_path}")
            
            # 打开Word文档
            print(f"正在处理Word文档: {word_template_path}")
            
            try:
                # 每次处理新的组合时，重新从模板创建文档对象
                # 这确保了每个组合都会生成一个独立的文档
                doc = Document(word_template_path)
                print(f"成功从模板创建新文档")
            except Exception as e:
                print(f"无法打开Word文档: {e}")
                error_count += 1
                continue
            
            # 填充文档的其余部分将在这里添加...
            
            # 1) 获取该组数据中最晚的完成日期
            date_col = column_mapping.get('完成日期')
            if date_col:
                # 确保日期列是日期类型
                group_df[date_col] = pd.to_datetime(group_df[date_col], errors='coerce')
                latest_date = group_df[date_col].max()
                
                if pd.isna(latest_date):
                    print(f"警告: 委托单编号 {order_number} 没有有效的完成日期")
                    year, month, day = datetime.now().year, datetime.now().month, datetime.now().day
                else:
                    # 将日期转换为年、月、日
                    year = latest_date.year
                    month = latest_date.month
                    day = latest_date.day
                    print(f"找到最晚完成日期: {year}年{month}月{day}日")
            else:
                print("警告: 未找到完成日期列")
                year, month, day = datetime.now().year, datetime.now().month, datetime.now().day
            
            # 获取相关数据
            inspection_numbers = group_df[column_mapping.get('检件编号')].dropna().tolist() if '检件编号' in column_mapping else []
            weld_numbers = group_df[column_mapping.get('焊口编号')].dropna().tolist() if '焊口编号' in column_mapping else []
            welder_numbers = group_df[column_mapping.get('焊工号')].dropna().tolist() if '焊工号' in column_mapping else []
            # 对规格数据进行去重处理
            specifications = list(group_df[column_mapping.get('规格')].dropna().unique()) if '规格' in column_mapping else []
            print(f"规格列去重前数量: {len(group_df[column_mapping.get('规格')].dropna().tolist() if '规格' in column_mapping else [])}")
            print(f"规格列去重后数量: {len(specifications)}")
            print(f"去重后规格列数据: {specifications}")
            
            # 获取分组的第一个值
            grade_level = ""
            if '合格级别' in column_mapping:
                grade_levels = group_df[column_mapping['合格级别']].dropna().tolist()
                if grade_levels:
                    grade_level = grade_levels[0]
                    print(f"找到合格级别: {grade_level}")
            
            inspection_ratio = ""
            if '检测比例' in column_mapping:
                ratios = group_df[column_mapping['检测比例']].dropna().tolist()
                if ratios:
                    # 转换为百分数格式
                    try:
                        ratio_value = float(ratios[0])
                        inspection_ratio = f"{ratio_value*100:.0f}%"
                    except (ValueError, TypeError):
                        inspection_ratio = str(ratios[0])
                    print(f"找到检测比例: {inspection_ratio}")
            
            welding_method = ""
            if '焊接方法' in column_mapping:
                methods = group_df[column_mapping['焊接方法']].dropna().tolist()
                if methods:
                    welding_method = methods[0]
                    print(f"找到焊接方法: {welding_method}")
            
            inspection_time = ""
            if '检测时机' in column_mapping:
                times = group_df[column_mapping['检测时机']].dropna().tolist()
                if times:
                    inspection_time = times[0]
                    print(f"找到检测时机: {inspection_time}")
            
            # 根据射线类型设置相关参数
            if ray_type == "γ射线":
                focus_size = "3*3"
                lead_screen = "柯达0.1*2"
                film_grade = "柯达MX125"
                ray_source = "γ射线"
            else:  # X射线
                focus_size = "2.5*2.5"
                lead_screen = "0.03*2"
                film_grade = "锐科R400"
                ray_source = ""  # 当γ射线的值为空时，射源种类值也设置为空
            
            print(f"射线类型: {ray_type}, 焦点尺寸: {focus_size}, 铅增感屏: {lead_screen}, 胶片等级: {film_grade}, 射源种类值: {ray_source}")
            
            # 替换文档中的值
            print("\n==== 开始替换文档中的值 ====")
            
            # 将EPKJ拼接委托单编号代替委托单编号值
            committee_order = f"EPKJ-{order_number}"
            replaced = False
            
            # 遍历所有段落，替换关键词
            for paragraph in doc.paragraphs:
                if "委托单编号值" in paragraph.text:
                    paragraph.text = paragraph.text.replace("委托单编号值", committee_order)
                    print(f"已将段落中的'委托单编号值'替换为'{committee_order}'")
                    replaced = True
                
                if "射源种类值" in paragraph.text:
                    paragraph.text = paragraph.text.replace("射源种类值", ray_source)
                    print(f"已将段落中的'射源种类值'替换为'{ray_source}'")
                    replaced = True
                
                if "合格级别值" in paragraph.text:
                    paragraph.text = paragraph.text.replace("合格级别值", grade_level)
                    print(f"已将段落中的'合格级别值'替换为'{grade_level}'")
                    replaced = True
                
                if "检测比例值" in paragraph.text:
                    paragraph.text = paragraph.text.replace("检测比例值", inspection_ratio)
                    print(f"已将段落中的'检测比例值'替换为'{inspection_ratio}'")
                    replaced = True
                
                if "焊接方法值" in paragraph.text:
                    paragraph.text = paragraph.text.replace("焊接方法值", welding_method)
                    print(f"已将段落中的'焊接方法值'替换为'{welding_method}'")
                    replaced = True
                
                if "检测时机值" in paragraph.text:
                    paragraph.text = paragraph.text.replace("检测时机值", inspection_time)
                    print(f"已将段落中的'检测时机值'替换为'{inspection_time}'")
                    replaced = True
                
                if "焦点尺寸值" in paragraph.text:
                    paragraph.text = paragraph.text.replace("焦点尺寸值", focus_size)
                    print(f"已将段落中的'焦点尺寸值'替换为'{focus_size}'")
                    replaced = True
                
                if "铅增感屏值" in paragraph.text:
                    paragraph.text = paragraph.text.replace("铅增感屏值", lead_screen)
                    print(f"已将段落中的'铅增感屏值'替换为'{lead_screen}'")
                    replaced = True
                
                if "胶片等级值" in paragraph.text:
                    paragraph.text = paragraph.text.replace("胶片等级值", film_grade)
                    print(f"已将段落中的'胶片等级值'替换为'{film_grade}'")
                    replaced = True
                
                # 新增的5个参数替换
                if "工程名称值" in paragraph.text and project_name:
                    paragraph.text = paragraph.text.replace("工程名称值", project_name)
                    print(f"已将段落中的'工程名称值'替换为'{project_name}'")
                    replaced = True
                
                if "委托单位值" in paragraph.text and entrusting_unit:
                    paragraph.text = paragraph.text.replace("委托单位值", entrusting_unit)
                    print(f"已将段落中的'委托单位值'替换为'{entrusting_unit}'")
                    replaced = True
                
                if "操作指导书编号值" in paragraph.text and operation_guide_number:
                    paragraph.text = paragraph.text.replace("操作指导书编号值", operation_guide_number)
                    print(f"已将段落中的'操作指导书编号值'替换为'{operation_guide_number}'")
                    replaced = True
                
                if "承包单位值" in paragraph.text and contracting_unit:
                    paragraph.text = paragraph.text.replace("承包单位值", contracting_unit)
                    print(f"已将段落中的'承包单位值'替换为'{contracting_unit}'")
                    replaced = True
                
                if "设备型号值" in paragraph.text and equipment_model:
                    paragraph.text = paragraph.text.replace("设备型号值", equipment_model)
                    print(f"已将段落中的'设备型号值'替换为'{equipment_model}'")
                    replaced = True
            
            # 遍历表格中的单元格，替换关键词
            for table in doc.tables:
                for i, row in enumerate(table.rows):
                    for j, cell in enumerate(row.cells):
                        for paragraph in cell.paragraphs:
                            if "委托单编号值" in paragraph.text:
                                paragraph.text = paragraph.text.replace("委托单编号值", committee_order)
                                print(f"已将表格单元格中的'委托单编号值'替换为'{committee_order}'")
                                replaced = True
                            
                            if "射源种类值" in paragraph.text:
                                paragraph.text = paragraph.text.replace("射源种类值", ray_source)
                                print(f"已将表格单元格中的'射源种类值'替换为'{ray_source}'")
                                replaced = True
                            
                            if "合格级别值" in paragraph.text:
                                paragraph.text = paragraph.text.replace("合格级别值", grade_level)
                                print(f"已将表格单元格中的'合格级别值'替换为'{grade_level}'")
                                replaced = True
                            
                            if "检测比例值" in paragraph.text:
                                paragraph.text = paragraph.text.replace("检测比例值", inspection_ratio)
                                print(f"已将表格单元格中的'检测比例值'替换为'{inspection_ratio}'")
                                replaced = True
                            
                            if "焊接方法值" in paragraph.text:
                                paragraph.text = paragraph.text.replace("焊接方法值", welding_method)
                                print(f"已将表格单元格中的'焊接方法值'替换为'{welding_method}'")
                                replaced = True
                            
                            if "检测时机值" in paragraph.text:
                                paragraph.text = paragraph.text.replace("检测时机值", inspection_time)
                                print(f"已将表格单元格中的'检测时机值'替换为'{inspection_time}'")
                                replaced = True
                            
                            if "焦点尺寸值" in paragraph.text:
                                paragraph.text = paragraph.text.replace("焦点尺寸值", focus_size)
                                print(f"已将表格单元格中的'焦点尺寸值'替换为'{focus_size}'")
                                replaced = True
                            
                            if "铅增感屏值" in paragraph.text:
                                paragraph.text = paragraph.text.replace("铅增感屏值", lead_screen)
                                print(f"已将表格单元格中的'铅增感屏值'替换为'{lead_screen}'")
                                replaced = True
                            
                            if "胶片等级值" in paragraph.text:
                                paragraph.text = paragraph.text.replace("胶片等级值", film_grade)
                                print(f"已将表格单元格中的'胶片等级值'替换为'{film_grade}'")
                                replaced = True
                            
                            # 新增的5个参数替换（表格单元格）
                            if "工程名称值" in paragraph.text and project_name:
                                paragraph.text = paragraph.text.replace("工程名称值", project_name)
                                print(f"已将表格单元格中的'工程名称值'替换为'{project_name}'")
                                replaced = True
                            
                            if "委托单位值" in paragraph.text and entrusting_unit:
                                paragraph.text = paragraph.text.replace("委托单位值", entrusting_unit)
                                print(f"已将表格单元格中的'委托单位值'替换为'{entrusting_unit}'")
                                replaced = True
                            
                            if "操作指导书编号值" in paragraph.text and operation_guide_number:
                                paragraph.text = paragraph.text.replace("操作指导书编号值", operation_guide_number)
                                print(f"已将表格单元格中的'操作指导书编号值'替换为'{operation_guide_number}'")
                                replaced = True
                            
                            if "承包单位值" in paragraph.text and contracting_unit:
                                paragraph.text = paragraph.text.replace("承包单位值", contracting_unit)
                                print(f"已将表格单元格中的'承包单位值'替换为'{contracting_unit}'")
                                replaced = True
                            
                            if "设备型号值" in paragraph.text and equipment_model:
                                paragraph.text = paragraph.text.replace("设备型号值", equipment_model)
                                print(f"已将表格单元格中的'设备型号值'替换为'{equipment_model}'")
                                replaced = True
            
            if not replaced:
                print("警告: 未找到需要替换的关键词，可能需要检查Word模板中的占位符命名。")
            
            # 填写日期（洗片人、拍片人、审核人）
            for table in doc.tables:
                for i, row in enumerate(table.rows):
                    for j, cell in enumerate(row.cells):
                        # 处理日期
                        date_patterns = ["洗片人", "拍片人", "审核人"]
                        for pattern in date_patterns:
                            if pattern in cell.text:
                                print(f"找到{pattern}单元格: 表格行{i+1}, 列{j+1}")
                                
                                # 检查单元格中的所有段落
                                date_found = False
                                for paragraph in cell.paragraphs:
                                    if "年" in paragraph.text and "月" in paragraph.text and "日" in paragraph.text:
                                        print(f"找到日期段落: {paragraph.text}")
                                        
                                        # 创建新的文本，确保只有一个年月日
                                        new_text = paragraph.text
                                        # 确保年月日前没有数字
                                        new_text = re.sub(r'\d*年', '年', new_text)
                                        new_text = re.sub(r'\d*月', '月', new_text)
                                        new_text = re.sub(r'\d*日', '日', new_text)
                                        
                                        # 在年月日前插入正确的数字
                                        new_text = new_text.replace('年', f'{year}年')
                                        new_text = new_text.replace('月', f'{month}月')
                                        new_text = new_text.replace('日', f'{day}日')
                                        
                                        paragraph.text = new_text
                                        date_found = True
                                        print(f"已更新{pattern}日期为 {year}年{month}月{day}日")
                                        break
                                
                                # 如果没有找到日期段落，尝试创建新段落
                                if not date_found:
                                    print(f"未在{pattern}单元格中找到日期段落，尝试添加")
                                    # 添加新段落
                                    p = cell.add_paragraph(f"{year}年{month}月{day}日")
                                    print(f"已添加{pattern}日期: {year}年{month}月{day}日")
            
            # 查找表头行，确定各列的位置
            for table in doc.tables:
                column_indices = {}
                header_row_index = -1
                
                # 添加完整的表格内容打印，帮助调试
                print("\n调试：打印表格内容以找到规格列")
                for i, row in enumerate(table.rows):
                    row_text = []
                    for j, cell in enumerate(row.cells):
                        row_text.append(f"[{j}]'{cell.text}'")
                    if len(row_text) > 0:  # 只打印非空行
                        print(f"行 {i}: {', '.join(row_text)}")
                
                # 查找规格表头(mm×mm)位于第10行左右，检件编号等位于第18行左右
                spec_column_index = -1
                for i in range(9, 12):  # 在第9-11行范围内查找规格列
                    if i < len(table.rows):
                        for j, cell in enumerate(table.rows[i].cells):
                            cell_text = cell.text.strip()
                            if "检件规格" in cell_text and ("mm×mm" in cell_text or "mm*mm" in cell_text):
                                spec_column_index = j
                                print(f"找到规格列（透照参数表格）：第{i+1}行，第{j+1}列，文本：{cell_text}")
                                break
                
                # 查找包含"检件编号"、"焊缝编号"、"焊工号"的行
                for i, row in enumerate(table.rows):
                    header_found = False
                    for j, cell in enumerate(row.cells):
                        cell_text = cell.text.strip()
                        
                        # 添加更详细的调试信息，输出表格单元格文本内容
                        if "检件规格" in cell_text or "规格" in cell_text:
                            print(f"找到可能的规格列： 行 {i+1}, 列 {j+1}, 文本: '{cell_text}'")
                        
                        if "检件编号" in cell_text:
                            column_indices["检件编号"] = j
                            header_row_index = i
                            header_found = True
                            print(f"找到检件编号列: 行 {i+1}, 列 {j+1}, 文本: '{cell_text}'")
                        elif "焊缝编号" in cell_text or "焊口编号" in cell_text:
                            column_indices["焊口编号"] = j
                            header_found = True
                            print(f"找到焊缝编号列: 行 {i+1}, 列 {j+1}, 文本: '{cell_text}'")
                        elif "焊工号" in cell_text:
                            column_indices["焊工号"] = j
                            header_found = True
                            print(f"找到焊工号列: 行 {i+1}, 列 {j+1}, 文本: '{cell_text}'")
                        elif "备注" in cell_text:
                            column_indices["备注"] = j
                            header_found = True
                            print(f"找到备注列: 行 {i+1}, 列 {j+1}, 文本: '{cell_text}'")
                        elif "检件规格" in cell_text or "规格" in cell_text or "检件规格(mm×mm)" in cell_text or "检件规格(mm*mm)" in cell_text or "检件规格(mm" in cell_text:
                            column_indices["检件规格"] = j
                            header_found = True
                            print(f"找到检件规格列: 行 {i+1}, 列 {j+1}, 文本: '{cell_text}'")
                        elif "透照参数序号" in cell_text:
                            column_indices["透照参数序号"] = j
                            header_found = True
                            print(f"找到透照参数序号列: 行 {i+1}, 列 {j+1}, 文本: '{cell_text}'")
                    
                    if header_found and header_row_index >= 0:
                        print(f"找到表头行: 第{header_row_index+1}行")
                        
                        # 将"焊口编号"的键名更新为"焊缝编号"以保持一致性
                        if "焊口编号" in column_indices:
                            column_indices["焊缝编号"] = column_indices.pop("焊口编号")
                            
                        print(f"列索引: {column_indices}")
                        break
                
                # 如果找到表头行，处理数据填充
                if header_row_index >= 0 and column_indices:
                    # 获取可用于填充数据的行
                    data_rows = []
                    for i in range(header_row_index + 1, len(table.rows)):
                        if i < len(table.rows):
                            # 检查是否是空行或包含特殊标记的行
                            if "以下空白" in table.rows[i].cells[0].text if len(table.rows[i].cells) > 0 else False:
                                print(f"找到'以下空白'行: 第{i+1}行")
                                break
                            # 添加可用于填充数据的行
                            data_rows.append(i)
                    
                    print(f"找到{len(data_rows)}行可用于填充数据")
                    
                    # 确定需要填充的数据行数
                    data_count = len(group_df)
                    print(f"需要填充{data_count}行数据")
                    
                    # 如果Word表格中的行数不足，需要添加新行
                    rows_needed = data_count - len(data_rows)
                    if rows_needed > 0:
                        print(f"需要添加{rows_needed}行到表格中")
                        # 找到最后一行的索引
                        last_row_idx = data_rows[-1] if data_rows else header_row_index
                        
                        # 添加新行
                        for _ in range(rows_needed):
                            # 在表格末尾添加一行
                            new_row = table.add_row()
                            data_rows.append(len(table.rows) - 1)  # 添加新行的索引
                    
                    # 处理每一行数据
                    for i in range(data_count):
                        if i < len(data_rows):
                            row_idx = data_rows[i]
                            row = table.rows[row_idx]
                            
                            # 1. 填写检件编号
                            if "检件编号" in column_indices and i < len(inspection_numbers):
                                col_idx = column_indices["检件编号"]
                                if col_idx < len(row.cells):
                                    cell = row.cells[col_idx]
                                    if cell.paragraphs:
                                        cell.paragraphs[0].text = str(inspection_numbers[i])
                                        print(f"已更新第{row_idx+1}行检件编号: {inspection_numbers[i]}")
                            
                            # 2. 填写焊缝编号
                            if "焊缝编号" in column_indices and i < len(weld_numbers):
                                col_idx = column_indices["焊缝编号"]
                                if col_idx < len(row.cells):
                                    cell = row.cells[col_idx]
                                    if cell.paragraphs:
                                        cell.paragraphs[0].text = str(weld_numbers[i])
                                        print(f"已更新第{row_idx+1}行焊缝编号: {weld_numbers[i]}")
                            
                            # 3. 填写焊工号
                            if "焊工号" in column_indices and i < len(welder_numbers):
                                col_idx = column_indices["焊工号"]
                                if col_idx < len(row.cells):
                                    cell = row.cells[col_idx]
                                    if cell.paragraphs:
                                        cell.paragraphs[0].text = str(welder_numbers[i])
                                        print(f"已更新第{row_idx+1}行焊工号: {welder_numbers[i]}")
                            
                            # 4. 填写备注（填入完成日期）
                            if "备注" in column_indices and date_col and i < len(group_df):
                                col_idx = column_indices["备注"]
                                if col_idx < len(row.cells):
                                    cell = row.cells[col_idx]
                                    # 获取对应行的完成日期
                                    if not pd.isna(group_df[date_col].iloc[i]):
                                        date_value = group_df[date_col].iloc[i]
                                        if isinstance(date_value, pd.Timestamp):
                                            # 如果是日期类型，格式化为字符串
                                            formatted_date = date_value.strftime("%Y年%m月%d日")
                                        else:
                                            # 如果不是日期类型，直接转为字符串
                                            formatted_date = str(date_value)
                                        
                                        if cell.paragraphs:
                                            cell.paragraphs[0].text = formatted_date
                                            print(f"已更新第{row_idx+1}行备注（完成日期）: {formatted_date}")
                            
                            # 5. 填写透照参数序号（规格数量）
                            if "透照参数序号" in column_indices:
                                col_idx = column_indices["透照参数序号"]
                                if col_idx < len(row.cells):
                                    cell = row.cells[col_idx]
                                    if cell.paragraphs:
                                        # 获取规格去重后的数量
                                        spec_count = len(specifications)
                                        # 直接填写规格总数，不再根据行号判断
                                        param_index = spec_count
                                        cell.paragraphs[0].text = str(param_index)
                                        print(f"已更新第{row_idx+1}行透照参数序号: {param_index}")
            
                    # 如果找到了规格列，在透照参数表中填写规格信息
                    if spec_column_index >= 0 and len(specifications) > 0:
                        # 透照参数表格一般在第10-15行，我们从第10行开始填充规格数据
                        start_row = 10
                        print(f"开始在透照参数表中填充去重后的{len(specifications)}种规格数据")
                        
                        # 清空现有规格数据
                        for i in range(5):  # 最多清空5行
                            if start_row + i < len(table.rows) and spec_column_index < len(table.rows[start_row + i].cells):
                                cell = table.rows[start_row + i].cells[spec_column_index]
                                if cell.paragraphs:
                                    cell.paragraphs[0].text = ""
                        
                        # 填入去重后的规格数据
                        for i in range(min(len(specifications), 5)):  # 最多填充5行
                            if start_row + i < len(table.rows) and spec_column_index < len(table.rows[start_row + i].cells):
                                cell = table.rows[start_row + i].cells[spec_column_index]
                                if cell.paragraphs:
                                    cell.paragraphs[0].text = str(specifications[i])
                                    print(f"已更新透照参数表第{start_row+i+1}行检件规格(mm×mm): {specifications[i]}")
                                    
                                    # 如果是X射线模式，则查找并填充X射线参数
                                    if ray_type == "X射线" and xray_params_df is not None:
                                        # 查找与规格匹配的X射线参数
                                        xray_params = find_xray_params_by_spec(xray_params_df, specifications[i])
                                        if xray_params:
                                            # 填充各项X射线参数
                                            param_columns = {
                                                '透照方式': 10,  # 透照方式列索引
                                                '焦距': 15,      # 焦距列索引
                                                '管电压源能量': 28, # 管电压源能量列索引
                                                '管电流源活度': 35, # 管电流源活度列索引
                                                '曝光时间': 40,   # 曝光时间列索引
                                                '有效片长': 23    # 有效片长列索引
                                            }
                                            
                                            # 填入对应参数
                                            for param_name, col_idx in param_columns.items():
                                                if param_name in xray_params and col_idx < len(table.rows[start_row + i].cells):
                                                    value = xray_params[param_name]
                                                    cell = table.rows[start_row + i].cells[col_idx]
                                                    if cell.paragraphs:
                                                        cell.paragraphs[0].text = str(value)
                                                        print(f"已更新第{start_row+i+1}行{param_name}: {value}")
                                        else:
                                            print(f"未找到规格 {specifications[i]} 的X射线参数")
                                    
                                    # 如果是γ射线模式，则查找并填充γ射线参数
                                    elif ray_type == "γ射线" and gamma_params_df is not None:
                                        # 查找与规格匹配的γ射线参数
                                        gamma_params = find_gamma_params_by_spec(gamma_params_df, specifications[i])
                                        if gamma_params:
                                            # 填充各项γ射线参数
                                            param_columns = {
                                                '透照方式': 10,  # 透照方式列索引
                                                '焦距': 15,      # 焦距列索引
                                                '管电流源活度': 35, # 管电流源活度列索引 (对应Excel中的源强)
                                                '有效片长': 23    # 有效片长列索引 (对应Excel中的一次透照长度)
                                            }
                                            
                                            # 填入对应参数
                                            for param_name, col_idx in param_columns.items():
                                                if param_name in gamma_params and col_idx < len(table.rows[start_row + i].cells):
                                                    value = gamma_params[param_name]
                                                    cell = table.rows[start_row + i].cells[col_idx]
                                                    if cell.paragraphs:
                                                        cell.paragraphs[0].text = str(value)
                                                        print(f"已更新第{start_row+i+1}行{param_name}: {value}")
                                        else:
                                            print(f"未找到规格 {specifications[i]} 的γ射线参数")
                    
                    if ray_type == "X射线":
                        print("X射线参数处理完成")
                    else:
                        print("γ射线参数处理完成")
            
            print("==== 文档填充完成 ====\n")
            
            # 保存文档
            try:
                print(f"\n正在保存文档到: {report_output_path}")
                doc.save(report_output_path)
                print(f"文档已成功保存至: {report_output_path}")
                success_count += 1
            except Exception as e:
                print(f"错误: 无法保存文档: {e}")
                error_count += 1
                
        except Exception as e:
            print(f"错误: 处理委托单编号 {order_number} 和射线类型 {ray_type} 时出错: {e}")
            error_count += 1
    
    print(f"\n处理完成: 共处理{len(groups)}个组合，成功生成{success_count}份报告，失败{error_count}份")
    if error_count > 0:
        print(f"警告: 有{error_count}个组合处理失败，请检查日志")
    
    return success_count > 0

def find_xray_params_by_spec(xray_params_df, spec_value):
    """
    根据检件规格在X射线参数表中查找对应的参数
    
    Args:
        xray_params_df: X射线参数表DataFrame
        spec_value: 检件规格值
    
    Returns:
        dict: 包含匹配到的各项参数值的字典，若未找到匹配则返回空字典
    """
    if xray_params_df is None or len(xray_params_df) == 0:
        print(f"警告: X射线参数表为空")
        return {}
    
    # 打印所有列以便调试
    print(f"X射线参数表所有列: {xray_params_df.columns.tolist()}")
    
    # 检查列名是否符合预期
    required_columns = ['检件规格', '透照方式', '焦 距（mm）', '源强（管电压）', '源活', '曝光时间', '一次透照长度(mm)']
    missing_columns = [col for col in required_columns if col not in xray_params_df.columns]
    if missing_columns:
        print(f"警告: X射线参数表缺少所需列: {missing_columns}")
        return {}
    
    # 清理规格值以便更准确匹配
    cleaned_spec = re.sub(r'[×*\s]', '', str(spec_value)).lower()
    print(f"查找规格值: '{spec_value}' (清理后: '{cleaned_spec}')")
    
    # 在检件规格列中查找匹配的行
    for idx, row in xray_params_df.iterrows():
        row_spec = str(row['检件规格'])
        cleaned_row_spec = re.sub(r'[×*\s]', '', row_spec).lower()
        
        # 如果找到完全匹配或部分匹配
        if cleaned_spec == cleaned_row_spec or cleaned_spec in cleaned_row_spec or cleaned_row_spec in cleaned_spec:
            print(f"在X射线参数表中找到匹配的规格: '{row_spec}'")
            
            # 返回匹配行的所有所需参数
            params = {
                '透照方式': row['透照方式'],
                '焦距': row['焦 距（mm）'],
                '管电压源能量': row['源强（管电压）'],
                '管电流源活度': row['源活'],
                '曝光时间': row['曝光时间'],
                '有效片长': row['一次透照长度(mm)']
            }
            
            print(f"匹配的X射线参数: {params}")
            return params
    
    print(f"警告: 在X射线参数表中未找到匹配的规格: '{spec_value}'")
    return {}

def find_gamma_params_by_spec(gamma_params_df, spec_value):
    """
    根据检件规格在γ射线参数表中查找对应的参数
    
    Args:
        gamma_params_df: γ射线参数表DataFrame
        spec_value: 检件规格值
    
    Returns:
        dict: 包含匹配到的各项参数值的字典，若未找到匹配则返回空字典
    """
    if gamma_params_df is None or len(gamma_params_df) == 0:
        print(f"警告: γ射线参数表为空")
        return {}
    
    # 打印所有列以便调试
    print(f"γ射线参数表所有列: {gamma_params_df.columns.tolist()}")
    
    # 检查列名是否符合预期
    required_columns = ['检件规格', '透照方式', '焦 距（mm）', '源强（管电压）', '一次透照长度(mm)']
    missing_columns = [col for col in required_columns if col not in gamma_params_df.columns]
    if missing_columns:
        print(f"警告: γ射线参数表缺少所需列: {missing_columns}")
        return {}
    
    # 清理规格值以便更准确匹配
    cleaned_spec = re.sub(r'[×*\s]', '', str(spec_value)).lower()
    print(f"查找规格值: '{spec_value}' (清理后: '{cleaned_spec}')")
    
    # 在检件规格列中查找匹配的行
    for idx, row in gamma_params_df.iterrows():
        row_spec = str(row['检件规格'])
        cleaned_row_spec = re.sub(r'[×*\s]', '', row_spec).lower()
        
        # 如果找到完全匹配或部分匹配
        if cleaned_spec == cleaned_row_spec or cleaned_spec in cleaned_row_spec or cleaned_row_spec in cleaned_spec:
            print(f"在γ射线参数表中找到匹配的规格: '{row_spec}'")
            
            # 返回匹配行的所有所需参数
            params = {
                '透照方式': row['透照方式'],
                '焦距': row['焦 距（mm）'],
                '管电流源活度': row['源强（管电压）'],  # 源强（管电压）列实际是源活度值
                '有效片长': row['一次透照长度(mm)']
            }
            
            print(f"匹配的γ射线参数: {params}")
            return params
    
    print(f"警告: 在γ射线参数表中未找到匹配的规格: '{spec_value}'")
    return {}

def main():
    # 创建命令行参数解析器
    parser = argparse.ArgumentParser(description='将Excel数据填入Word文档')
    parser.add_argument('-e', '--excel', default="生成器/Excel/4_生成器台账-射线检测记录.xlsx", 
                        help='Excel表格路径 (默认: 生成器/Excel/4_生成器台账-射线检测记录.xlsx)')
    parser.add_argument('-w', '--word', default="生成器/wod/4_射线检测记录.docx", 
                        help='Word模板文档路径 (默认: 生成器/wod/4_射线检测记录.docx)')
    parser.add_argument('-o', '--output', 
                        help='输出目录 (可选，默认为"生成器/输出报告/4_射线检测记录"目录)')
    
    # 新增5个命令行参数
    parser.add_argument('-p', '--project', 
                        help='工程名称，用于替换文档中的"工程名称值"')
    parser.add_argument('-u', '--entrusting_unit', 
                        help='委托单位，用于替换文档中的"委托单位值"')
    parser.add_argument('-g', '--guide_number', 
                        help='操作指导书编号，用于替换文档中的"操作指导书编号值"')
    parser.add_argument('-c', '--contracting_unit', 
                        help='承包单位，用于替换文档中的"承包单位值"')
    parser.add_argument('-m', '--equipment_model', 
                        help='设备型号，用于替换文档中的"设备型号值"')
    
    # 解析命令行参数
    args = parser.parse_args()
    
    # 处理Excel到Word的转换
    success = process_excel_to_word(
        args.excel, args.word, args.output,
        args.project, args.entrusting_unit,
        args.guide_number, args.contracting_unit,
        args.equipment_model
    )
    
    # 返回状态码
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()

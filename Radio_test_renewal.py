import pandas as pd
import os
import sys
import argparse
from docx import Document
from datetime import datetime
import re

def find_column_with_keyword(df, keyword):
    """查找包含指定关键字的列"""
    # 首先尝试精确匹配
    exact_matches = [col for col in df.columns if col.strip() == keyword.strip()]
    if exact_matches:
        return exact_matches[0]
    
    # 如果没有精确匹配，再尝试部分匹配
    partial_matches = [col for col in df.columns if keyword.lower() in col.lower()]
    return partial_matches[0] if partial_matches else None

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
    return f"{template_name}_{order_number}_{ray_mark}_续表_生成结果.docx"

def process_excel_to_word(excel_path, word_template_path, output_path=None, project_name=None, client_name=None, instruction_number=None):
    """将Excel数据填入Word文档
    
    Args:
        excel_path: Excel表格路径
        word_template_path: Word模板文档路径
        output_path: 输出目录路径（如果为None，将自动生成）
        project_name: 工程名称，用于替换Word文档中的"工程名称值"
        client_name: 委托单位，用于替换Word文档中的"委托单位值"
        instruction_number: 操作指导书编号，用于替换Word文档中的"操作指导书编号值"
    
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
        output_dir = os.path.join("生成器", "输出报告", "5_射线检测记录续")
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
    
    # 打印所有列名，帮助调试
    print(f"Excel表格列名: {list(df.columns)}")
    
    # 定义需要查找的列关键字
    column_keywords = {
        '完成日期': '完成日期',
        '委托单编号': '委托单编号',
        '检件编号': '检件编号',
        '焊口编号': '焊口编号',
        '焊工号': '焊工号',
        '规格': '规格',
        'γ射线': 'γ射线',
        '张数': '张数'  # 确保精确匹配"张数"列
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
            '规格': 'G',
            'γ射线': 'P',
            '张数': 'M'  # 明确指定张数列为M列
        }
        
        for key in missing_columns:
            col_letter = possible_columns.get(key)
            if col_letter:
                col_idx = ord(col_letter) - ord('A')
                if col_idx < len(df.columns):
                    column_mapping[key] = df.columns[col_idx]
                    print(f"使用列位置找到: '{key}' -> '{df.columns[col_idx]}'")
                else:
                    print(f"警告: 列位置 {col_letter} 超出范围，无法找到 '{key}'")
    
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
            print(f"委托单编号 {order_number} 没有明确的射线类型，处理为X射线")
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
            inspection_numbers = []
            weld_numbers = []
            welder_numbers = []
            specifications = []
            sheet_counts = []
            
            # 直接从DataFrame中获取数据
            for i in range(len(group_df)):
                if i < len(group_df):
                    # 获取当前行的各列值
                    inspection_number = group_df.iloc[i][column_mapping['检件编号']] if '检件编号' in column_mapping else ""
                    weld_number = group_df.iloc[i][column_mapping['焊口编号']] if '焊口编号' in column_mapping else ""
                    welder_number = group_df.iloc[i][column_mapping['焊工号']] if '焊工号' in column_mapping else ""
                    specification = group_df.iloc[i][column_mapping['规格']] if '规格' in column_mapping else ""
                    
                    # 添加到对应的列表
                    inspection_numbers.append(str(inspection_number))
                    weld_numbers.append(str(weld_number))
                    welder_numbers.append(str(welder_number))
                    specifications.append(str(specification))
                    
                    # 处理张数
                    if '张数' in column_mapping:
                        sheet_count_raw = group_df.iloc[i][column_mapping['张数']]
                        print(f"行 {i+1} 检件编号 {inspection_number} 原始张数值: '{sheet_count_raw}' (类型: {type(sheet_count_raw).__name__})")
                        
                        # 确保张数是数值类型
                        try:
                            if pd.isna(sheet_count_raw):
                                print(f"警告: 行 {i+1} 检件编号 {inspection_number} 的张数值为NaN，默认为1")
                                sheet_count = 1
                            else:
                                # 尝试从"180*80/3张"这种格式中提取数字
                                if isinstance(sheet_count_raw, str) and '张' in sheet_count_raw:
                                    # 使用正则表达式提取斜杠后、张字前的数字
                                    match = re.search(r'/(\d+)张', sheet_count_raw)
                                    if match:
                                        sheet_count = int(match.group(1))
                                        print(f"从字符串 '{sheet_count_raw}' 中提取到张数: {sheet_count}")
                                    else:
                                        # 如果没有找到匹配的模式，默认为1
                                        print(f"无法从 '{sheet_count_raw}' 中提取张数，默认为1")
                                        sheet_count = 1
                                else:
                                    # 尝试直接转换为整数
                                    sheet_count = int(float(sheet_count_raw))
                                    print(f"行 {i+1} 检件编号 {inspection_number} 的张数值转换为: {sheet_count}")
                        except (ValueError, TypeError) as e:
                            print(f"警告: 行 {i+1} 检件编号 {inspection_number} 的张数值 '{sheet_count_raw}' 转换失败: {e}，默认为1")
                            sheet_count = 1
                        sheet_counts.append(sheet_count)
                    else:
                        sheet_counts.append(1)  # 默认为1
            
            if '张数' in column_mapping:
                print(f"\n张数列名: '{column_mapping['张数']}'")
                print(f"张数列所有值: {group_df[column_mapping['张数']].tolist()}")
                print(f"最终获取到的张数数据: {sheet_counts}")
            else:
                print("未找到张数列，默认所有记录的张数为1")
            
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
                # 替换工程名称
                if "工程名称值" in paragraph.text and project_name:
                    paragraph.text = paragraph.text.replace("工程名称值", project_name)
                    print(f"已将段落中的'工程名称值'替换为'{project_name}'")
                # 替换委托单位
                if "委托单位值" in paragraph.text and client_name:
                    paragraph.text = paragraph.text.replace("委托单位值", client_name)
                    print(f"已将段落中的'委托单位值'替换为'{client_name}'")
                # 替换操作指导书编号
                if "操作指导书编号值" in paragraph.text and instruction_number:
                    paragraph.text = paragraph.text.replace("操作指导书编号值", instruction_number)
                    print(f"已将段落中的'操作指导书编号值'替换为'{instruction_number}'")
            
            # 遍历表格中的单元格，替换关键词
            for table in doc.tables:
                for i, row in enumerate(table.rows):
                    for j, cell in enumerate(row.cells):
                        for paragraph in cell.paragraphs:
                            if "委托单编号值" in paragraph.text:
                                paragraph.text = paragraph.text.replace("委托单编号值", committee_order)
                                print(f"已将表格单元格中的'委托单编号值'替换为'{committee_order}'")
                                replaced = True
                            # 替换工程名称
                            if "工程名称值" in paragraph.text and project_name:
                                paragraph.text = paragraph.text.replace("工程名称值", project_name)
                                print(f"已将表格单元格中的'工程名称值'替换为'{project_name}'")
                            # 替换委托单位
                            if "委托单位值" in paragraph.text and client_name:
                                paragraph.text = paragraph.text.replace("委托单位值", client_name)
                                print(f"已将表格单元格中的'委托单位值'替换为'{client_name}'")
                            # 替换操作指导书编号
                            if "操作指导书编号值" in paragraph.text and instruction_number:
                                paragraph.text = paragraph.text.replace("操作指导书编号值", instruction_number)
                                print(f"已将表格单元格中的'操作指导书编号值'替换为'{instruction_number}'")
            
            if not replaced:
                print("警告: 未找到需要替换的关键词，可能需要检查Word模板中的占位符命名。")
            
            # 填写日期（评片人、审核人）
            for table in doc.tables:
                for i, row in enumerate(table.rows):
                    for j, cell in enumerate(row.cells):
                        # 处理日期
                        date_patterns = ["评片人", "审核人"]
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
                
                # 查找包含"检件编号"、"焊缝编号"、"焊工号"等的行
                for i, row in enumerate(table.rows):
                    header_found = False
                    for j, cell in enumerate(row.cells):
                        cell_text = cell.text.strip()
                        
                        # 打印表格单元格内容，帮助调试
                        print(f"表格单元格[{i},{j}]内容: '{cell_text}'")
                        
                        if "检件编号" in cell_text:
                            column_indices["检件编号"] = j
                            header_row_index = i
                            header_found = True
                            print(f"找到检件编号列: 行 {i+1}, 列 {j+1}")
                        elif "焊缝编号" in cell_text or "焊口编号" in cell_text:
                            column_indices["焊缝编号"] = j
                            header_found = True
                            print(f"找到焊缝编号列: 行 {i+1}, 列 {j+1}")
                        elif "焊工号" in cell_text:
                            column_indices["焊工号"] = j
                            header_found = True
                            print(f"找到焊工号列: 行 {i+1}, 列 {j+1}")
                        elif "规格" in cell_text:
                            column_indices["规格"] = j
                            header_found = True
                            print(f"找到规格列: 行 {i+1}, 列 {j+1}")
                        elif "片号" in cell_text:
                            column_indices["片号"] = j
                            header_found = True
                            print(f"找到片号列: 行 {i+1}, 列 {j+1}")
                    
                    if header_found and header_row_index >= 0:
                        print(f"找到表头行: 第{header_row_index+1}行")
                        print(f"列索引: {column_indices}")
                        break
                
                # 如果没有找到某些列，尝试通过位置确定
                if "片号" not in column_indices and header_row_index >= 0:
                    # 片号通常在焊缝编号和焊工号之间，尝试通过位置确定
                    if "焊缝编号" in column_indices and "焊工号" in column_indices:
                        expected_pos = min(column_indices["焊缝编号"] + 1, column_indices["焊工号"])
                        column_indices["片号"] = expected_pos
                        print(f"通过位置推断片号列: 列 {expected_pos+1}")
                    elif len(table.rows[header_row_index].cells) > 2:
                        # 如果没有找到焊缝编号和焊工号，但表格有足够的列，假设片号在第3列
                        column_indices["片号"] = 2
                        print(f"默认片号列位置: 列 3")
                
                if "焊缝编号" not in column_indices and header_row_index >= 0:
                    # 如果没有找到焊缝编号列，但找到了检件编号列，假设焊缝编号在检件编号后一列
                    if "检件编号" in column_indices and len(table.rows[header_row_index].cells) > 1:
                        column_indices["焊缝编号"] = column_indices["检件编号"] + 1
                        print(f"通过位置推断焊缝编号列: 列 {column_indices['焊缝编号']+1}")
                
                print(f"最终确定的列索引: {column_indices}")
                
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
                            # 在最后一行之后添加一行
                            new_row = table.add_row()
                            data_rows.append(len(table.rows) - 1)  # 添加新行的索引
                    
                    # 对于张数≥6的情况，我们需要确保有足够的行来填写所有片号
                    # 检查是否需要添加额外的行来显示完整的片号序列
                    extra_rows_needed = 0
                    extra_rows_for_inspection = {}  # 记录每个检件编号需要的额外行数
                    
                    for i in range(data_count):
                        if i < len(sheet_counts):
                            # 修改逻辑：当张数为2或3时，也需要额外行
                            if sheet_counts[i] >= 6 or sheet_counts[i] in [2, 3]:
                                # 每个符合条件的检件编号需要张数个行
                                extra_needed = sheet_counts[i] - 1  # 减1是因为已经计算了一行
                                extra_rows_needed += extra_needed
                                # 记录该检件编号需要的额外行数
                                inspection_number = inspection_numbers[i]
                                extra_rows_for_inspection[inspection_number] = extra_needed
                    
                    if extra_rows_needed > 0:
                        print(f"为了显示完整的片号序列，需要额外添加{extra_rows_needed}行")
                    
                    # 处理每一行数据
                    row_index = 0  # 用于跟踪当前处理的行索引
                    processed_rows = 0  # 已处理的行数
                    
                    # 首先计算每个检件编号需要的行数
                    inspection_rows_needed = {}
                    for i in range(len(inspection_numbers)):
                        current_inspection = inspection_numbers[i]
                        current_sheet_count = sheet_counts[i] if i < len(sheet_counts) else 1
                        
                        # 对于张数≥6的情况，需要生成多行
                        # 修改逻辑：当张数为2或3时，也生成对应数量的行
                        if current_sheet_count >= 6 or current_sheet_count in [2, 3]:
                            rows_to_generate = current_sheet_count
                        else:
                            rows_to_generate = 1
                        
                        if current_inspection in inspection_rows_needed:
                            inspection_rows_needed[current_inspection] += rows_to_generate
                        else:
                            inspection_rows_needed[current_inspection] = rows_to_generate
                    
                    print(f"每个检件编号需要的行数: {inspection_rows_needed}")
                    
                    # 处理每个检件编号
                    for i in range(data_count):
                        if i < len(inspection_numbers):
                            current_inspection = inspection_numbers[i]
                            current_weld = weld_numbers[i] if i < len(weld_numbers) else ""
                            current_welder = welder_numbers[i] if i < len(welder_numbers) else ""
                            current_spec = specifications[i] if i < len(specifications) else ""
                            current_sheet_count = sheet_counts[i] if i < len(sheet_counts) else 1
                            
                            # 对于张数≥6或张数为2、3的情况，需要生成多行
                            # 修改逻辑：当张数为2或3时，也生成对应数量的行
                            if current_sheet_count >= 6 or current_sheet_count in [2, 3]:
                                rows_to_generate = current_sheet_count
                            else:
                                rows_to_generate = 1
                            
                            # 检查是否需要为当前检件编号添加行
                            if row_index + rows_to_generate > len(data_rows):
                                # 计算需要添加的行数
                                rows_needed = row_index + rows_to_generate - len(data_rows)
                                print(f"为检件编号 {current_inspection} 添加 {rows_needed} 行")
                                
                                # 在当前位置插入新行
                                for _ in range(rows_needed):
                                    # 如果当前位置有效，在当前位置之后插入新行
                                    if row_index > 0 and row_index <= len(data_rows):
                                        # 获取当前行的索引
                                        current_row_idx = data_rows[row_index - 1] if row_index - 1 < len(data_rows) else len(table.rows) - 1
                                        
                                        # 在当前行之后插入新行
                                        new_row = table.add_row()
                                        
                                        # 将新行移动到当前行之后
                                        # 注意：python-docx不直接支持在特定位置插入行，所以我们需要记录行索引
                                        new_row_idx = len(table.rows) - 1
                                        data_rows.insert(row_index, new_row_idx)
                                    else:
                                        # 如果是在末尾添加行
                                        new_row = table.add_row()
                                        data_rows.append(len(table.rows) - 1)
                            
                            # 填写当前检件编号的所有行
                            for j in range(rows_to_generate):
                                if row_index < len(data_rows):
                                    row_idx = data_rows[row_index]
                                    row = table.rows[row_idx]
                                    
                                    # 1. 填写检件编号
                                    if "检件编号" in column_indices:
                                        col_idx = column_indices["检件编号"]
                                        if col_idx < len(row.cells):
                                            cell = row.cells[col_idx]
                                            if cell.paragraphs:
                                                cell.paragraphs[0].text = str(current_inspection)
                                                print(f"已更新第{row_idx+1}行检件编号: {current_inspection}")
                                    
                                    # 2. 填写焊缝编号
                                    if "焊缝编号" in column_indices:
                                        col_idx = column_indices["焊缝编号"]
                                        if col_idx < len(row.cells):
                                            cell = row.cells[col_idx]
                                            if cell.paragraphs:
                                                # 确保单元格内容被完全替换
                                                if cell.paragraphs[0].text:
                                                    cell.paragraphs[0].text = ""
                                                cell.paragraphs[0].text = str(current_weld)
                                                print(f"已更新第{row_idx+1}行焊缝编号: {current_weld}")
                                    
                                    # 3. 填写焊工号
                                    if "焊工号" in column_indices:
                                        col_idx = column_indices["焊工号"]
                                        if col_idx < len(row.cells):
                                            cell = row.cells[col_idx]
                                            if cell.paragraphs:
                                                cell.paragraphs[0].text = str(current_welder)
                                                print(f"已更新第{row_idx+1}行焊工号: {current_welder}")
                                    
                                    # 4. 填写规格
                                    if "规格" in column_indices:
                                        col_idx = column_indices["规格"]
                                        if col_idx < len(row.cells):
                                            cell = row.cells[col_idx]
                                            if cell.paragraphs:
                                                cell.paragraphs[0].text = str(current_spec)
                                                print(f"已更新第{row_idx+1}行规格: {current_spec}")
                                    
                                    # 5. 填写片号
                                    if "片号" in column_indices:
                                        col_idx = column_indices["片号"]
                                        if col_idx < len(row.cells):
                                            cell = row.cells[col_idx]
                                            
                                            # 确定当前行是该检件编号的第几个实例
                                            current_index_in_group = j
                                            
                                            # 根据张数规则确定片号
                                            film_number = ""
                                            
                                            if current_sheet_count in [1, 4, 5]:
                                                # 不填写片号
                                                film_number = ""
                                                print(f"张数为 {current_sheet_count}，片号保持为空")
                                            elif current_sheet_count == 2:
                                                # 依次填写1，2
                                                if current_index_in_group < 2:
                                                    film_number = str(current_index_in_group + 1)
                                                    print(f"张数为 2，当前是第 {current_index_in_group + 1} 个实例，片号为: {film_number}")
                                            elif current_sheet_count == 3:
                                                # 依次填写1，2，3
                                                if current_index_in_group < 3:
                                                    film_number = str(current_index_in_group + 1)
                                                    print(f"张数为 3，当前是第 {current_index_in_group + 1} 个实例，片号为: {film_number}")
                                            elif current_sheet_count >= 6:
                                                # 依次填写1-2，2-3，3-4，...，(N-1)-N，N-1
                                                if current_index_in_group < current_sheet_count:
                                                    if current_index_in_group < current_sheet_count - 1:
                                                        # 对于前N-1个实例，填写 i-(i+1)
                                                        film_number = f"{current_index_in_group + 1}-{current_index_in_group + 2}"
                                                    else:
                                                        # 对于第N个实例，填写 N-1
                                                        film_number = f"{current_sheet_count}-1"
                                                    print(f"张数为 {current_sheet_count}，当前是第 {current_index_in_group + 1} 个实例，片号为: {film_number}")
                                            
                                            # 打印当前单元格状态
                                            print(f"片号单元格当前内容: '{cell.text}'")
                                            
                                            # 确保单元格内容被完全替换
                                            try:
                                                # 先清空单元格的所有内容
                                                for p in cell.paragraphs:
                                                    p.clear()
                                                
                                                # 如果没有段落，添加一个新段落
                                                if len(cell.paragraphs) == 0:
                                                    p = cell.add_paragraph()
                                                
                                                # 设置片号文本
                                                run = cell.paragraphs[0].add_run(film_number)
                                                
                                                if film_number:
                                                    print(f"已更新第{row_idx+1}行片号: '{film_number}'")
                                                else:
                                                    print(f"第{row_idx+1}行片号保留为空")
                                            except Exception as e:
                                                print(f"设置片号时出错: {e}")
                                                # 尝试另一种方式
                                                try:
                                                    if len(cell.paragraphs) > 0:
                                                        cell.paragraphs[0].text = film_number
                                                    else:
                                                        cell.text = film_number
                                                    print(f"使用备用方法设置片号: '{film_number}'")
                                                except Exception as e2:
                                                    print(f"备用方法也失败: {e2}")
                                    
                                    # 6. 填写像质计灵敏度
                                    # 查找表格中的"像质计灵敏度"列
                                    sensitivity_col_idx = -1
                                    for j, cell in enumerate(table.rows[header_row_index].cells):
                                        if "像质计" in cell.text and "灵敏度" in cell.text:
                                            sensitivity_col_idx = j
                                            print(f"找到像质计灵敏度列: 行 {header_row_index+1}, 列 {j+1}")
                                            break
                                    
                                    if sensitivity_col_idx >= 0 and sensitivity_col_idx < len(row.cells):
                                        # 查找对应规格的像质计灵敏度值
                                        sensitivity_value = find_sensitivity_value(current_spec, ray_type)
                                        
                                        if sensitivity_value:
                                            # 填写像质计灵敏度值
                                            cell = row.cells[sensitivity_col_idx]
                                            
                                            try:
                                                # 先清空单元格的所有内容
                                                for p in cell.paragraphs:
                                                    p.clear()
                                                
                                                # 如果没有段落，添加一个新段落
                                                if len(cell.paragraphs) == 0:
                                                    p = cell.add_paragraph()
                                                
                                                # 设置像质计灵敏度文本
                                                run = cell.paragraphs[0].add_run(sensitivity_value)
                                                print(f"已更新第{row_idx+1}行像质计灵敏度: '{sensitivity_value}'")
                                            except Exception as e:
                                                print(f"设置像质计灵敏度时出错: {e}")
                                                # 尝试另一种方式
                                                try:
                                                    if len(cell.paragraphs) > 0:
                                                        cell.paragraphs[0].text = sensitivity_value
                                                    else:
                                                        cell.text = sensitivity_value
                                                    print(f"使用备用方法设置像质计灵敏度: '{sensitivity_value}'")
                                                except Exception as e2:
                                                    print(f"备用方法也失败: {e2}")
                                    
                                    row_index += 1
                                    processed_rows += 1
            
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

# 新增函数：查找像质计灵敏度
def find_sensitivity_value(specification, ray_type):
    """
    根据规格和射线类型查找对应的像质计灵敏度值
    
    Args:
        specification: 规格值
        ray_type: 射线类型 ("X射线" 或 "γ射线")
    
    Returns:
        str: 像质计灵敏度值，如果未找到则返回空字符串
    """
    try:
        # 根据射线类型选择不同的Excel文件
        if ray_type == "X射线":
            excel_path = "生成器/Excel/4_生成器X射线指导书模版.xlsx"
            print(f"查找X射线像质计灵敏度，使用文件: {excel_path}")
        else:  # γ射线
            excel_path = "生成器/Excel/4_生成器γ射线指导书模版.xlsx"
            print(f"查找γ射线像质计灵敏度，使用文件: {excel_path}")
        
        # 检查文件是否存在
        if not os.path.exists(excel_path):
            print(f"错误: 像质计灵敏度查询文件不存在: {excel_path}")
            return ""
        
        # 读取Excel文件
        df = pd.read_excel(excel_path)
        
        # 打印列名，帮助调试
        print(f"文件 {excel_path} 的列名: {list(df.columns)}")
        
        # 检查A列是否存在
        if len(df.columns) == 0:
            print(f"错误: Excel文件 {excel_path} 没有任何列")
            return ""
        
        # 使用A列作为规格列（第一列）
        spec_column = df.columns[0]
        print(f"使用A列 '{spec_column}' 作为规格列")
        
        # 使用I列作为像质计灵敏度列（第9列，因为索引从0开始）
        if len(df.columns) <= 8:
            print(f"错误: Excel文件 {excel_path} 没有足够的列数来使用I列，当前列数: {len(df.columns)}")
            return ""
        
        sensitivity_column = df.columns[8]  # I列
        print(f"使用I列 '{sensitivity_column}' 作为像质计灵敏度列")
        
        # 清理规格字符串，便于匹配
        clean_spec = specification.strip()
        print(f"开始查找规格值: '{clean_spec}'")
        
        # 在规格列中查找匹配项
        for idx, row_spec in enumerate(df[spec_column]):
            if pd.isna(row_spec):
                continue
                
            row_spec_str = str(row_spec).strip()
            
            # 检查是否精确匹配
            if clean_spec == row_spec_str:
                # 找到匹配项，获取对应的像质计灵敏度值
                sensitivity = df.iloc[idx][sensitivity_column]
                if pd.isna(sensitivity):
                    print(f"警告: 规格 '{clean_spec}' 对应的像质计灵敏度值为空")
                    return ""
                
                print(f"找到规格 '{clean_spec}' 对应的像质计灵敏度值: {sensitivity}")
                return str(sensitivity)
        
        # 如果没有找到精确匹配，尝试部分匹配
        for idx, row_spec in enumerate(df[spec_column]):
            if pd.isna(row_spec):
                continue
                
            row_spec_str = str(row_spec).strip()
            
            # 提取规格中的数字部分
            spec_numbers = re.findall(r'\d+\.?\d*', clean_spec)
            row_spec_numbers = re.findall(r'\d+\.?\d*', row_spec_str)
            
            # 检查是否部分匹配（检查规格的数字部分是否匹配）
            if spec_numbers and row_spec_numbers and spec_numbers == row_spec_numbers:
                # 找到部分匹配项，获取对应的像质计灵敏度值
                sensitivity = df.iloc[idx][sensitivity_column]
                if pd.isna(sensitivity):
                    print(f"警告: 规格 '{row_spec_str}' (部分匹配 '{clean_spec}') 对应的像质计灵敏度值为空")
                    return ""
                
                print(f"找到规格 '{row_spec_str}' (部分匹配 '{clean_spec}') 对应的像质计灵敏度值: {sensitivity}")
                return str(sensitivity)
            
            # 尝试更宽松的匹配：只要数字部分有重叠即可
            if spec_numbers and row_spec_numbers:
                # 检查是否有共同的数字
                common_numbers = set(spec_numbers).intersection(set(row_spec_numbers))
                if common_numbers:
                    sensitivity = df.iloc[idx][sensitivity_column]
                    if pd.isna(sensitivity):
                        continue
                    
                    print(f"找到规格 '{row_spec_str}' (宽松匹配 '{clean_spec}') 对应的像质计灵敏度值: {sensitivity}")
                    return str(sensitivity)
        
        print(f"警告: 未找到规格 '{clean_spec}' 对应的像质计灵敏度值")
        return ""
    
    except Exception as e:
        print(f"查找像质计灵敏度时出错: {e}")
        return ""

def main():
    # 创建命令行参数解析器
    parser = argparse.ArgumentParser(description='将Excel数据填入Word文档')
    parser.add_argument('-e', '--excel', default="生成器/Excel/5_生成器评片记录续表模版.xlsx", 
                        help='Excel表格路径 (默认: 生成器/Excel/5_生成器评片记录续表模版.xlsx)')
    parser.add_argument('-w', '--word', default="生成器/wod/5_射线检测记录_续.docx", 
                        help='Word模板文档路径 (默认: 生成器/wod/5_射线检测记录_续.docx)')
    parser.add_argument('-o', '--output', 
                        help='输出目录 (可选，默认为"生成器/输出报告/5_射线检测记录续"目录)')
    parser.add_argument('-p', '--project', 
                        help='工程名称 (用于替换Word文档中的"工程名称值")')
    parser.add_argument('-c', '--client', 
                        help='委托单位 (用于替换Word文档中的"委托单位值")')
    parser.add_argument('-i', '--instruction', 
                        help='操作指导书编号 (用于替换Word文档中的"操作指导书编号值")')
    
    # 解析命令行参数
    args = parser.parse_args()
    
    # 处理Excel到Word的转换
    success = process_excel_to_word(args.excel, args.word, args.output, 
                                   args.project, args.client, args.instruction)
    
    # 返回状态码
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()

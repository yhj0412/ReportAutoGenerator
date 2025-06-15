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

def process_excel_to_word(excel_path, word_template_path, output_path):
    """将Excel数据填入Word文档
    
    Args:
        excel_path: Excel表格路径
        word_template_path: Word模板文档路径
        output_path: 输出Word文档路径
    
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
    
    # 确保输出目录存在
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"创建输出目录: {output_dir}")
    
    # 读取Excel数据
    print(f"正在读取Excel文件: {excel_path}")
    df = pd.read_excel(excel_path)
    
    # 打印所有列名，帮助调试
    print(f"Excel表格列名: {list(df.columns)}")
    
    # 1) 获取B列(完成日期)的最晚日期
    # 确保日期列是日期类型
    date_col = find_column_with_keyword(df, '完成日期')
    if not date_col:
        print("警告: 未找到'完成日期'列")
        return False
        
    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    latest_date = df[date_col].max()
    
    if pd.isna(latest_date):
        print("警告: 未找到有效的完成日期")
        return False
    
    # 将日期转换为年、月、日
    year = latest_date.year
    month = latest_date.month
    day = latest_date.day
    
    print(f"找到最晚完成日期: {year}年{month}月{day}日")
    
    # 2) 获取相关列的数据，使用模糊匹配
    # 定义需要查找的列关键字
    column_keywords = {
        '委托单编号': '委托单编号',
        '检件编号': '检件编号',
        '焊口编号': '焊口编号',
        '焊工号': '焊工号',
        '张数': '张数'
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
        return False
    
    # 获取所有数据列的值
    # 获取C列(委托单编号)的值
    order_numbers = df[column_mapping['委托单编号']].dropna().tolist()
    print(f"找到{len(order_numbers)}个委托单编号")
    
    # 获取D列(检件编号)的值
    test_batch_numbers = df[column_mapping['检件编号']].dropna().tolist()
    print(f"找到{len(test_batch_numbers)}个检件编号")
    
    # 获取E列(焊口编号)的值
    weld_numbers = df[column_mapping['焊口编号']].dropna().tolist()
    print(f"找到{len(weld_numbers)}个焊口编号")
    
    # 获取F列(焊工号)的值
    welder_numbers = df[column_mapping['焊工号']].dropna().tolist()
    print(f"找到{len(welder_numbers)}个焊工号")
    
    # 获取M列(张数)的值
    sheet_counts = df[column_mapping['张数']].dropna().tolist()
    print(f"找到{len(sheet_counts)}个张数")
    
    # 打开Word文档
    print(f"正在处理Word文档: {word_template_path}")
    
    # 检查文件扩展名，使用不同的方法处理.doc和.docx文件
    if word_template_path.lower().endswith('.doc'):
        # 对于.doc文件，需要先转换为.docx
        temp_docx_path = word_template_path + 'x'
        print(f"检测到.doc文件，尝试转换为.docx: {temp_docx_path}")
        
        try:
            # 尝试直接打开.doc文件
            doc = Document(word_template_path)
            doc.save(temp_docx_path)
            print(f"成功转换.doc为.docx")
            doc = Document(temp_docx_path)
        except Exception as e:
            print(f"无法直接打开.doc文件: {e}")
            print("请将.doc文件转换为.docx格式后重试")
            return False
    else:
        # 对于.docx文件，直接打开
        doc = Document(word_template_path)
    
    # 处理表格
    for table in doc.tables:
        # 查找日期字段
        for i, row in enumerate(table.rows):
            for j, cell in enumerate(row.cells):
                # 1) 处理"检测人"日期
                if "检测人" in cell.text:
                    print(f"找到检测人单元格: 第{i+1}行, 第{j+1}列")
                    
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
                            print("已更新检测人日期")
                            break
                    
                    # 如果没有找到日期段落，尝试创建新段落
                    if not date_found:
                        print("未在检测人单元格中找到日期段落，尝试其他方法...")
                        # 添加新段落
                        p = cell.add_paragraph(f"{year}年{month}月{day}日")
                        print("已添加检测人日期")
                
                # 2) 处理"审核"日期
                if "审核" in cell.text:
                    print(f"找到审核单元格: 第{i+1}行, 第{j+1}列")
                    
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
                            print("已更新审核日期")
                            break
                    
                    # 如果没有找到日期段落，尝试创建新段落
                    if not date_found:
                        print("未在审核单元格中找到日期段落，尝试其他方法...")
                        # 添加新段落
                        p = cell.add_paragraph(f"{year}年{month}月{day}日")
                        print("已添加审核日期")
        
        # 查找表头行，确定各列的位置
        column_indices = {}
        header_row_index = -1
        
        # 查找包含"委托单编号"、"检测批号"等的行
        for i, row in enumerate(table.rows):
            header_found = False
            for j, cell in enumerate(row.cells):
                if "委托单编号" in cell.text:
                    column_indices["委托单编号"] = j
                    header_row_index = i
                    header_found = True
                elif "检测批号" in cell.text:
                    column_indices["检测批号"] = j
                    header_found = True
                elif "焊口号" in cell.text:
                    column_indices["焊口号"] = j
                    header_found = True
                elif "焊工号" in cell.text:
                    column_indices["焊工号"] = j
                    header_found = True
                elif "返修张/处数" in cell.text:
                    column_indices["返修张/处数"] = j
                    header_found = True
            
            if header_found and header_row_index >= 0:
                break
        
        print(f"找到表头行: 第{header_row_index+1}行")
        print(f"列索引: {column_indices}")
        
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
            data_count = min(len(order_numbers), len(test_batch_numbers), len(weld_numbers), 
                            len(welder_numbers), len(sheet_counts))
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
                    
                    # 1. 填写委托单编号
                    if "委托单编号" in column_indices and i < len(order_numbers):
                        col_idx = column_indices["委托单编号"]
                        if col_idx < len(row.cells):
                            cell = row.cells[col_idx]
                            if cell.paragraphs:
                                cell.paragraphs[0].text = str(order_numbers[i])
                                print(f"已更新第{row_idx+1}行委托单编号: {order_numbers[i]}")
                    
                    # 2. 填写检测批号
                    if "检测批号" in column_indices and i < len(test_batch_numbers):
                        col_idx = column_indices["检测批号"]
                        if col_idx < len(row.cells):
                            cell = row.cells[col_idx]
                            if cell.paragraphs:
                                cell.paragraphs[0].text = str(test_batch_numbers[i])
                                print(f"已更新第{row_idx+1}行检测批号: {test_batch_numbers[i]}")
                    
                    # 3. 填写焊口号
                    if "焊口号" in column_indices and i < len(weld_numbers):
                        col_idx = column_indices["焊口号"]
                        if col_idx < len(row.cells):
                            cell = row.cells[col_idx]
                            if cell.paragraphs:
                                cell.paragraphs[0].text = str(weld_numbers[i])
                                print(f"已更新第{row_idx+1}行焊口号: {weld_numbers[i]}")
                    
                    # 4. 填写焊工号
                    if "焊工号" in column_indices and i < len(welder_numbers):
                        col_idx = column_indices["焊工号"]
                        if col_idx < len(row.cells):
                            cell = row.cells[col_idx]
                            if cell.paragraphs:
                                cell.paragraphs[0].text = str(welder_numbers[i])
                                print(f"已更新第{row_idx+1}行焊工号: {welder_numbers[i]}")
                    
                    # 5. 填写返修张/处数
                    if "返修张/处数" in column_indices and i < len(sheet_counts):
                        col_idx = column_indices["返修张/处数"]
                        if col_idx < len(row.cells):
                            cell = row.cells[col_idx]
                            if cell.paragraphs:
                                cell.paragraphs[0].text = str(sheet_counts[i])
                                print(f"已更新第{row_idx+1}行返修张/处数: {sheet_counts[i]}")
    
    # 保存文档
    doc.save(output_path)
    print(f"文档已保存至: {output_path}")
    return True

def main():
    # 创建命令行参数解析器
    parser = argparse.ArgumentParser(description='将Excel数据填入Word文档')
    parser.add_argument('-e', '--excel', default="生成器/Excel/2_生成器台账_结果.xlsx", 
                        help='Excel表格路径 (默认: 生成器/Excel/2_生成器台账_结果.xlsx)')
    parser.add_argument('-w', '--word', default="生成器/wod/2_新-聚乙烯结果_改.docx", 
                        help='Word模板文档路径 (默认: 生成器/wod/2_新-聚乙烯结果_改.docx)')
    parser.add_argument('-o', '--output', default="生成器/wod/生成的结果报告.docx", 
                        help='输出Word文档路径 (默认: 生成器/wod/生成的结果报告.docx)')
    
    # 解析命令行参数
    args = parser.parse_args()
    
    # 处理Excel到Word的转换
    success = process_excel_to_word(args.excel, args.word, args.output)
    
    # 返回状态码
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main() 
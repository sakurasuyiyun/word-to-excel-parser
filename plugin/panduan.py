import re
import openpyxl
import os
from docx import Document

def parse_true_false(word_file, excel_file):
    if not os.path.exists(word_file):
        print("文件路径不存在，请检查后重试。")

    """将判断题从 Word 文件转换为 Excel 格式，忽略解析内容"""
    doc = Document(word_file)
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "True or False Questions"
    
    # 设置表头
    sheet.append(["题目", "正确", "错误", "答案"])
    
    current_question = None  # 用于存储当前题目
    for paragraph in doc.paragraphs:
        line = paragraph.text.strip()

        # 如果是答案行
        if line.startswith("答案："):
            match = re.search(r"(正确|错误)", line)
            if match:
                answer = "A" if match.group(1) == "正确" else "B"
            else:
                answer = ""  # 未识别答案

            # 如果有题目，则记录进表格
            if current_question:
                sheet.append([
                    current_question,
                    "对",  # 固定列名
                    "错",  # 固定列名
                    answer
                ])
                current_question = None  # 处理完当前题目，清空

        # 如果是题目行（非答案和非解析）
        elif line and not line.startswith("答案："):
            # 检查当前是否有未处理的题目
            if current_question:  # 上一个题目未配对答案
                sheet.append([current_question, "对", "错", ""])
            
            current_question = line  # 记录新题目

    # 如果最后一题没有答案也要记录
    if current_question:
        sheet.append([current_question, "对", "错", ""])
    
    # 保存 Excel 文件
    workbook.save(excel_file)
    print(f"成功将判断题保存到 {excel_file}")
from docx import Document
import re
from datetime import datetime

def set_font(paragraph, font_name="宋体", font_size=11):
    """
    设置段落字体样式
    :param paragraph: 段落对象
    :param font_name: 字体名称，默认为宋体
    :param font_size: 字号，默认为11
    """
    for run in paragraph.runs:
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)  # 设置中文字体
        run.font.size = Pt(font_size)

def add_question_numbers_docx(input_file, output_file):
    # 打开输入文件
    doc = Document(input_file)

    # 提取文本内容
    input_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    if not input_text.strip():
        print("输入文件为空或没有有效内容！")
        return

    # 匹配题目块的正则表达式
    pattern = r"(.*?\nA\..*?\nB\..*?\nC\..*?\nD\..*?\n答案：.*?)(?=\n.*?\nA\.|\Z)"
    matches = re.findall(pattern, input_text, re.S)

    if not matches:
        print("未找到任何匹配的题目格式，请检查输入内容是否符合格式要求！")
        return

    # 为每个题目块添加序号
    numbered_questions = []
    for i, match in enumerate(matches, start=1):
        # 去掉题目块内多余的换行符
        cleaned_match = re.sub(r"\n+", "\n", match.strip())
        lines = cleaned_match.split("\n")
        lines[0] = f"{i}. {lines[0]}"  # 为题目第一行添加序号
        numbered_questions.append("\n".join(lines))

    # 写入到输出文件
    output_doc = Document()
    for question in numbered_questions:
        for line in question.split("\n"):
            para = output_doc.add_paragraph(line)
            set_font(para, font_name="宋体", font_size=11)
        output_doc.add_paragraph("")  # 添加空行分隔题目

    output_doc.save(output_file)
    print(f"已处理完成，结果保存到 {output_file}")

# 根据当前时间戳生成文件名
# timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

# 示例用法
# input_file = "input.docx"  # 输入文件名
# output_file = f"output_{timestamp}.docx"

# add_question_numbers_docx(input_file, output_file)
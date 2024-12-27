import sys
from datetime           import datetime
from plugin.choice      import multiple_choice_question
from plugin.number      import add_question_numbers_docx
from plugin.panduan     import parse_true_false
from plugin._methods    import del_file, check_number_isvalid

def switch(value, output_docx_path, output_excel_path):
    value = int(value)
    match value:
        case 1:
            multiple_choice_question(output_docx_path, output_excel_path)
            return 
        case 2:
            parse_true_false(output_docx_path, output_excel_path)
            return 
        case _:
            return False

if __name__ == "__main__":
    # 根据当前时间戳生成文件名
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    input_path = "input.docx"
    output_docx_path = f"exam_{timestamp}.docx"
    output_excel_path = f"exam_{timestamp}.xlsx"

    print("请选择题型：1-选择题 2-判断题")
    question_type = input("请输入题型编号：").strip()

    check = check_number_isvalid(2, question_type)
    if(check == False):
        print("输入有误")
        sys.exit()

    print("是否需要添加题目序号？如果题目已经有序号则无需生成序号")
    add_nums = input("[y/n]：").strip()
    if add_nums == 'y' or add_nums == 'Y':
        add_question_numbers_docx(input_path, output_docx_path, question_type)
        switch(question_type, output_docx_path, output_excel_path)
        del_file(output_docx_path)
    else:
        switch(question_type, output_docx_path, output_excel_path)
        
from pathlib import Path

def del_file(word_file):
    file_path = Path(word_file)
    if file_path.exists():  # 确保文件存在
        file_path.unlink()
    else:
        print(f"文件 {file_path} 不存在")

def check_number_isvalid(max_num, user_input):
    # 检查输入是否为非负整数
    if user_input.isdigit() == False:
        return False
    user_input = int(user_input)
    return 1 <= user_input <= max_num
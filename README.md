# word-to-excel-parser

#### 介绍
将word文档转成excel表格

#### 安装教程
1.  将代码拉到本地
2.  确保电脑已经安装Python(version >=3.10)
3.  创建虚拟隔离环境

    ```Linux/MacOS
    python3 -m venv venv
    source venv/bin/activate
4.  安装依赖

    ```python
    pip install -r requirements.txt
    ```

#### 使用说明
1.  创建一个`input.docx`文档文件在根目录下
    ###### **选择题**文档格式内容如下：
        1.xxxxx
        A.xxx
        B.xxx
        C.xxx
        D.xxx
        答案：x
    ###### **判断题**文档格式内容如下（**未实现**）：
        1.xxxxx
        正确/错误
2.  在`main.py`根目录下运行

    ```python
    python3 main.py
    ```

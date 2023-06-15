import os
import win32com.client as win32
import pandas as pd

# 创建Word应用对象
word = win32.gencache.EnsureDispatch('Word.Application')

def get_word_page_count(file_path):
    doc = word.Documents.Open(file_path)
    page_count = doc.ComputeStatistics(2)
    doc.Close(False)
    return page_count

# 指定要处理的目录
dir_path = 'test'

# 创建一个空的DataFrame来存储结果
df = pd.DataFrame(columns=['File Name', 'Page Count'])

# 遍历目录下的所有文件
for file_name in os.listdir(dir_path):
    # 只处理.docx文件
    if file_name.endswith('.docx'):
        file_path = os.path.join(dir_path, file_name)
        page_count = get_word_page_count(file_path)
        df = df.append({'File Name': file_name, 'Page Count': page_count}, ignore_index=True)

# 保存结果到Excel文件
df.to_excel('output.xlsx', index=False)

# 关闭Word应用对象
word.Quit()

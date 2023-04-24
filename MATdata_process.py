import openpyxl

# 删除多余列
file_path = "result\\result.xlsx"  # 你的Excel文件路径
sheet_name = "Sheet1"  # 你要操作的工作表名称
columns_to_delete = [2, 4, 5, 6, 7, 8, 10, 11, 12, 13, 14, 16, 17, 18, 19, 20, 26, 27, 28, 29, 30, 31, 32]  # 要删除的列索引列表（基于1的索引，例如：1表示第一列）
output_path = "result\\result-output.xlsx"

delete_columns_from_excel(file_path, output_path, sheet_name, columns_to_delete)


# 替换文字
# file_path = "result\\result-output.xlsx"  # 你的Excel文件路径
# sheet_name = "Sheet1"  # 你要操作的工作表名称
find_str = "\n质量标准\nQuality Standard"  # 要查找的字符串
replace_str = ""  # 用于替换的字符串

replace_string_in_excel(file_path, sheet_name, find_str, replace_str)

# 增加数据列

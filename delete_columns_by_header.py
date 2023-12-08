import openpyxl

def delete_columns_by_header(file_path, output_path, sheet_name, headers_to_delete):
    """
    根据Excel工作表中第一行的内容删除指定的列

    :param file_path: Excel文件的路径
    :param sheet_name: 要操作的工作表名称
    :param headers_to_delete: 一个包含要删除的列标题的列表
    """
    # 加载工作簿和工作表
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook[sheet_name]

    # 获取第一行的所有列标题
    headers = [cell.value for cell in sheet[1]]

    # 确定要删除的列索引
    columns_to_delete = [headers.index(header) + 1 for header in headers_to_delete if header in headers]

    # 按降序排列列索引，以便在删除时保持正确的索引
    columns_to_delete.sort(reverse=True)

    for col_index in columns_to_delete:
        # 删除指定的列
        sheet.delete_cols(col_index)

    # 保存更改后的工作簿
    workbook.save(output_path)

# 示例用法
file_path = "result\\result.xlsx"  # 你的Excel文件路径
sheet_name = "Sheet1"  # 你要操作的工作表名称
headers_to_delete = ["0-1-1", "0-1-2"]  # 要删除的列标题列表
output_path = "result\\result-output.xlsx"

delete_columns_by_header(file_path, output_path, sheet_name, headers_to_delete)

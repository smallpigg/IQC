import openpyxl

def delete_columns_from_excel(file_path, output_path, sheet_name, columns_to_delete):
    """
    删除Excel工作表中的指定列

    :param file_path: Excel文件的路径
    :param sheet_name: 要操作的工作表名称
    :param columns_to_delete: 一个包含要删除的列索引的列表
    """
    # 加载工作簿和工作表
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook[sheet_name]

    # 按降序排列列索引，以便在删除时保持正确的索引
    columns_to_delete.sort(reverse=True)

    for col_index in columns_to_delete:
        # 删除指定的列
        sheet.delete_cols(col_index)

    # 保存更改后的工作簿
    workbook.save(output_path)
    # print(f"删除了{len(columns_to_delete)}列")

# 示例用法
file_path = "result\\result.xlsx"  # 你的Excel文件路径
sheet_name = "Sheet1"  # 你要操作的工作表名称
columns_to_delete = [2, 4, 5, 6, 7, 8, 10, 11, 12, 13, 14, 16, 17, 18, 19, 20, 26, 27, 28, 29, 30, 31, 32]  # 要删除的列索引列表（基于1的索引，例如：1表示第一列）
output_path = "result\\result-output.xlsx"

delete_columns_from_excel(file_path, output_path, sheet_name, columns_to_delete)



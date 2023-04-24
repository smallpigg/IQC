import openpyxl

def delete_columns_with_specific_text(file_path, sheet_name, text_to_match):
    """
    删除Excel工作表中列名包含特定文本的列

    :param file_path: Excel文件的路径
    :param sheet_name: 要操作的工作表名称
    :param text_to_match: 要匹配的文本
    """
    # 加载工作簿和工作表
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook[sheet_name]

    # 获取第一行（标题行）的值
    header_row = list(sheet.iter_rows(min_row=1, max_row=1, values_only=True))[0]

    # 查找包含特定文本的列索引
    columns_to_delete = [index for index, value in enumerate(header_row, start=1) if value and text_to_match in str(value)]

    # 按降序排列列索引，以便在删除时保持正确的索引
    columns_to_delete.sort(reverse=True)

    for col_index in columns_to_delete:
        # 删除指定的列
        sheet.delete_cols(col_index)

    # 保存更改后的工作簿
    workbook.save(file_path)
    print(f"删除了{len(columns_to_delete)}列")

# 示例用法
file_path = "your_excel_file.xlsx"  # 你的Excel文件路径
sheet_name = "Sheet1"  # 你要操作的工作表名称
text_to_match = "example"  # 要匹配的文本

delete_columns_with_specific_text(file_path, sheet_name, text_to_match)

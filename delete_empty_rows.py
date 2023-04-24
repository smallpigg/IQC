from docxtpl import DocxTemplate  # pip install docxtpl
import docx

def delete_empty_rows(table, n):
    # 增加对n和表格列数的比较
    for row in table.rows:
        X_cell = row.cells[n-1]
        if X_cell.text == 'nan' or X_cell.text == '':
            row._element.getparent().remove(row._element)
            
# 示例
# file_path = "1.docx"

# checkbox_value = 1
# if checkbox_value:
#     doc = docx.Document(file_path)
#     for table in doc.tables:
#         delete_empty_rows(table, 2)
#     doc.save(file_path)
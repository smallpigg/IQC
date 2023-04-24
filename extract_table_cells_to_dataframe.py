import pandas as pd
from docx import Document

def extract_table_cells_to_dataframe(file_path, table_index):
    doc = Document(file_path)
    table = doc.tables[table_index - 1]
    data = []
    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            data.append((f"table{table_index}_{i}_{j}", cell.text))
    df = pd.DataFrame(data, columns=["column_name", "cell_content"])
    return df

# # 使用示例
# word_file = "wordfiles\AGA-MAT-001-A版 扫描枪 质量标准.docx"
# n = -1
# df = extract_table_cells_to_dataframe(word_file, n)
# print(df)

# # 将DataFrame保存为Excel文件
# df.T.to_excel('result\output.xlsx', index=False)
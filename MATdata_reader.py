import os
import docx
import pandas as pd
from pathlib import Path
import read_words_tables as rt

base_dir = Path(__file__).parent
# word_path = base_dir / "wordfiles"base_dir / "result" / "result.csv"
# path_list = os.listdir(word_path)
# docx_list = [os.path.join(word_path,str(i)) for i in path_list if str(i).endswith('docx')]
# output_path = base_dir / "result" / "result.csv"
# process_files(docx_list, output_path)

input_folder = "word_files"  # 替换为包含Word文档的文件夹路径
output_file = base_dir / "result" / "result.xlsx"  # 替换为你想要保存结果的Excel文件路径
table_numbers = [0, -1, 1]  # 替换为你想要读取的表格序号列表

rt.process_batch_word_files(input_folder, output_file, table_numbers)

















# for single_path in docx_list:
#     df1 = extract_table_cells_to_dataframe(single_path,1)
#     df2 = extract_table_cells_to_dataframe(single_path,-1)
#     df = df1 +df2
#     document = docx.Document(single_path)
#     tables = document.tables

#     filename = [os.path.basename(single_path)]
#     tablenum = len(tables)

#     table1 = tables[0]
#     cells = table1._cells
#     cells1 = tables[-1]._cells
#     cells = cells + cells1

#     cells_text = [filename + [tablenum] + [cell.text for cell in cells]]

#     df = pd.DataFrame(cells_text)
#     lfn = len(cells) + 2
#     rownum = ["filename", "numfubiao", "version", "date", "person", "changes"]
#     k = 2
#     for i in range(1, len(table1.rows)):
#         k += 1
#         for j in range(1, len(table1.columns)):
#             rownum.append(k + len(table1.columns))
#             k += 1

#     df.rename(columns={0: "filename", 1: "tablenum", lfn - 4: "version", lfn - 3: "date", lfn - 2: "person", lfn - 1: "changes"}, inplace=True)
#     pd_data.append(df[rownum])
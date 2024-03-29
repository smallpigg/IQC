import os
import docx
import pandas as pd
from pathlib import Path

"""
保存word文件内容到csv文件，其中一列数据包含了正文的表格数量
"""

base_dir = Path(__file__).parent
word_path = base_dir / "word_files"
path_list = os.listdir(word_path)
docx_list = [os.path.join(word_path,str(i)) for i in path_list if str(i).endswith('docx')]
output_path = base_dir / "result" / "result_TableNum.csv"

pd_data = []
# rownum1 = [4,5,6,7]
# rownum = [9,10,11,12,13,16,17,18,19,20,23,24,25,26,27,-4,-3,-2,-1]

# numfubiao= 0

for single_path in docx_list:
    document = docx.Document(single_path)
    tables = document.tables

    filename = list([single_path[42:]])

    numfubiao = len(tables)
    # print(filename, len(tables))
    # if len(tables)>4:
    #     numfubiao += 1
    # print("有附表的文件共有", numfubiao, "个")

    table1 = tables[0]

    cells = table1._cells
    cells1 = tables[len(tables) - 1]._cells
    cells = cells + cells1

    cells_text = [filename + list([len(tables)]) + [cell.text for cell in cells]]
    # cells_text = filename + cells_text

    # print(filename)
    # print(cells_text)

    df = pd.DataFrame(cells_text)

    lfn = len(cells)+2
    rownum = ["filename", "numfubiao", "version", "date", "person", "changes"]
    k = 2
    for i in range(1, len(table1.rows)):
        k = k + 1
        for j in range(1, len(table1.columns)):
            rownum.append(k + len(table1.columns))
            k = k + 1

    # print(rownum)

    df.rename(columns={0: "filename", 1: "numfubiao", lfn - 4: "version", lfn - 3: "date", lfn - 2: "person", lfn - 1: "changes"}, inplace=True)
    pd_data.append(df[rownum])

    # print(filename, "done!")

pd_data = pd.concat(pd_data)
pd_data.to_csv(output_path, encoding='utf_8_sig',index=False)
import docx
import pandas as pd
import os
from pathlib import Path
import numpy as np

base_dir = Path(__file__).parent
word_path = base_dir / "wordfiles"
path_list = os.listdir(word_path)
docx_list = [os.path.join(word_path,str(i)) for i in path_list if str(i).endswith('docx')]
output_path = word_path / "result.csv"

pd_data = []
# rownum1 = [4,5,6,7]
# rownum = [9,10,11,12,13,16,17,18,19,20,23,24,25,26,27,-4,-3,-2,-1]

for single_path in docx_list:
    document = docx.Document(single_path)
    tables = document.tables

    table1 = tables[0]
    cells = table1._cells

    row_count = len(table1.rows)
    col_count = len(table1.columns)

    # rownum = [lfn - 4, lfn - 3, lfn - 2, lfn - 1]
    # for i in range(1,len(table1.rows)):
    #     for j in range(2,len(table1.columns)):
    #         rownum = rownum.append(len(table1.rows)*(j-1)+2)

    print(table1.rows[0])
    print(table1.columns[0])

    print(row_count)
    print(col_count)




    cells_text = [[cell.text for cell in cells]]

    df = pd.DataFrame(cells_text)

    print(df)
import os
import docx
import pandas as pd
from pathlib import Path


def process_files(file_paths, output_path):
    pd_data = []

    for single_path in file_paths:
        document = docx.Document(single_path)
        tables = document.tables

        filename = [os.path.basename(single_path)]
        tablenum = len(tables)

        table1 = tables[0]
        cells = table1._cells
        cells1 = tables[-1]._cells
        cells = cells + cells1

        cells_text = [filename + [tablenum] + [cell.text for cell in cells]]

        df = pd.DataFrame(cells_text)
        lfn = len(cells) + 2
        rownum = ["filename", "tablenum", "version", "date", "person", "changes"]
        k = 2
        for i in range(1, len(table1.rows)):
            k += 1
            for j in range(1, len(table1.columns)):
                rownum.append(k + len(table1.columns))
                k += 1

        df.rename(columns={0: "filename", 1: "tablenum", lfn - 4: "version", lfn - 3: "date", lfn - 2: "person", lfn - 1: "changes"}, inplace=True)
        pd_data.append(df[rownum])
    print("1", pd_data)
    pd_data = pd.concat(pd_data)
    print("2", pd_data)
    pd_data.to_csv(output_path, encoding='utf_8_sig', index=False)


if __name__ == "__main__":
    base_dir = Path(__file__).parent
    word_path = base_dir / "wordfiles"
    path_list = os.listdir(word_path)
    docx_list = [os.path.join(word_path,str(i)) for i in path_list if str(i).endswith('docx')]
    output_path = base_dir / "result" / "result.csv"
    process_files(docx_list, output_path)


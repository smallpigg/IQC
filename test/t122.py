from pathlib import Path
import docx
import pandas as pd

base_dir = Path(__file__).parent
output_dir = base_dir / "OUTPUT"
excel_path = base_dir / "wordtest.xlsx"

output_dir.mkdir(exist_ok=True)
df = pd.read_excel(excel_path, sheet_name="Sheet1")

for record in df.to_dict(orient="records"):
    for i in range(1, 7):
        print(record['检验项目'+str(i)])
        list_string = [record['检验项目'+str(i)]]
        string_set = set(['材料', '产品包装', '单证资料', '规格型号', '合格证明'])
        if all([word in string_set for word in list_string]):
            print('this is ture')
        else:
            print('this is false')

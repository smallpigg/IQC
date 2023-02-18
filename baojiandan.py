from pathlib import Path
import pandas as pd
from docxtpl import DocxTemplate

base_dir = Path(__file__).parent
word_template_path1 = base_dir / "TB-SJ-QS 8.2.6-02 报检单（A版）.docx"
output_dir = base_dir / "OUTPUT"
output_dir.mkdir(exist_ok=True)

excel_path = base_dir / "bjd.xlsx"
df = pd.read_excel(excel_path, sheet_name="Sheet1")

# df["编写日期"] = pd.to_datetime(df["编写日期"]).dt.date

df["报检单文件名称"] = df["物料名称"]
for i in range(0, len(df)):
    str1 = str(df.loc[i, 'No.']) + ' ' + str(df.loc[i, '物料名称']) + '.docx'
    df.loc[i, '报检单文件名称'] = str1

for record in df.to_dict(orient="records"):
    doc = DocxTemplate(word_template_path1)
    doc.render(record)

    output_path = output_dir / f"{record['报检单文件名称']}"
    doc.save(output_path)

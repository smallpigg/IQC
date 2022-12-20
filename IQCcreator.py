from pathlib import Path

import pandas as pd  # pip install pandas openpyxl
from docxtpl import DocxTemplate  # pip install docxtpl

base_dir = Path(__file__).parent
word_template_path1 = base_dir / "电子_进货检验作业指导书-Template-dmq.docx"
word_template_path2 = base_dir / "电子_进货检验记录单-Template-common.docx"
output_dir1 = base_dir / "OUTPUT_电子"

word_template_path3 = base_dir / "机械_进货检验作业指导书-Template.docx"
word_template_path4 = base_dir / "机械_进货检验记录单-Template.docx"
output_dir2 = base_dir / "OUTPUT_机械"

excel_path = base_dir / "list.xlsx"

# Create output folder for the word documents
output_dir1.mkdir(exist_ok=True)
output_dir2.mkdir(exist_ok=True)

# Convert Excel sheet to pandas dataframe
df = pd.read_excel(excel_path, sheet_name="Sheet1")
# df2 = pd.read_excel(excel_path, sheet_name="机械")

df["编写日期"] = pd.to_datetime(df["编写日期"]).dt.date
# df2["编写日期"] = pd.to_datetime(df2["编写日期"]).dt.date


# Iterate over each row in df and render word document
for record in df.to_dict(orient="records"):
    doc = DocxTemplate(word_template_path1)
    doc.render(record)
    output_path = output_dir1 / f"{record['IQC文件名']}"
    doc.save(output_path)

    doc = DocxTemplate(word_template_path2)
    doc.render(record)
    output_path = output_dir1 / f"{record['进货检验记录单文件名']}"
    doc.save(output_path)

for record in df.to_dict(orient="records"):
    doc = DocxTemplate(word_template_path3)
    doc.render(record)
    output_path = output_dir2/ f"{record['IQC文件名']}"
    doc.save(output_path)

    doc = DocxTemplate(word_template_path4)
    doc.render(record)
    output_path = output_dir2 / f"{record['进货检验记录单文件名']}"
    doc.save(output_path)

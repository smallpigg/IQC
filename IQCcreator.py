from pathlib import Path

import pandas as pd  # pip install pandas openpyxl
from docxtpl import DocxTemplate  # pip install docxtpl

pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 100)
pd.set_option('display.width', 1000)

base_dir = Path(__file__).parent
IQC_001_path = base_dir / "template/IQC-001.docx"
TB_IQC_001_path = base_dir / "template/TB-IQC-001.docx"
TZD_AAA_B_IQC_path = base_dir / "template/TZD-AAA-B-IQC.docx"
TZD_AAA_B_TB_IQC_path = base_dir / "template/TZD-AAA-B-TB-IQC.docx"
TZD_ABA_B_IQC_path = base_dir / "template/TZD-ABA-B-IQC.docx"
TZD_ABA_B_TB_IQC_path = base_dir / "template/TZD-ABA-B-TB-IQC.docx"


output_dir = base_dir / "OUTPUT"


excel_path = base_dir / "list_V02.xlsx"

# Create output folder for the word documents
output_dir.mkdir(exist_ok=True)

# Convert Excel sheet to pandas dataframe
df = pd.read_excel(excel_path, sheet_name="Sheet1")
# df2 = pd.read_excel(excel_path, sheet_name="机械")
df["修改日期"] = pd.to_datetime(df["修改日期"]).dt.date
# df2["编写日期"] = pd.to_datetime(df2["编写日期"]).dt.date

# 增加IQC文件编号
df["IQC文件编号"] = df["质量标准编号"]
for i in range(0, len(df)):
    str1 = df.loc[i, '质量标准编号']
    str1 = str1.replace("MAT","IQC",1)
    df.loc[i, 'IQC文件编号'] = str1

# 增加物料名称
df["IQC物料名称"] = df["质量标准文件名称"]
for i in range(0, len(df)):
    str1 = df.loc[i, '质量标准文件名称']
    str1 = str1.split(" ",2)
    df.loc[i, 'IQC物料名称'] = str1[1]

# 增加IQC文件名称
df["IQC文件名称"] = df["质量标准文件名称"]
for i in range(0, len(df)):
    str1 = df.loc[i, '质量标准文件名称']
    str1 = str1.replace("MAT","IQC",1)
    str1 = str1.replace("质量标准", "进货检验作业指导书", 1)
    df.loc[i, 'IQC文件名称'] = str1

# 增加IQC文件记录文件名称
df["IQC记录文件名称"] = df["质量标准文件名称"]
for i in range(0, len(df)):
    str1 = df.loc[i, '质量标准文件名称']
    str1 = str1.replace("MAT","IQC",1)
    str1 = "TB-"+str1.replace("质量标准", "进货检验记录", 1)
    df.loc[i, 'IQC记录文件名称'] = str1

print(df)



# Iterate over each row in df and render word document
for record in df.to_dict(orient="records"):
    doc = DocxTemplate(IQC_001_path)
    doc.render(record)
    output_path = output_dir / f"{record['IQC文件名称']}"
    doc.save(output_path)

    doc = DocxTemplate(TB_IQC_001_path)
    doc.render(record)
    output_path = output_dir / f"{record['IQC记录文件名称']}"
    doc.save(output_path)

    doc = DocxTemplate(TZD_AAA_B_IQC_path)
    doc.render(record)
    output_path = output_dir / f"{record['IQC物料名称']}"
    doc.save(output_path)
#
#     doc = DocxTemplate(word_template_path2)
#     doc.render(record)
#     output_path = output_dir / f"{record['IQC记录文件名称']}"
#     doc.save(output_path)
#
#
# TB_IQC_JX_001_path = base_dir / "template/TB-IQC-JX-001.docx"
# TZD_AAA_B_IQC_path = base_dir / "template/TZD-AAA-B-IQC.docx"
# TZD_AAA_B_TB_IQC__path = base_dir / "template/TZD-AAA-B-TB-IQC.docx"
# TZD_ABA_B_IQC_path = base_dir / "template/TZD-ABA-B-IQC.docx"
# TZD_ABA_B_TB_IQC_path = base_dir / "template/TZD-ABA-B-TB-IQC.docx"
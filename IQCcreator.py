from pathlib import Path
from docxtpl import DocxTemplate  # pip install docxtpl
import docx
import pandas as pd
# import os
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.text import WD_TAB_ALIGNMENT
# from docx.enum.style import WD_STYLE_TYPE
# from docx.enum.table import WD_TABLE_ALIGNMENT
# from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
# from docx.shared import Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
def set_cell_border(cell, **kwargs):
    """
    Set cell`s border
    Usage:
    set_cell_border(
        cell,
        top={"sz": 12, "val": "single", "color": "#FF0000", "space": "0"},
        bottom={"sz": 12, "color": "#00FF00", "val": "single"},
        left={"sz": 24, "val": "dashed", "shadow": "true"},
        right={"sz": 12, "val": "dashed"},
    )
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # check for tag existnace, if none found, then create one
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)

    # list over all available tags
    for edge in ('left', 'top', 'right', 'bottom', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)

            # check for tag existnace, if none found, then create one
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)

            # looks like order of attributes is important
            for key in ["sz", "val", "color", "space", "shadow"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))

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

df["申请日期"] = pd.to_datetime(df["申请日期"]).dt.date
# df["编写日期"] = pd.to_datetime(df["编写日期"]).dt.date
df["A版日期"] = pd.to_datetime(df["A版日期"]).dt.date

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

# 增加IQC文件记录通知单文件名称
df["IQC文件通知单名称"] = df["质量标准文件名称"]
for i in range(0, len(df)):
    str1 = df.loc[i, '质量标准文件名称']
    str1 = str1.replace("MAT","IQC",1)
    str1 = str1.replace("质量标准", "进货检验作业指导书文件记录更改通知单", 1)
    df.loc[i, 'IQC文件通知单名称'] = str1

# 增加TBIQC文件记录通知单文件名称
df["IQC记录文件通知单名称"] = df["质量标准文件名称"]
for i in range(0, len(df)):
    str1 = df.loc[i, '质量标准文件名称']
    str1 = str1.replace("MAT","IQC",1)
    str1 = "TB-"+str1.replace("质量标准", "进货检验记录文件记录更改通知单", 1)
    df.loc[i, 'IQC记录文件通知单名称'] = str1

IQC_001_path = 'C:\\Users\\Zz\PycharmProjects\\IQC\\template\\IQC-001.docx'
# Iterate over each row in df and render word document
# IQC_TB_path = 'C:\\Users\\Zz\PycharmProjects\\IQC\\template\\IQC-TB-001.docx'



for record in df.to_dict(orient="records"):
    i = 2
    while i <= 7:
        if str(record['检验项目'+str(i+1)]) == "nan":
            str1 = '00'+str(i-1)
            IQC_path = IQC_001_path.replace("001",str1,1)
            doc = DocxTemplate(IQC_path)
            doc.render(record)
            output_path = output_dir / f"{record['IQC文件名称']}"
            doc.save(output_path)
            break
        else:
            i += 1



    doc = DocxTemplate(TB_IQC_001_path)
    doc.render(record)
    output_path = output_dir / f"{record['IQC记录文件名称']}"
    doc.save(output_path)

    if record['质量标准编号'][:3] == 'AAA':
        doc = DocxTemplate(TZD_AAA_B_IQC_path)
        doc.render(record)
        output_path = output_dir / f"{record['IQC文件通知单名称']}"
        doc.save(output_path)
        doc = DocxTemplate(TZD_AAA_B_TB_IQC_path)
        doc.render(record)
        output_path = output_dir / f"{record['IQC记录文件通知单名称']}"
        doc.save(output_path)
        # print('AAA', record['质量标准编号'][:3])
    else:
        doc = DocxTemplate(TZD_ABA_B_IQC_path)
        doc.render(record)
        output_path = output_dir / f"{record['IQC文件通知单名称']}"
        doc.save(output_path)
        doc = DocxTemplate(TZD_ABA_B_TB_IQC_path)
        doc.render(record)
        output_path = output_dir / f"{record['IQC记录文件通知单名称']}"
        doc.save(output_path)
        # print('ABA', record['质量标准编号'][:3])
    print(record['IQC文件编号']," is done!")
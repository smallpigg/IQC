import os
import shutil
from docx import Document
from docx.shared import Inches
import pandas as pd
from PIL import Image

def save_images_from_docx(filepath, output_folder):
    doc = Document(filepath)
    images = []
    
    for i, shape in enumerate(doc.inline_shapes):
        if shape.type == 3:  # Type 3 indicates an image
            image_path = os.path.join(output_folder, f"{os.path.splitext(os.path.basename(filepath))[0]}_{i+1}.png")
            with open(image_path, "wb") as image_file:
                image_file.write(shape._inline.graphic.graphicData.pic.blipFill.blip.blob)
            images.append(image_path)
    
    return images

def process_batch_word_files(input_folder, output_file, table_numbers):
    dfs = []
    image_folder = "images"
    
    if not os.path.exists(image_folder):
        os.makedirs(image_folder)

    for filename in os.listdir(input_folder):
        if filename.endswith(".docx"):
            filepath = os.path.join(input_folder, filename)
            df = read_word_tables_to_df(filepath, table_numbers)
            
            # Save images from the document and add their filenames to the DataFrame
            image_paths = save_images_from_docx(filepath, image_folder)
            for i, image_path in enumerate(image_paths):
                df[f"图片{i+1}"] = os.path.basename(image_path)
            
            dfs.append(df)

    result_df = pd.concat(dfs, ignore_index=True)
    result_df.to_excel(output_file, index=False)

# 示例用法
# input_folder = "word_files"  # 替换为包含Word文档的文件夹路径
# output_file = "output.xlsx"  # 替换为你想要保存结果的Excel文件路径
# table_numbers = [1, 3]  # 替换为你想要读取的表格序号列表

# process_batch_word_files(input_folder, output_file, table_numbers)

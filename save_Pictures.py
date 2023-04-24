import docx
import os, re
from pathlib import Path

def get_pictures(word_path, result_path):
    """
    图片提取
    :param word_path: word路径
    :param result_path: 结果路径
    :return:
    """
    doc = docx.Document(word_path)
    dict_rel = doc.part._rels
    for rel in dict_rel:
        rel = dict_rel[rel]
        if "image" in rel.target_ref:
            if not os.path.exists(result_path):
                os.makedirs(result_path)
            img_name = re.findall("/(.*)", rel.target_ref)[0]
            word_name = os.path.splitext(word_path)[0]
            if os.sep in word_name:
                new_name = word_name.split('\\')[-1]
            else:
                new_name = word_name.split('/')[-1]
            img_name = f'{new_name}_{img_name}'
            with open(f'{result_path}/{img_name}', "wb") as f:
                f.write(rel.target_part.blob)


# 示例用法
base_dir = Path(__file__).parent
word_path = base_dir / "word_files"
images_path = base_dir / "images"
# path_list = os.listdir(word_path)
# docx_list = [os.path.join(word_path,str(i)) for i in path_list if str(i).endswith('docx')]
# output_path = word_path / "images"

os.chdir(word_path)
spam=os.listdir(os.getcwd())
for i in spam:
    get_pictures(str(i), images_path)
    # print(str(i), images_path)


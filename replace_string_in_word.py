from docx import Document

def replace_string_in_word(source_file, target_file, old_string, new_string):
    doc = Document(source_file)

    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if old_string in run.text:
                run.text = run.text.replace(old_string, new_string)

    doc.save(target_file)
    
# 示例
replace_string_in_word('test/source.docx', 'test/target.docx', '\n', 'nihao')
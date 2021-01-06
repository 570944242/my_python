# coding=utf-8

from docx import Document

'''
自用修改尚硅谷docx文档的代码
old_doc,new_doc,old_str,new_str
旧文件的路径,新文件路径,要替换的字符串,替换成新的字符串
'''


def replace_doc(old_doc, new_doc, old_str, new_str):
    document = Document(old_doc)
    document.settings.odd_and_even_pages_header_footer = False
    for section in document.sections:
        section.different_first_page_header_footer = False
        section.header.is_linked_to_previous = True
        section.footer.is_linked_to_previous = True
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            if "尚硅谷" in run.text:
                run.text = run.text.replace(old_str, new_str)
                print(run.text)
    document.save(new_doc)


replace_doc('input/尚硅谷大数据.docx', 'input/尚硅谷大数据tmp.docx', '尚硅谷', '宋词')

# -*- coding: utf-8 -*-
from docx import Document

class DocxHelper:
    def __init__(self, docx_path, make_copy = False):
        try:
            self.doc = Document(docx_path)
        except Exception as e:
            print(str(e))

        self.doc_path = docx_path
        self.make_copy = make_copy

    def save(self):
        if self.make_copy:
            formatter = "{}_Copy.docx"
            str_path = str(self.doc_path)
            start = str_path.find(".docx")
            self.doc.save(formatter.format(str_path[0: int(start)]))
        else:    
            self.doc.save(self.doc_path)

    def is_string_exists(self, find_text):
        for paragraph in self.doc.paragraphs:
            if find_text in paragraph.text:
                return True
        
        for table in self.doc.tables:
            for row in table.rows:  
                for cell in row.cells:
                    if find_text in cell.text:
                        return True
                    for paragraph in cell.paragraphs:
                        if find_text in paragraph.text:
                            return True
        return False      
    
    def find_replace(self, find_text, replace_text):
        for paragraph in self.doc.paragraphs:
            self.replace_in_paragrah(paragraph, find_text, replace_text)

        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    self.find_replace_cell(cell, find_text, replace_text)

    def replace_in_paragrah(self, paragraph, search_text, replace_text):
        full_text = paragraph.text
        if search_text in full_text:
            # 开始替换逻辑
            start_index = full_text.find(search_text)
            end_index = start_index + len(search_text)
            current_index = 0
            text_replaced = False

            for run in paragraph.runs:
                run_end_index = current_index + len(run.text)

                if text_replaced or run_end_index < start_index:
                    # 如果这个run完全在替换文本的前面，或者文本已经被替换，跳过
                    current_index += len(run.text)
                    continue

                # 计算run中需要替换文本的部分
                run_start_index = max(start_index - current_index, 0)
                run_end_index = min(end_index - current_index, len(run.text))

                # 替换文本
                run.text = run.text[:run_start_index] + replace_text + run.text[run_end_index:]
                
                text_replaced = True
                current_index += len(run.text)              

    def find_paragrahs(self, find_text):
        res = []
        for paragraph in self.doc.paragraphs:
            if find_text in paragraph.text:
                res.append(paragraph.text)

        for table in self.doc.tables:
            for row in table.rows:  
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if find_text in paragraph.text:
                            res.append(paragraph.text)
        return res
    
    def async_find_replace(self, find_text, replace_text):
        self.find_replace(find_text, replace_text)
                    
    def find_replace_cell(self, cell, find_text, replace_text):
        for paragraph in cell.paragraphs:
            self.replace_in_paragrah(paragraph, find_text, replace_text)
    

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
            for run in paragraph.runs:
                if find_text in run.text:
                    return True
        
        for table in self.doc.tables:
            for row in table.rows:  
                for cell in row.cells:
                    if find_text in cell.text:
                        return True
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            if find_text in run.text:
                                return True
        return False      
    
    def find_replace(self, find_text, replace_text):
        for paragraph in self.doc.paragraphs:
            for run in paragraph.runs:
                if find_text in run.text:
                    run.text = run.text.replace(find_text, replace_text)

        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    self.find_replace_cell(cell, find_text, replace_text)

    def find_paragrahs(self, find_text):
        res = []
        for paragraph in self.doc.paragraphs:
            for run in paragraph.runs:
                if find_text in run.text:
                    res.append(paragraph.text)
                    break 

        for table in self.doc.tables:
            for row in table.rows:  
                for cell in row.cells:
                    # if find_text in cell.text:
                    #     res.append(cell.text)
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            if find_text in run.text:
                                res.append(paragraph.text)
                                break       
        return res
    
    def async_find_replace(self, find_text, replace_text):
        self.find_replace(find_text, replace_text)
                    
    def find_replace_cell(self, cell, find_text, replace_text):
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                if find_text in run.text:
                    run.text = run.text.replace(find_text, replace_text)
    
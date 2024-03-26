
# -*- coding: utf-8 -*-
from PyQt6.QtCore import QObject, pyqtSignal
from PyQt6.QtCore import QRunnable

class WorkerSignals(QObject):
     finished = pyqtSignal(int, bool)
     docx_search_finish=pyqtSignal(list)


class Worker(QRunnable):
    def __init__(self, docxHelper, index, search_text):
        super(Worker, self).__init__()
        self.docxHelper = docxHelper
        self.search_text = search_text
        self.num = index
        self.signals = WorkerSignals()

    def run(self):
        res = self.docxHelper.is_string_exists(self.search_text)
        self.signals.finished.emit(self.num, res)  # 任务完成时发出信号

class DocxSearchWorker(QRunnable):
    def __init__(self, docxHelper, search_text):
        super(DocxSearchWorker, self).__init__()
        self.docxHelper = docxHelper
        self.search_text = search_text
        self.signals = WorkerSignals()
    
    def run(self):
        paragraphs = self.docxHelper.find_paragrahs(self.search_text)
        self.signals.docx_search_finish.emit(paragraphs)
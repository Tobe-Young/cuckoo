
from PyQt6.QtCore import QObject, pyqtSignal
from PyQt6.QtCore import QRunnable
import pypandoc
import os

class WorkerSignals(QObject):
    convert_finish = pyqtSignal(bool)

class Doc2PdfHelper(QRunnable):
    def __init__(self, src_dir, dest_dir):
        super().__init__()
        self.src_dir = src_dir
        self.dest_dir = dest_dir
        self.signals = WorkerSignals()

    def run(self):
        for doc_path in self.src_dir.rglob('*'):
            if doc_path.suffix.lower() in ['.docx', '.doc']:
                relative_path = doc_path.relative_to(self.src_dir)
                dest_path = self.dest_dir / relative_path.with_suffix('.pdf')
                dest_path.parent.mkdir(parents=True, exist_ok=True)
                self.convert_docx_to_pdf(doc_path, dest_path)
        
        print("All conversions completed.")       
        self.signals.convert_finish.emit(True)

    def convert_docx_to_pdf(self, src_file, dest_file):
        pypandoc.convert_file(str(src_file), 'pdf', outputfile=str(dest_file))
        
           
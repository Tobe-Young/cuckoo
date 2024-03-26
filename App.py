# -*- coding: utf-8 -*-
import sys
from pathlib import Path
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QApplication, QLabel, QWidget, QMainWindow,  QSplitter, QListWidget, QTextEdit, QHBoxLayout, QWidget
from PyQt6.QtWidgets import QLineEdit, QVBoxLayout,QPushButton, QSpacerItem, QSizePolicy, QFileDialog
from PyQt6.QtWidgets import QListWidgetItem, QCheckBox, QMessageBox, QStatusBar
from PyQt6.QtGui import QIcon
from PyQt6.QtCore import QThreadPool
from PyQt6.QtCore import QSettings
from docx import Document
import DocxUtils
import Icon
from Runnables import Worker, DocxSearchWorker

class ChaTyWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("")
        self.setWindowTitle('Docx查替')
        self.setGeometry(100, 100, 1000, 800)
        self.const_working_dir = "当前工作目录：{}"
        self.filtered_files_list = []
        self.make_file_copy = False
        self.docxHelper = None
        self.find_text = ""
        self.replace_text = ""
        self.docx_dict = {}
        self.is_select_all = False
        self.is_searching_selected_file = False
        self.is_quering_files = False
        # 线程池设置为最多4个线程
        self.threadPool = QThreadPool()
        self.threadPool.setMaxThreadCount(8)
        self.settings = QSettings('DoubleInc', 'ChaTy')
        self.initUI()
        last_work_dir = self.settings.value("last_work_dir", "")
        if not last_work_dir == "":
            self.selected_dir = last_work_dir
            self.search_selected_dir()
            
    def initUI(self):
        # 创建主窗口的中心部件，并设置布局
        centralWidget = QWidget()
        self.setCentralWidget(centralWidget)
        mainLayout = QVBoxLayout(centralWidget)
        mainLayout.setSpacing(0)
        mainLayout.setContentsMargins(0, 0, 0, 0)
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)

        # 创建包含两个输入框的容器
        containerHWidget = QWidget()
        containerHWidget.setFixedHeight(30)
        containerHLayout = QHBoxLayout(containerHWidget)
        containerHLayout.setContentsMargins(12, 0, 12, 0)
        btnSelectDir = QPushButton("选择文件夹")
        btnSelectDir.clicked.connect(self.on_select_dir_clicked)
        containerHLayout.addWidget(btnSelectDir)
        self.workingDirLbl = QLabel(self.const_working_dir.format(""))
        self.workingDirLbl.setStyleSheet('QLabel {max-width: 800px;}')
        containerHLayout.addWidget(self.workingDirLbl)
        self.save_as_copy_checkbox = QCheckBox("保存为新文件")
        self.save_as_copy_checkbox.setChecked(self.make_file_copy)
        self.save_as_copy_checkbox.stateChanged.connect(self.on_checkbox_state_changed)
        containerHLayout.addItem(QSpacerItem(20, 30, QSizePolicy.Policy.Expanding))
        containerHLayout.addWidget(self.save_as_copy_checkbox)
        
        inputContainer = QWidget()
        inputContainer.setFixedHeight(30)
        inputLayout = QHBoxLayout(inputContainer)
        inputLayout.setSpacing(0)
        inputLayout.setContentsMargins(12, 0, 12, 0)
        findTextLabel = QLabel("查找内容：")
        replaceTextLbl = QLabel("替换内容：")
        self.input1 = QLineEdit()
        self.input1.textChanged.connect(self.on_find_text_changed)
        startSearchBtn = QPushButton("开始查找")
        startSearchBtn.clicked.connect(self.on_start_search)
        self.input2 = QLineEdit()
        self.input2.textChanged.connect(self.on_replace_text_changed)
        btnConfirm = QPushButton("开始替换")
        btnConfirm.clicked.connect(self.on_start_replacing)
        inputLayout.addWidget(findTextLabel)
        inputLayout.addWidget(self.input1)
        inputLayout.addWidget(startSearchBtn)
        inputLayout.addWidget(replaceTextLbl)
        inputLayout.addWidget(self.input2)
        inputLayout.addWidget(btnConfirm)

        # 创建水平分割器
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # 左侧文件列表
        leftWidget = QWidget()
        leftLayout = QVBoxLayout(leftWidget)
        leftLayout.setContentsMargins(12, 0, 0, 0)
        self.selectAllCheckBox = QCheckBox("选中所有")
        self.selectAllCheckBox.stateChanged.connect(self.on_select_all_checkbox_state_changed)
        self.fileListWidget = QListWidget()
        # 文件列表选中变化时连接到槽函数
        self.fileListWidget.currentTextChanged.connect(self.on_file_selected)
        leftLayout.addWidget(self.selectAllCheckBox)
        leftLayout.addWidget(self.fileListWidget)
        
        # 右侧文件预览
        self.filePreview = QTextEdit()
        self.filePreview.setReadOnly(True)
        
        # 添加到分割器
        self.splitter.addWidget(leftWidget)
        self.splitter.addWidget(self.filePreview)

        # 将输入容器和分割器添加到主布局
        mainLayout.addWidget(containerHWidget)
        mainLayout.addWidget(inputContainer)
        mainLayout.addWidget(self.splitter)

        totalWidth = self.splitter.width()
        leftWidth = int(totalWidth * 0.25)
        rightWidth = int(totalWidth - leftWidth)
        self.splitter.setSizes([leftWidth, rightWidth])

        self.center_on_screen()

    def center_on_screen(self):
        screen = self.screen().geometry()
        size = self.geometry()
        x = (screen.width() - size.width()) / 2
        y = (screen.height() - size.height()) / 2
        self.move(int(x), int(y))

    def on_select_all_checkbox_state_changed(self, state):
        if self.is_select_all == (state == 2) :
            return
        self.is_select_all = not self.is_select_all
        state = Qt.CheckState.Checked if self.is_select_all else Qt.CheckState.Unchecked
        list_count = self.fileListWidget.count()
        for i in range(list_count):
            list_item = self.fileListWidget.item(i)
            list_item.setCheckState(state)

    def on_checkbox_state_changed(self, state):
        if state == 2:
            self.make_file_copy = True
        else:
            self.make_file_copy = False

    def on_find_text_changed(self, text):
        self.find_text = text
 
    def get_docx_helper(self, str_path):
        if not str_path in self.docx_dict:
            helper = DocxUtils.DocxHelper(str_path, self.make_file_copy)
            self.docx_dict[str_path] = helper
            return helper
        return self.docx_dict[str_path]
            
    def on_replace_text_changed(self, text):
        self.replace_text = text

    def on_filed_finished(self, paragraphs):
        html = ""
        for paragraph in paragraphs:
            html += "<p>"
            html += paragraph.replace(self.find_text, f"<b>{self.find_text}</b>")
            html += "</p>"
        self.filePreview.setHtml(html)
        self.is_searching_selected_file = False

    def on_file_selected(self, fileName):
        # 这里可以添加实际的文件读取逻辑
        # 为了简化示例，我们只是在文本编辑器中显示所选文件名
        if self.is_searching_selected_file:
            return
        
        self.is_searching_selected_file = True

        index = self.fileListWidget.currentRow()
        file_path = self.filtered_files_list[index]
        self.statusBar.showMessage(f"当前文件路径{file_path}")
        self.filePreview.setHtml("检索当前选中文件内容中......")
        if self.find_text is None or self.find_text == "":
            return
        
        docxHelper = self.get_docx_helper(str(file_path))
        worker = DocxSearchWorker(docxHelper, self.find_text)
        worker.signals.docx_search_finish.connect(self.on_filed_finished)
        self.threadPool.start(worker)

    def on_select_dir_clicked(self):
        self.selected_dir = QFileDialog.getExistingDirectory()
        self.settings.setValue("last_work_dir", str(self.selected_dir))
        self.search_selected_dir()
    
    def search_selected_dir(self):
        self.workingDirLbl.setText(self.const_working_dir.format(str(self.selected_dir)))
        self.docx_files = list(ChaTyWindow.search_docx_in_dir(self.selected_dir))
        self.filtered_files_list = self.docx_files.copy()
        self.refresh_listwidget()
    
    def check_search_result(self, index, has_found):
        self.search_count += 1
        if has_found :
            docx = self.docx_files[index]
            self.filtered_files_list.append(docx)
            listItem = QListWidgetItem(f'{docx.name}')
            listItem.setFlags(listItem.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            state = Qt.CheckState.Checked if self.is_select_all else Qt.CheckState.Unchecked
            listItem.setCheckState(state) # 设置初始状态为未选中
            self.fileListWidget.addItem(listItem)

        if self.search_count == len(self.docx_files):
            if self.fileListWidget.count() > 0:
                self.fileListWidget.setCurrentRow(0)         
                self.on_file_selected("")
            self.is_quering_files = False

    def on_start_search(self):
        if not hasattr(self, 'docx_files'):
            QMessageBox.information(self, "警告", "请选择工作目录")
            return
        if self.is_quering_files:
            self.statusBar.showMessage(f"搜索当前目录中....")
            return
        
        self.is_quering_files = True
        text = self.find_text
        self.filePreview.setHtml("")
        self.fileListWidget.clear()
        self.filtered_files_list = []
        self.search_count = 0

        state = Qt.CheckState.Checked if self.is_select_all else Qt.CheckState.Unchecked
        if text == "":
            self.filtered_files_list = self.docx_files.copy()
            self.filePreview.setHtml("没有选择文件")
            for docx in self.filtered_files_list:
                listItem = QListWidgetItem(f'{docx.name}')
                listItem.setFlags(listItem.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                listItem.setCheckState(state) # 设置初始状态为未选中
                self.fileListWidget.addItem(listItem)    
        else:
            count = len(self.docx_files)
            for i in range(count):
                docx = self.docx_files[i]
                docxHelper = self.get_docx_helper(str(docx))
                worker = Worker(docxHelper, i, text)
                worker.signals.finished.connect(self.check_search_result)
                self.threadPool.start(worker)      

    def on_start_replacing(self):
        self.docxHelper = None
        toReplaceFiles = []
        for index in range(len(self.filtered_files_list)):
            fileItem = self.fileListWidget.item(index)
            if fileItem.checkState() == Qt.CheckState.Checked:
                toReplaceFiles.append(self.filtered_files_list[index])

        for docx in toReplaceFiles:
            docxHelper = self.get_docx_helper(docx)
            docxHelper.make_copy = self.make_file_copy
            docxHelper.async_find_replace(self.find_text, self.replace_text)
            docxHelper.save()
        QMessageBox.information(self, "替换", "替换完成")

        self.refresh_listwidget()
        
    def refresh_listwidget(self):
        self.docx_files = list(ChaTyWindow.search_docx_in_dir(self.selected_dir))
        self.filePreview.setHtml("没有选择文件")
        self.fileListWidget.clear()
        for file in self.docx_files:
             listItem = QListWidgetItem(f'{file.name}')
             listItem.setFlags(listItem.flags() | Qt.ItemFlag.ItemIsUserCheckable)
             listItem.setCheckState(Qt.CheckState.Unchecked) # 设置初始状态为未选中
             self.fileListWidget.addItem(listItem)

    def search_docx_in_dir(workingDir):
        path = Path(workingDir)
        if not path.exists():
            return []
        docx_files = path.rglob('*.docx')
        res = []
        for file in docx_files:
            if not file.name.startswith('~$'):
                res.append(file)
        return res

def main():
     app = QApplication(sys.argv)
     app.setWindowIcon(QIcon(Icon.get_icon()))
     chaTy = ChaTyWindow()
     chaTy.show()
     sys.exit(app.exec())

if __name__ == "__main__":
    main()
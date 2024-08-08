import sys, os, traceback

from PySide6.QtCore import *
from PySide6.QtWidgets import * #QApplication, QMainWindow, QPushButton, QFileDialog
from PySide6.QtGui import *
from pprint import pprint

from modules import pdf_utils

if sys.platform == "linux" or sys.platform == "linux2":
    pass

elif sys.platform == "win32":
    import ctypes
    myappid = u'mycompany.myproduct.subproduct.version'  # arbitrary string
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

elif sys.platform == "darwin":
    pass

# Subclass QMainWindow to customize your application's main window
class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()

        self.setWindowTitle("PDF Utilities")
        self.setMinimumSize(QSize(300,100))
        self.setIcon()
        self.layout = QGridLayout()

        self.btn_select = QPushButton("Select Files")
        self.btn_select.clicked.connect(self.selectFiles)
        
        self.btn_merge = QPushButton("Merge Files")
        self.btn_merge.clicked.connect(self.start_merge_pdf_thread)
        
        self.btn_convert = QPushButton("Convert")
        self.btn_convert.clicked.connect(self.convert_pdf_thread)
        
        self.txt_edit = QTextEdit()
        self.txt_edit.setReadOnly(True)
        
        self.layout.addWidget(self.btn_select,0,0,1,2)
        self.layout.addWidget(self.txt_edit,1,0,1,2)
        self.layout.addWidget(self.btn_merge,2,0)
        self.layout.addWidget(self.btn_convert,2,1)

        self.btn_convert.setDisabled(True)
        self.btn_merge.setDisabled(True)
        
        widget = QWidget()
        widget.setLayout(self.layout)
        self.setCentralWidget(widget)

        self.list_selected_files = None

    def setIcon(self):
        appIcon = QIcon()
        appIcon.addFile("resources/icon.png")
        self.setWindowIcon(appIcon)

    def log(self, msg, color="black", end="<br>"):
        self.txt_edit.textCursor().insertHtml("<font color='{}'>{}</font>{}".format(color,msg,end))

    def selectFiles(self):
        self.clear()
        self.btn_merge.setDisabled(True)
        self.btn_convert.setDisabled(True)
        
        filename = QFileDialog()
        filename.setFileMode(QFileDialog.ExistingFiles)
        names = filename.getOpenFileNames(self, "Select PDF Files", ".", "PDF (*.pdf)")
        self.list_selected_files = names[0]
        names = [os.path.basename(x) for x in self.list_selected_files]
        names.sort()
        self.log("<br>".join(names))

        if len(self.list_selected_files) > 0:
            if len(self.list_selected_files) > 1:
                self.btn_merge.setDisabled(False)
            self.btn_convert.setDisabled(False)

    def start_merge_pdf_thread(self): 
        self.btn_merge.setDisabled(True)
        
        if len(self.list_selected_files) < 2:
            self.log("Please select atleast 2 files for merging!", "red")

        else:    
            self.thread = pdf_utils.mergePDFsThread(self.list_selected_files, "merged.pdf")
            self.thread.log.connect(self.log)
            self.thread.thread_finished.connect(self.done_merge_pdf_thread)
            self.thread.start()
    
    def done_merge_pdf_thread(self):
        self.btn_merge.setDisabled(False)
        self.btn_convert.setDisabled(False)

    def convert_pdf_thread(self):
        self.thread = pdf_utils.pdfToDocxThread(self.list_selected_files)
        self.thread.msg.connect(self.log)
        self.thread.start()

    def clear(self):
        self.txt_edit.clear()

app = QApplication(sys.argv)
app.setStyle("Fusion")
window = MainWindow()
window.show()
app.exec()
sys.exit()
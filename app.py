import sys, os

from PySide6.QtCore import QSize, Qt
from PySide6.QtWidgets import * #QApplication, QMainWindow, QPushButton, QFileDialog
from PySide6.QtGui import *
from pprint import pprint

from modules import pdf_utils

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
        self.btn_merge.clicked.connect(self.mergeFiles)
        
        self.btn_convert = QPushButton("Convert")
        self.btn_convert.clicked.connect(self.convertFiles)
        
        self.txt_edit = QTextEdit()
        self.txt_edit.setReadOnly(True)
        
        self.layout.addWidget(self.btn_select,0,0,1,2)
        self.layout.addWidget(self.txt_edit,1,0,1,2)
        self.layout.addWidget(self.btn_merge,2,0)
        self.layout.addWidget(self.btn_convert,2,1)
        
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
        filename = QFileDialog()
        filename.setFileMode(QFileDialog.ExistingFiles)
        names = filename.getOpenFileNames(self, "Select PDF Files", ".", "PDF (*.pdf)")
        self.list_selected_files = names[0]
        names = [os.path.basename(x) for x in self.list_selected_files]
        names.sort()
        self.log("<br>".join(names))

    def mergeFiles(self):
        output = pdf_utils.mergePDFs(self.list_selected_files, "merged.pdf")
        self.log(f"Merged PDF created: {output}", color="blue")

    def convertFiles(self):
        pdf_utils.pdfToDocx(self.list_selected_files)
        self.log("File(s) converted.", color="blue")

app = QApplication(sys.argv)
window = MainWindow()
window.show()
app.exec()
sys.exit()
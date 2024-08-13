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
        
        self.txt_edit = QTextEdit()
        self.txt_edit.setReadOnly(True)

        self.btn_merge = QPushButton("Merge Files")
        self.btn_merge.clicked.connect(self.start_merge_pdf_thread)
        
        self.btn_convert = QPushButton("Convert")
        self.btn_convert.clicked.connect(self.convert_pdf_thread)
        
        self.btn_ocr = QPushButton("OCR")
        self.btn_ocr.clicked.connect(self.start_ocr_thread)
        
        self.layout.addWidget(self.btn_select,0,0,1,2)
        self.layout.addWidget(self.txt_edit,1,0,1,2)
        self.layout.addWidget(self.btn_merge,2,0)
        self.layout.addWidget(self.btn_convert,2,1)
        self.layout.addWidget(self.btn_ocr,3,0)

        self.btn_ocr.setDisabled(True)
        self.btn_convert.setDisabled(True)
        self.btn_merge.setDisabled(True)
        
        self.central_widget = QWidget()
        self.central_widget.setLayout(self.layout)
        self.setCentralWidget(self.central_widget)

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
            self.btn_ocr.setDisabled(False)

    def start_merge_pdf_thread(self): 
        self.btn_merge.setDisabled(True)
        
        if len(self.list_selected_files) < 2:
            self.log("Please select atleast 2 files for merging!", "red")

        else:    
            self.thread = pdf_utils.mergePDFsThread(self.list_selected_files, "merged.pdf")
            self.thread.log.connect(self.log)
            self.thread.thread_finished.connect(self.done_merge_pdf_thread)
            self.thread.start()
            self.hide()
            self.loading_screen()
    
    def done_merge_pdf_thread(self):
        self.btn_merge.setDisabled(False)
        self.btn_convert.setDisabled(False)
        self.loading_screen_window.close()
        self.show()

    def convert_pdf_thread(self):
        self.thread = pdf_utils.pdfToDocxThread(self.list_selected_files)
        self.thread.msg.connect(self.log)
        self.thread.start()

    def start_ocr_thread(self):
        self.btn_convert.setDisabled(True)
        self.btn_ocr.setDisabled(True)
        self.thread = pdf_utils.ocrPdfThread(self.list_selected_files)
        self.thread.msg.connect(self.log)
        self.thread.thread_finished.connect(self.done_ocr_thread)
        self.thread.start()        
        self.hide()
        self.loading_screen()

    def done_ocr_thread(self):
        self.loading_screen_window.close()
        self.show()
        self.btn_convert.setDisabled(False)
        self.btn_ocr.setDisabled(False)
        self.log("OCR task(s) completed!")

    def loading_screen(self):
        self.loading_screen_window = QWidget()
        self.loading_screen_window.setWindowIcon(QIcon("resources/icon.png"))
        self.loading_screen_window.setWindowTitle("Processing")
        self.loading_screen_window.setStyleSheet("background-color: white;")
        self.loading_screenlayout = QVBoxLayout()
        self.loading_screen_label = QLabel()
        self.loading_screen_label.setGeometry(QRect(25, 25, 200, 200)) 
        self.loading_screen_label.setMinimumSize(QSize(250, 250)) 
        self.loading_screen_label.setMaximumSize(QSize(250, 250)) 
        self.loading_screenlayout.addWidget(self.loading_screen_label)
        movie = QMovie("resources/loading.gif")
        self.loading_screen_label.setMovie(movie)
        self.loading_screen_label.setScaledContents(True)
        movie.start()
        self.loading_screen_window.setLayout(self.loading_screenlayout)
        self.loading_screen_window.show()
        self.loading_screen_window.exec()
        self.disable_all()

    def disable_all(self):
        self.btn_merge.setDisabled(True)
        self.btn_convert.setDisabled(True)
        self.btn_ocr.setDisabled(True)
        self.btn_select.setDisabled(True)

    def clear(self):
        self.txt_edit.clear()

if __name__ == "__main__":    
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    app.exec()
    sys.exit()
from PyPDF2 import PdfMerger
from pdf2docx import Converter
from pdf2docx import parse
import os
import time
import ocrmypdf

from PySide6.QtWidgets import *
from PySide6.QtCore import *
from PySide6.QtGui import *

class mergePDFsThread(QThread):

    log = Signal(str,str)
    thread_finished = Signal(int)

    def __init__(self, list_pdfs, outname_filename):
        super().__init__()
        self.list_pdfs        = list_pdfs
        self.outname_filename = outname_filename

    def run(self):
        merger = PdfMerger()

        for pdf in self.list_pdfs:
            merger.append(open(pdf, "rb"))

        if os.path.isfile(self.outname_filename):
            suffix = str(int(time.time()))
            self.outname_filename = self.outname_filename.replace(".", "_"+suffix+".")

        with open(self.outname_filename, "wb") as fout:
            merger.write(fout)

        self.log.emit("Merging Files..", "blue")
        time.sleep(0.2)
        self.log.emit(f"merged PDF: {self.outname_filename}", "blue")
        print("mergePDFsThread finished!")
        self.thread_finished.emit(1)

    def stop(self):
        self.is_running = False
        self.terminate()

class pdfToDocxThread(QThread):
    
    msg = Signal(str,str)

    def __init__(self, list_pdfs):
        super().__init__()
        self.list_pdfs = list_pdfs
    
    def run(self):
        if len(self.list_pdfs) > 0:
            for i in  self.list_pdfs:
                parse(i, i+".docx", start=0, end=None)
        
        self.msg.emit("Conversion Completed!","blue")

# class loadingAnimationThread(QThread):

#     def __init__(self):
#         super().__init__()

#     def run(self):


class ocrPdfThread(QThread):
    msg = Signal(str,str)
    thread_finished = Signal(int)

    def __init__(self, list_of_files):
        super().__init__()
        self.list_of_files = list_of_files
    
    def run(self):
        for file in self.list_of_files:
            suffix = str(int(time.time()))
            output_file = file.replace(".", "_"+suffix+"_ocr.")

            ocrmypdf.ocr(file, output_file, language=["eng","hin"])
            self.msg.emit(f"OCR Completed for {file}", "blue")
        self.thread_finished.emit(1)

def pdfToDocx(list_pdfs):
    if len(list_pdfs) > 0:
        for i in  list_pdfs:
            parse(i, i+".docx", start=0, end=None)
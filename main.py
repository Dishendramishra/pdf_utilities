from PyPDF2 import PdfMerger
from pdf2docx import Converter
from pdf2docx import parse
import glob
import os
import time

def mergePDFs(list_pdfs, outname_filename):
    merger = PdfMerger()
    for pdf in list_pdfs:
        merger.append(open(pdf, "rb"))

    if os.path.isfile(outname_filename):
        suffix = str(int(time.time()))
        outname_filename = outname_filename.replace(".", "_"+suffix+".")

    with open(outname_filename, "wb") as fout:
        merger.write(fout)

def pdfToDocx(pdf_filename, docx_filename):
    parse(pdf_filename, docx_filename, start=0, end=None)


# if __name__ == "__main__":
# mergePDFs(glob.glob("*.pdf"),"merged.pdf")
# pdfToDocx("page2.pdf", "page2.docx")
from PyPDF2 import PdfFileMerger, PdfFileReader
import os

merger = PdfFileMerger()

mPath = r"C:\Users\ALESSANDROAlves\Box\Abbott Latam\Leadership\Demand Management Bucket\ME-003273"

file1 = os.path.join(mPath,"ME-003273 - Page 1-10.pdf")
file2 = os.path.join(mPath,"ME-003273 - Pag 11.pdf")
file3 = os.path.join(mPath,"ME-003273 - Page 12.pdf")

filenames = [file1, file2, file3]


for filename in filenames:
    merger.append(PdfFileReader(open(filename, 'rb')))

merger.write("document-output.pdf")
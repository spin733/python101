import os
import re
from tkinter import filedialog
import win32com.client

#작업할 대상 폴더 선택
foldername = filedialog.askdirectory()
foldername = foldername.replace('/', '\\')

#엑셀파일 필터
files = [f for f in os.listdir(foldername) if re.match(r'.*\.xls', f)]
excel = win32com.client.Dispatch("Excel.Application")
#엑셀을 PDF 파일로 저장
for file in files:
    wb = excel.Workbooks.Open(os.path.join(foldername,file))
    pre,ext = os.path.splitext(file)
    rfolder = os.getcwd() + "\\Result"
    wb.ExportAsFixedFormat(0, os.path.join(rfolder,pre + ".pdf"))
    wb.Close()
excel.Quit()
#

# PPT파일 필터
files = [f for f in os.listdir(foldername) if re.match(r'.*[.]ppt', f)]
ppttoPDF = 32
powerpoint = win32com.client.Dispatch("Powerpoint.Application")
#PPT PDF로 전환
for file in files:
    deck = powerpoint.Presentations.Open(os.path.join(foldername,file))
    pre, ext = os.path.splitext(file)
    rfolder = os.getcwd() + "\\Result"
    deck.SaveAs(os.path.join(rfolder,pre + ".pdf"), ppttoPDF)  # formatType = 32 for ppt to pdf
    deck.Close()
powerpoint.Quit()

from fpdf import FPDF
pdf = FPDF('L')

# 이미지파일 필터
files = [f for f in os.listdir(foldername) if re.match(r'.*[.](jpg|png|gif)', f)]
#이미지 PDF파일로 저장
for file in files:
    pdf.add_page()
    pdf.image(os.path.join(foldername,file), 0, 0, 297, 210)
rfolder = os.getcwd() + "\\Result\\IMG2PDF.pdf"
pdf.output(rfolder, "F")

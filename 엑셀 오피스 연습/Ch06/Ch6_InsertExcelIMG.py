from tkinter import filedialog
import os
import openpyxl

#엑셀 워크북 & 워크시트 설정해주기
wb = openpyxl.load_workbook('Ch6_Sample.xlsx')
ws = wb.worksheets[0]


#폴더 선택하고 파일명 리스트 가져오기
foldername = filedialog.askdirectory()
filearr = os.listdir(foldername)
i=1
for file in filearr:
    # 이미지 삽입하기
    img = openpyxl.drawing.image.Image(foldername + '/' + file)
    img.anchor = 'A' + str(i)
    IMGRation = img.width / img.height
    img.width = 250
    img.height = 250 / IMGRation
    ws.add_image(img)
    ws.row_dimensions[i].height = img.height * 0.75
    i=i+1

wb.save('Ch6_Sample_Result.xlsx')
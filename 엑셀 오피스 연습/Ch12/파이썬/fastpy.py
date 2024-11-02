import os
import re
import win32com.client

def AutoPDF (TPath,RPath) :
    #작업할 대상 폴더 선택
    TPath = TPath.replace('/', '\\')
    RPath = RPath.replace('/', '\\')
    
    #엑셀파일 필터
    files = [f for f in os.listdir(TPath) if re.match(r'.*\.xls', f)]
    excel = win32com.client.Dispatch("Excel.Application")
    #엑셀을 PDF 파일로 저장
    for file in files:
        wb = excel.Workbooks.Open(os.path.join(TPath,file))
        pre,ext = os.path.splitext(file)
        wb.ExportAsFixedFormat(0, os.path.join(RPath,pre + ".pdf"))
        wb.Close()
    excel.Quit()
    #
    
    # PPT파일 필터
    files = [f for f in os.listdir(TPath) if re.match(r'.*[.]ppt', f)]
    ppttoPDF = 32
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    #PPT PDF로 전환
    for file in files:
        deck = powerpoint.Presentations.Open(os.path.join(TPath,file))
        pre, ext = os.path.splitext(file)
        deck.SaveAs(os.path.join(RPath,pre + ".pdf"), ppttoPDF)  # formatType = 32 for ppt to pdf
        deck.Close()
    powerpoint.Quit()
    
    from fpdf import FPDF
    pdf = FPDF('L')
    
    # 이미지파일 필터
    files = [f for f in os.listdir(TPath) if re.match(r'.*[.](jpg|png|gif)', f)]
    #이미지 PDF파일로 저장
    for file in files:
        pdf.add_page()
        pdf.image(os.path.join(TPath,file), 0, 0, 297, 210)
    pdf.output(RPath + "\\IMG2PDF.pdf", "F")

def AutoFont(TPath, RPath, EFont, HFont) :
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")

    files = [f for f in os.listdir(TPath) if re.match(r'.*[.]ppt', f)]

    for file in files:
        TFile = os.path.join(TPath,file)
        ppt= powerpoint.Presentations.Open(TFile)

        for sld in ppt.Slides :
            for shp in sld.shapes :
                #Text Frame이 있을 때
                if shp.HasTextFrame == -1 :
                    shp.TextFrame.TextRange.Font.Name = EFont
                    shp.TextFrame.TextRange.Font.NameFarEast = HFont
                #Table 이 있을 때
                if shp.HasTable == -1 :
                    for col in shp.Table.Columns:
                        for cell in col.cells:
                            cell.Shape.TextFrame.TextRange.Font.Name = EFont
                            cell.Shape.TextFrame.TextRange.Font.NameFarEast = HFont
                #Group 안에 있을때
                try:
                    for GI in shp.GroupItems:
                        if GI.HasTextFrame == -1 :
                            GI.TextFrame.TextRange.Font.Name = EFont
                            GI.TextFrame.TextRange.Font.NameFarEast = HFont
                except:
                    pass
        RFile = os.path.join(RPath,file)
        ppt.SaveAs(RFile)
        ppt.Close()

import win32com.client

powerpoint = win32com.client.Dispatch("Powerpoint.Application")
TPath = "D:\\Python\\FastCampus\\Ch11_FontSample.pptx"
RPath = "D:\\Python\\FastCampus\\Ch11_FontResult.pptx"

EFont = "Arial"
HFont = "궁서체"

ppt= powerpoint.Presentations.Open(TPath)

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
ppt.SaveAs(RPath)
ppt.Close()

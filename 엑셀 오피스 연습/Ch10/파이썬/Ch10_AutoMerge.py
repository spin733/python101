import pandas as pd

df_ref = pd.read_excel('Ch10_MergeMaster.xlsx')
df_merge = pd.read_excel('Ch10_Merge_Result.xlsx')

from tkinter import filedialog
import os
import re

foldername = filedialog.askdirectory()
files = [f for f in os.listdir(foldername) if re.match(r'.*\.xls', f)]

#엑셀을 PDF 파일로 저장
for file in files:
    df_work = pd.read_excel(os.path.join(foldername,file))
    for index, row in df_ref.iterrows():
        df_work.rename(columns={row[1]:row[0]}, inplace=True) #df_work 칼럼 이름을 replace
    df_frame = pd.DataFrame(df_work, columns=df_merge.columns.values)
    df_merge = pd.concat([df_merge,df_frame])

df_merge.to_excel("Ch10_FinalResult.xlsx", index = False)
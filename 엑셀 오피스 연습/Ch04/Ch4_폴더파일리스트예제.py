from tkinter import filedialog

folder_selected = filedialog.askdirectory()

import os

arr = os.listdir(folder_selected)

for filename in arr:
    print (filename)
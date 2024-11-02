from tkinter import *
from tkinter import font
import os
from tkinter import filedialog
import fastpy

def TargetPathUp() :
    foldername = filedialog.askdirectory()
    tl.config(text=foldername)

def ResultPathUp() :
    foldername = filedialog.askdirectory()
    rl.config(text=foldername)

def hCurFont (evt) :
    txt = hselect.get(hselect.curselection())
    hfl.config(text=txt)

def eCurFont (evt) :
    txt = eselect.get(eselect.curselection())
    efl.config(text=txt)

def RunAutoPDF () :
    fastpy.AutoPDF(tl.cget("text"),rl.cget("text"))

def RunAutoFont () :
    fastpy.AutoFont(tl.cget("text"),rl.cget("text"),efl.cget("text"),hfl.cget("text"))

# def whichSelected () :
#     print "At %s of %d" % (select.curselection(), len(phonelist))
#     return int(select.curselection()[0])

def makeWindow () :
    global tl, rl, hfl, efl, hselect, eselect
    win = Tk()
    win.title('Auto GUI')
    fonts = list(font.families())

    frame1 = Frame(win,padx=10,pady=10)
    frame1.pack()

    Label(frame1, text="Target Path : ", width=10).grid(row=0, column=0, sticky=W)
    tl = Label(frame1, text=os.getcwd() + "\\Target",relief="groove", width=40)
    tl.grid(row=0, column=1, sticky=W)
    Button(frame1, text="Select Path", width=10, command= TargetPathUp).grid(row=0, column=2, sticky=W)

    Label(frame1, text="Result Path : ", width=10).grid(row=1, column=0, sticky=W)
    rl = Label(frame1, text=os.getcwd() + "\\Result", relief="groove", width=40)
    rl.grid(row=1, column=1, sticky=W)
    Button(frame1, text="Select Path", width=10, command= ResultPathUp).grid(row=1, column=2, sticky=W)

    frame2 = Frame(win,padx=10,pady=10)  # Row of buttons
    frame2.pack()
    Button(frame2,text="Auto Font", width=20, command=RunAutoFont).grid(row=0,column=0,sticky=W)
    Label(frame2,width=10).grid(row=0,column=1,sticky=W)
    Button(frame2, text="Auto PDF", width=20, command=RunAutoPDF).grid(row=0, column=2, sticky=W)

    frame3 = Frame(win,padx=10)
    frame3.pack()
    Label(frame3, text="선택된 한글 폰트 : ", width=15).grid(row=0, column=0, sticky=W)
    hfl = Label(frame3, text="Spoqa Han Sans Neo Medium", width=35, anchor="w")
    hfl.grid(row=0, column=1, sticky=W)

    frame4 = Frame(win,padx=10,pady=10)       # select of names
    frame4.pack()
    hscroll = Scrollbar(frame4, orient=VERTICAL)
    hselect = Listbox(frame4, yscrollcommand=hscroll.set, height=6, width=50)
    hselect.bind('<<ListboxSelect>>', hCurFont)
    hscroll.config (command=hselect.yview)
    hscroll.pack(side=RIGHT, fill=Y)
    hselect.pack(side=LEFT,  fill=BOTH, expand=1)
    for item in fonts:
        hselect.insert (END, item)

    frame5 = Frame(win,padx=10)
    frame5.pack()
    Label(frame5, text="선택된 영문 폰트 : ", width=15).grid(row=0, column=0, sticky=W)
    efl = Label(frame5, text="Spoqa Han Sans Neo Medium", width=35, anchor="w")
    efl.grid(row=0, column=1, sticky=W)

    frame6 = Frame(win,padx=10,pady=10)  # select of names
    frame6.pack()
    escroll = Scrollbar(frame6, orient=VERTICAL)
    eselect = Listbox(frame6, yscrollcommand=escroll.set, height=6, width=50)
    eselect.bind('<<ListboxSelect>>', eCurFont)
    escroll.config(command=eselect.yview)
    escroll.pack(side=RIGHT, fill=Y)
    eselect.pack(side=LEFT, fill=BOTH, expand=1)
    for item in fonts:
        eselect.insert (END, item)

    return win

win = makeWindow()
win.resizable(0, 0)
win.mainloop()
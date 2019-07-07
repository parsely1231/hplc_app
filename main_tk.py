import os
from tkinter import *
from tkinter import ttk
from tkinter import filedialog


filepath = ''


def fileselect_button_clicked():  # 変換するファイルを選ぶ関数（ボタンを押した時に実行）
    ftype = [('text file', '*.txt')]
    idir = os.path.abspath(os.path.dirname(__file__))
    global filepath
    filepath = filedialog.askopenfilename(filetypes=ftype, initialdir=idir)
    filename.set(filepath)


def convert_button_clicked():  # 選んだファイルを読み込んでエクセルファイルに変換する（ボタンを押した時に実行）
    import openfile
    import create_xlsx


if __name__ == '__main__':
    root = Tk()
    root.title(u'HPLC txt to csv')
    root.geometry('600x400')

    static1 = ttk.Label(root, text='変換元のテキストファイルを選んでください')
    static1.pack()

    filename = StringVar()
    filename_entry = ttk.Entry(root, textvariable=filename, width=50)
    filename_entry.pack()

    fileselect_button = ttk.Button(root, text=u'file select', width=10, command=fileselect_button_clicked)
    fileselect_button.pack()

    convert_button = ttk.Button(root, text=u'convert', width=10, command=convert_button_clicked)
    convert_button.pack()

    root.mainloop()

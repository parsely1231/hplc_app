import os
from tkinter import *
from tkinter import ttk
from tkinter import filedialog


def fileselect_button_clicked():
    ftype = [('text file', '*.txt')]
    idir = os.path.abspath(os.path.dirname(__file__))
    filepath = filedialog.askopenfilename(filetypes = ftype, initialdir = idir)
    filename.set(filepath)


def convert_button_clicked():
    import convert


    wb = excel.Workbook()
    ws = wb.active

    row_num = 2
    y = row_num

    column_num = 2
    x = column_num

    for sample in samplelist:
        ws.cell(y, x, sample)
        y += 1
        ws.cell(y, x, 'RT')
        ws.cell(y, x + 1, 'Area')
        ws.cell(y, x + 2, 'Area%')
        # サンプル名をセルに記入し、一行下にRT、Area、Area%を記入

        for datalist in datatable_dict[sample]:
            # サンプル名に対応した２次元配列からリストを取り出し
            y += 1
            # 一行下に下げて
            for data in datalist:
                ws.cell(y, x, data)
                x += 1
                # リストから値を取り出して、列を右にずらしながらセルに記入
            l = len(datalist)
            x -= l
            # for data から抜けたら、ずらした分の列を左に戻す
        x += 5
        y = 2
        # for datalistから抜けたら列を右に十分にずらして、行を元（二行目）に戻す

    wb.save('sample.xlsx')


if __name__ == '__main__':

    root = Tk()
    root.title(u'HPLC txt to csv')
    root.geometry('600x400')

    static1 = ttk.Label(root, text = '変換元のテキストファイルを選んでください')
    static1.pack()

    filename = StringVar()
    filename_entry = ttk.Entry(root, textvariable = filename, width = 50)
    filename_entry.pack()

    fileselect_button = ttk.Button(root, text=u'file select', width = 10, command = fileselect_button_clicked)
    fileselect_button.pack()

    convert_button = ttk.Button(root, text = u'convert', width = 10, command = convert_button_clicked)
    convert_button.pack()

    root.mainloop()

import os
import openpyxl as excel
from tkinter import *
from tkinter import ttk
from tkinter import filedialog


datatable = []
# 2次元配列用の空リスト

datatable_dict = {}
# 2次元配列を複数入れるための辞書

samplelist = []
# サンプルネームの空リスト

wb = excel.Workbook()
ws = wb.active

filepath = ''


def fileselect_button_clicked():  # 変換するファイルを選ぶ関数（ボタンを押した時に実行）
    ftype = [('text file', '*.txt')]
    idir = os.path.abspath(os.path.dirname(__file__))
    global filepath
    filepath = filedialog.askopenfilename(filetypes=ftype, initialdir=idir)
    filename.set(filepath)


def readfile():
    with open(filepath, 'r') as file_open:
        n = -1

        for line in file_open:
            if line[0] == '#':
                line = line.replace('\n', '')
                global samplelist
                samplelist.append(line)
                n += 1
                continue

            elif line[0] in ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0']:
                str_list = line.split()
                float_list = [float(s) for s in str_list]
                global datatable
                datatable.append(float_list)
                # 読み込んだ一行のデータを、floatリストにする
                # その後２次元配列用の空リストに入れる
                continue

            datatable_name = samplelist[n]
            global datatable_dict
            datatable_dict[datatable_name] = datatable
            datatable = []
            # 直近のサンプル名をkeyにして作成中の２次元配列をvalueにする。
            # その後２次元配列の中身をからにする


def convert():
    row_num = 2
    y = row_num

    column_num = 2
    x = column_num

    for sample in samplelist:
        ws.cell(y, x, sample)
        y += 1
        ws.cell(y, x, 'name')
        ws.cell(y, x+1, 'RT')
        ws.cell(y, x+2, 'RRT')
        ws.cell(y, x+3, 'Area')
        ws.cell(y, x+4, 'Area%')
        ws.cell(y, x+5, '補正Area%')
        # サンプル名をセルに記入し、一行下にname, RT, RRT, Area, Area%, 補正Area%を記入

        rt_list = [datalist[0] for datalist in datatable_dict[sample]]  # ２次元配列の各リストの１つ目の要素をリスト化（RT）
        area_list = [datalist[1] for datalist in datatable_dict[sample]]  # ２次元配列の各リストの2つ目の要素をリスト化（RT）
        areaper_list = [datalist[2] for datalist in datatable_dict[sample]]  # ２次元配列の各リストの3つ目の要素をリスト化（RT）

        x += 1  # １列右にずらしてnameの下からRTの下に移動
        for rt_data in rt_list:
            # rt_listからrt_dataを取り出し
            y += 1
            ws.cell(y, x, rt_data)
            # 行を下にずらしながらセルに記入

        y += 2
        ws.cell(y, x-1, '基準RT')

        std_rt_x = x
        std_rt_y = y
        # std_rtのセル位置をあとで使う用の変数

        y = 4  # 3行目に戻る(RT等のカラムタイトルが書かれている行の１個下)
        x += 1  # 1列右にずらしてRTの下からRRTの下に移動

        length = len(rt_list)
        while y < 4+length:
            ws.cell(y, x,
                    '='
                    + ws.cell(y, x-1).coordinate
                    + '/'
                    + ws.cell(std_rt_y, std_rt_x).coordinate)
            y += 1
            # RTをstd_RTで割ってRRTを計算させる

        y = 3
        x += 1  # 1列右にずらしてRRTの下からAreaの下に移動

        for area_data in area_list:
            # area_listからarea_dataを取り出し
            y += 1
            ws.cell(y, x, area_data)
            # 行を下にずらしながらセルに記入

        y += 2
        ws.cell(y, x,
                '=SUM('
                + ws.cell(4, x).coordinate
                + ':'
                + ws.cell(y-2, x).coordinate
                + ')')
        # SUM関数を文字列としてエクセルに記入,total areaを計算させる。

        ws.cell(y, x-1, 'total Area')
        totalarea_x = x
        totalarea_y = y
        # totalareaのセル位置をあとで使う用の変数

        y = 3
        x += 1  # 1列右にずらしてAreaの下からArea%の下に移動

        for areaper_data in areaper_list:
            # areaper_listからareaper_dataを取り出し
            y += 1
            ws.cell(y, x, areaper_data)
            # 行を下にずらしながらセルに記入

        y = 4
        x += 1  # 1列右にずらしてAreaの下からArea%の下に移動

        while y < 4+length:
            ws.cell(y, x,
                    '='
                    + ws.cell(y, x-2).coordinate
                    + '/'
                    + ws.cell(totalarea_y, totalarea_x).coordinate)
            y += 1

        x += 2
        y = 2
        # fordatalistから抜けたら列を右に2列ずらして、行を元（二行目）に戻す


def convert_button_clicked():  # 選んだファイルを読み込んでエクセルファイルに変換する（ボタンを押した時に実行）
    readfile()
    convert()
    root.savefile = filedialog.asksaveasfilename(initialdir="/", title="Save as", filetypes=[("xlsx file", "*.xlsx")])
    wb.save(root.savefile)


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

import openpyxl as excel
from openfile import samplelist
from openfile import datatable_dict

wb = excel.Workbook()
ws = wb.active

row_no = 2
y = row_no

colmun_no = 2
x = colmun_no

for sample in samplelist:
    ws.cell(y, x, sample)
    y += 1
    ws.cell(y, x, 'RT')
    ws.cell(y, x+1, 'Area')
    ws.cell(y, x+2, 'Area%')
    #サンプル名をセルに記入し、一行下にRT、Area、Area%を記入

    for datalist in datatable_dict[sample]:
        #サンプル名に対応した２次元配列からリストを取り出し
        y += 1
        #一行下に下げて
        for data in datalist:
            ws.cell(y, x, data)
            x += 1
            #リストから値を取り出して、列を右にずらしながらセルに記入
        l = len(datalist)
        x -= l
        #for data から抜けたら、ずらした分の列を左に戻す
    x += 5
    y = 2
    #for datalistから抜けたら列を右に十分にずらして、行を元（二行目）に戻す

ws2 = wb.create_sheet('sheet2')


wb.save('sample.xlsx')


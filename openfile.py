with open('rawdata.txt', 'r') as file_open:
    #後で開くファイルをGUIで選択するようにする

    datatable = []
    #2次元配列用の空リスト

    datatable_list = {}
    #2次元配列を複数入れるための辞書

    samplelist = []
    #サンプルネームの空リスト

    x = -1

    for line in file_open:
        if line[0] == '#':
            line = line.replace('\n', '')
            samplelist.append(line)
            x += 1
            continue

        elif line[0] in ['1','2','3','4','5','6','7','8','9','0']:
            str_list = line.split()
            float_list = [float(s) for s in str_list]
            datatable.append(float_list)
            # 読み込んだ一行のデータを、floatリストにする
            # その後２次元配列用の空リストに入れる
            continue

        datatable_name = samplelist[x]
        datatable_list[datatable_name] = datatable
        datatable.clear()
        #直近のサンプル名をkeyにして作成中の２次元配列をvalueにする。
        #その後２次元配列の中身をからにする
print(datatable_list)
print("finish")
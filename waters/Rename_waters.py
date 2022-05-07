#########################################################
# watersから出力されるtxtとpdfファイルが対象 
#########################################################

print('ライブラリをロード中...\n')
import os
from glob import glob
import pandas as pd

#====================================
# テキストファイル内で先頭が#の列数を返す
def return_sharp_nrows(txt):
    file=open(txt,"r",encoding="shift_jis")  #shift_jis形式だった
    text=file.read()
    text_split = text.split("\n")
    nrows = 0
    for l in text_split:
        if l.startswith("#"):
            sharp_nrow = nrows
        nrows += 1
    return sharp_nrow

#====================================
# フォルダ内からデータリストを作る関数
# data_dicのデータ構造は
# { {file1.txt: {"df1":df1(サンプル名,分析日,注入量,装置メソッドなど), "df2":df2(rt,area,area%など)}, file2.txt:...}
def datalist_from_dirpath (folder):
    data_dic = {}
    os.chdir(folder)
    for file_txt in os.listdir():
        if not file_txt.endswith(".txt"):
            continue
        nrows = return_sharp_nrows(file_txt)
        df1=pd.read_csv(file_txt, header = None, nrows=nrows, encoding="shift_jis", sep="\t",  index_col=0)
        df2=pd.read_csv(file_txt, header = nrows, encoding="shift_jis", sep="\t", index_col="#")
        analysis_date = df1.loc["分析日"].values[0]
        data_dic[file_txt] = {"df1":df1, "df2":df2, "analysis_data":analysis_date}
    return data_dic

#============================
# 重複のあるファイルをリネーム
def rename(bfr, aft, rtn=False):
    """
    bfr: filepath before rename. e.g. /dir1/test1.txt
    aft: file name after rename. e.g. TEST
    """
    dirname = os.path.dirname(bfr) # get directory name.
    bfr = os.path.basename(bfr) # get basename.
    root, ext = os.path.splitext(bfr) # separate basename into file name and extension.
    
    if '/' in aft:
        aft = aft.replace("/","_")
    
    cnt = 1
    tmpaft = aft
    while True:
        if os.path.isfile(tmpaft + ext):
            tmpaft = aft + '_' + str(cnt)
            cnt = cnt +1
        else:
            aft = tmpaft
            break
    os.rename(bfr, dirname+aft+ext)

    for any in glob(root+'.*'): # any file same as bfr is renamed.
        r, e = os.path.splitext(any)
        os.rename(any, dirname+aft+e)
    
    if rtn:
        return bfr, dirname+aft+ext

#====================================
# フォルダ内からデータを収集
folder = input("整理したいフォルダを入力してください ")
data_dic = datalist_from_dirpath(folder)
data_dic = {k: data_dic[k] for k in sorted(data_dic, reverse=False, key=lambda x: data_dic[x]["analysis_data"])} # 分析日でソートする

#============================
#ファイル名を変更する
for data in data_dic:
    df1 = data_dic[data]["df1"]
    sample_name = df1.loc["サンプル名"].values[0]
    bft, aft = rename(bfr=data, aft=sample_name, rtn=True)
    print('before name: {}\t after name: {}'.format(bft, aft))
print('\nファイル名の変更終了')


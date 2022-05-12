import os
import pandas as pd
from glob import glob

#============================
# 重複のあるファイルをリネーム
def rename(bfr, aft, rtn=False):
    """
    Function to rename file name.
    bfr: filepath before rename. e.g. /dir1/test1.txt
    aft: file name after rename. e.g. TEST
    fin_aft: final file name after rename. e.g. TEST_1.txt
    """
    dirname = os.path.dirname(bfr) # get directory name.
    bfr = os.path.basename(bfr) # get basename.
    root, ext = os.path.splitext(bfr) # separate basename into file name and extension.
    
    # '/' is not allowed for file name.
    if '/' in aft:
        aft = aft.replace("/","_")
    
    cnt = 0
    while True:
        try:
            if cnt == 0:
                os.rename(bfr, dirname+aft+ext)
                fin_aft = dirname+aft+ext
                break
            os.rename(bfr, dirname+aft+'_'+str(cnt)+ext)
            fin_aft = dirname+aft+'_'+str(cnt)+ext
            break
        except:
            cnt = cnt +1
    
    for any in glob(root+'.*'): # any file same as bfr is renamed.
        r, e = os.path.splitext(any)
        r_, e_ = os.path.splitext(fin_aft)
        os.rename(any, r_+e)
        
    if rtn:
        return bfr, fin_aft

#============================
# Function to set standart retention time.
def input_rt():
    std_rt=input("基準となるrtを入力してください")
    try:
        std_rt=float(std_rt)
        return std_rt
    except:
        print("エラー: 数字を入力してください")
        return input_rt()

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
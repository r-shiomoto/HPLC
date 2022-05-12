#########################################################
# watersから出力されるtxtとpdfファイルが対象 
#########################################################

print('ライブラリをロード中...\n')
from function import rename, datalist_from_dirpath

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


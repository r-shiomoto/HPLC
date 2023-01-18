import os
from glob import glob
import pandas as pd

folder = input('整理したいフォルダのパスを入力してください')

print('ファイル名を変更します...\n')

file_list = glob(folder+'\\*.XLS')
for file in file_list:
    df = pd.read_excel(file, sheet_name='Manager Report')
    df.columns = [i for i in range(df.shape[1])]
    samplename = df.iloc[8,5]
    if '/' in samplename:
        samplename = samplename.replace('/','_')
    
    dirname, name_ext = os.path.split(file)
    name, ext = os.path.splitext(name_ext)
    new_file = dirname+'\\'+samplename+ext
    
    cnt = 1
    while True:
        try:
            os.rename(file, new_file)
            break
        except:
            dirname, name_ext = os.path.split(new_file)
            name, ext = os.path.splitext(name_ext)
            new_file = dirname+'\\'+samplename+f'_{cnt}'+ext
            cnt = cnt + 1
    
    print('{} -> {}'.format(file, new_file))
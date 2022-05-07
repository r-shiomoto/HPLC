#########################################################
#日立の.xlsファイルからプロファイルを作成
#分析日ごとに並び替え
#pick_upも作成
#########################################################
#データ構造は [[[サンプル名,分析日],[rrt,rt,area,area_percent],[rrt,rt,area,area_percent],,,],,,]
import os,openpyxl,xlrd
from openpyxl.styles.borders import Border, Side
from glob import glob
import datetime

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

def excel_date(num):   #シリアル値を日付に変換する
    from datetime import datetime, timedelta
    return(datetime(1899, 12, 30) + timedelta(days=num))

wb1=openpyxl.Workbook("")
sheet1=wb1.get_sheet_by_name("Sheet")
num=3

folder=input("整理したいフォルダのパスを入力してください")
os.chdir(folder)
data_list=[]
for d in os.listdir():
    if not d.lower().endswith(".xls"):
        continue
    peak_list=[]
    wb=xlrd.open_workbook(d)
    sheet2=wb.sheet_by_name("Manager Report")
    analysis=excel_date(sheet2.cell(4,2).value)  #分析日
    print("分析日は",analysis)
    sample=sheet2.cell(9,5).value  #サンプル名
    print("サンプル名は",sample)
    peak_list.append([sample,analysis,d])
    sheet3 = wb.sheet_by_name("Tables")
    print("rt","\t","area","\t","area%")
    for row in range(10,sheet3.nrows-5):   #一度rt,area,area%を見せてから、基準となるrtを選ばせる
        print(sheet3.cell(row,2).value,"\t",int(sheet3.cell(row,3).value),"\t",sheet3.cell(row,4).value)
    print()
    std_rt=float(input("基準となるrtを入力してください"))
    for row in range(10,sheet3.nrows-5):
        rt=sheet3.cell(row,2).value
        area=int(sheet3.cell(row,3).value)
        rrt=round(rt/std_rt,2)
        area_percent=sheet3.cell(row,4).value
        peak_list.append([rrt,rt,area,area_percent])
    data_list.append(peak_list)

#================
print()
rrt_list=[]   #rrtだけ抽出して、リストを作り、それをエクセルの左端に入力する
for m in data_list:   #   mは[[サンプル名,分析日],[rrt,rt,area,area_percent],[rrt,rt,area,area_percent]...] 
    for n in m:   #   nは[サンプル名,分析日],[rrt,rt,area,area_percent],[rrt,rt,area,area_percent]...
        if len(n)==4:
            rrt_list.append(n[0])
rrt_list=list(set(rrt_list))   #setで重複をまとめて、再度リスト化している
rrt_list.sort()   #sortで順番を並び替える
#print("rrt_listは",rrt_list)確認用
for i,j in zip(rrt_list,range(4,len(rrt_list)+4)):
    sheet1.cell(row=j,column=1).value=i   #左端にrrtを入力する 

#================
#分析日で並び替え
data_list=sorted(data_list, key=lambda x: x[0][1])
#==========================================

#========================================
#data_listをエクセルに書き込む
side = Side(style='thin')
border = Border(right=side)
bottom = Border(bottom=side)
for m in data_list:#   mは[[サンプル名,分析日],[rrt,rt,area,area_percent],[rrt,rt,area,area_percent]...] 
    sheet1.cell(row=1,column=num).value="サンプル名"
    sheet1.cell(row=2,column=num).value="分析日"
    sheet1.cell(row=1,column=num+1).value=m[0][0]  #サンプル名代入
    sheet1.cell(row=2,column=num+1).value=m[0][1]  #分析日を代入
    sheet1.cell(row=3,column=num).value="rrt"
    sheet1.cell(row=3,column=num+1).value="rt"
    sheet1.cell(row=3,column=num+2).value="area"
    sheet1.cell(row=3,column=num+3).value="area%"
    sheet1.cell(row=1,column=num+3).border = border
    sheet1.cell(row=2,column=num+3).border = border
    sheet1.cell(row=3,column=num+3).border = border
    for r in range(4,len(rrt_list)+4):
        sheet1.cell(row=r,column=num+3).border = border
        for n in m[1:]:   # nは[サンプル名,分析日],[rrt,rt,area,area_percent],[rrt,rt,area,area_percent]...
            if n[0]==sheet1.cell(row=r,column=1).value:
                sheet1.cell(row=r,column=num).value=n[0] 
                sheet1.cell(row=r,column=num+1).value=n[1]
                sheet1.cell(row=r,column=num+2).value=n[2]
                sheet1.cell(row=r,column=num+3).value=n[3]
            else:
                continue
    num+=4
sheet1.freeze_panes="C4"
for i in range(1,num):
    sheet1.cell(row=len(rrt_list)+3,column=i).border = bottom
#==========================================

#==========================================
#area%をピックアップしてまとめる
sheet4 = wb1.create_sheet("pick_up")
for i,j in zip(rrt_list,range(4,len(rrt_list)+4)):
    sheet4.cell(row=j,column=1).value=i   #左端にrrtを入力する 
num1=3
sheet4.cell(row=1,column=num1-1).value="サンプル名"
sheet4.cell(row=2,column=num1-1).value="分析日"
for m in data_list:#   mは[[サンプル名,分析日,ファイル名],[rrt,rt,area],[rrt,rt,area]...] 
    sheet4.cell(row=1,column=num1).value=m[0][0]  #サンプル名代入
    sheet4.cell(row=2,column=num1).value=m[0][1]  #分析日を代入
    sheet4.cell(row=3,column=num1).value="area%"
    for r in range(4,len(rrt_list)+4):
        for n in m[1:]:   # nは[サンプル名,分析日,ファイル名],[rrt,rt,area],[rrt,rt,area]...
            if n[0]==sheet1.cell(row=r,column=1).value:
                sheet4.cell(row=r,column=num1).value=n[3]
            else:
                continue
    num1+=1
sheet4.freeze_panes="C4"
#==========================================
dt_now = datetime.datetime.now()
dt_now = dt_now.strftime('%Y%m%d%H%M%S')
dirname = os.path.splitext(os.path.basename(os.getcwd()))[0]
filename = dirname + '_' + dt_now
wb1.save(".\\{}.xlsx".format(filename))
wb1.close()

#============================
#ファイル名を変更する
import os.path
for m in data_list:#   mは[[サンプル名,分析日,ファイル名],[rrt,rt,area],[rrt,rt,area]...] 
    sample_name = m[0][0]
    file_name = m[0][2]
    bft, aft = rename(bfr=file_name, aft=sample_name, rtn=True)
    print('before name: {}\t after name: {}'.format(bft, aft))
print('\nファイル名の変更終了')

print()
print("================================")
print("エクセルファイルを作成しました。")
print("================================")
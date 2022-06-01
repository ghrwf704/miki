import os

def yyyymmdd(path,x):
    import pathlib    
    import datetime
    path=path+x
    today=datetime.date.today()
    p=pathlib.Path(path)
    ctime=datetime.date.fromtimestamp(p.stat().st_ctime)
    atime=datetime.date.fromtimestamp(p.stat().st_atime)
    mtime=datetime.date.fromtimestamp(p.stat().st_mtime)
    #5年以上（365*5）離れていたらflgを立てる
    cflg=str(int((int((today-ctime).days)-(365*5))>0))
    aflg=str(int((int((today-atime).days)-(365*5))>0))
    mflg=str(int((int((today-mtime).days)-(365*5))>0))
    ctime=str(ctime)
    atime=str(atime)
    mtime=str(mtime)
    size = os.path.getsize(path) / 1000
    size = "{:,.0f}".format(size)
    return '''"{size}","{ctime}","{cflg}","{atime}","{aflg}","{mtime}","{mflg}"'''.format(size=str(size),ctime=str(ctime),cflg=str(int(cflg)),atime=str(atime),aflg=str(int(aflg)),mtime=str(mtime),mflg=str(int(mflg))).strip()
    
def dir_check(path):
    path=path+"\\"
    x=""
    for x in os.listdir(path): 
        if os.path.isdir(path wb[i].insert.columns+ x):  #パスに取り出したオブジェクトを足してフルパスに
            dir_check(path + x)
        elif os.path.isfile(path + x):  #isdirの代わりにisfileを使う
            f=open("tree.csv","a+", encoding='CP932', errors='ignore')
            f.write("\""+path+"\",\""+x+"\","+yyyymmdd(path,x)+"\n")
            f.close()

# path で、検索フォルダーを指定して実行！     複数フォルダー指定は、L23～L39をコピー追加し、ファイル名（L30とL36）を変更
path = u'Z:\情報システム部'      ####           
f=open("tree.csv","w", encoding='CP932', errors='ignore')
f.write('''"ディレクトリ","ファイル名","サイズKB","作成日付","作成F","最終アクセス日付","アクセスF","更新日付","更新F"\n''')
f.close()
dir_check(path)


"""                       フォルダー追加（2個目）
def dir_check(path):
    path=path+"\\"
    x=""
    for x in os.listdir(path): 
        if os.path.isdir(path + x):  #パスに取り出したオブジェクトを足してフルパスに
            dir_check(path + x)
        elif os.path.isfile(path + x):  #isdirの代わりにisfileを使う
            f=open("tree2.csv","a+", encoding='CP932', errors='ignore')             ####
            f.write("\""+path+"\",\""+x+"\","+yyyymmdd(path,x)+"\n")
            f.close()

# path で、検索フォルダーを指定して実行！ 
path = u'Z:\情報システム部\自部門用\●_予実管理\◎ｼｽﾃﾑ支払関連資料\随時支払'               ####          
f=open("tree2.csv","w", encoding='CP932', errors='ignore')                          ####
f.write('''"ディレクトリ","ファイル名","サイズKB","作成日付","作成F","最終アクセス日付","アクセスF","更新日付","更新F"\n''')
f.close()
dir_check(path)
"""
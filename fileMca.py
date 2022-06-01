# -*- coding: utf-8 -*-
"""
Created on Wed Nov 27 09:40:17 2019

@author: x3553yoshigae
"""

#社内サーバーを整理するにあたって、まずは現行のフォルダ構成やファイル更新日付などを出力する
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
    
    return '''"{ctime}","{cflg}","{atime}","{aflg}","{mtime}","{mflg}"'''.format(ctime=str(ctime),cflg=str(int(cflg)),atime=str(atime),aflg=str(int(aflg)),mtime=str(mtime),mflg=str(int(mflg))).strip()
    
def dir_check(path):
    path=path+"\\"
    x=""
    for x in os.listdir(path): 
        if os.path.isdir(path + x):  #パスに取り出したオブジェクトを足してフルパスに
            dir_check(path + x)
        elif os.path.isfile(path + x):  #isdirの代わりにisfileを使う
            f=open("tree.csv","a+")
            f.write("\""+path+"\",\""+x+"\","+yyyymmdd(path,x)+"\n")
            f.close()
                
path = 'D:'  #ディレクトリ一覧を取得したいディレクトリ

f=open("tree.csv","a+")
f.write('''"ディレクトリ","ファイル名","最終表示日付","cflg","変更日付","aflg","作成日付","mflg"\n''')
f.close()

dir_check(path)
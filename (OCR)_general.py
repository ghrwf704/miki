import subprocess
import re

import datetime
import os
import csv
import requests
import time
import numpy as np
import cv2
import pdf2image
import PyPDF2
import socks
import shutil
import send2trash

#imgのバイナリからpngファイルを生成
#cvの関数でなく、自分で定義（バグがあって2倍と表示は文字化けする）
def imwrite(filename, img, params=None):
    try:
        ext = os.path.splitext(filename)[1]
        result, n = cv2.imencode(ext, img, params)
        if result:
            with open(filename, mode='w+b') as f:
                n.tofile(f)
            return True
        else:
            return False
    except Exception as e:
        print(e)
        return False

#元間を少し修正、複数のPDFファイルでもページ分割するように修正
def split_pdf_pages(src_path):
    dst_basepath=src_path.replace(".pdf","")
    src_pdf = PyPDF2.PdfFileReader(src_path)
    for i in range(src_pdf.numPages):
        if i==src_pdf.numPages:
            return -1
        dst_pdf = PyPDF2.PdfFileWriter()
        dst_pdf.addPage(src_pdf.getPage(i))
        if os.path.isfile('{}_{}.pdf'.format(dst_basepath, i)):
            return -1
        with open('{}_{}.pdf'.format(dst_basepath, i), 'wb') as f:
            dst_pdf.write(f)

#正規表現(使っていないけど念のため)
def regmatch(src,pat):
	search_result = re.search(pat, src)
	if search_result:
	  return search_result.group()
	else:
	  return 0

#ファイル名を受け取り、画像のアップロードとOCR（文字化）してJSON形式で結果を返す
def readImg(temp,i):
    subprocess.Popen("mogrify -format png "+os.getcwd()+"\\undone\\*.jpg", shell=True)
    os.rename("undone\\"+temp,"undone\\temp"+str(i)+".png")
    #アップロード用のphpファイルに接続
    url = 'http://mikipulley.sakura.ne.jp/microsoft/azuru/ocr/img/upload.php'
    file = {'upfile': open(os.getcwd()+"\\undone\\"+"temp"+str(i)+".png", 'rb')}
    res = requests.post(url, files=file)#画像ファイルなどをアップするのに必要
    #ocr用のコマンドを作成
    url = "https://x8g7u5g3e0.apigw.ntruss.com/custom/v1/86/9b358ba9f00b3bebfadb934df4a2bdbe7f9d77e00ae5b4c96dfd38bd4b854dd2/general"
    payload = "{\r\n    \"images\": [\r\n      {\r\n        \"format\": \"png\",\r\n        \"name\": \"All\",\r\n        \"data\": null,\r\n        \"url\": \""+ 'http://mikipulley.sakura.ne.jp/microsoft/azuru/ocr/img/files/'+"temp"+str(i)+".png"+"\"\r\n      }\r\n    ],\r\n    \"lang\": \"ja\",\r\n    \"requestId\": \"string\",\r\n    \"resultType\": \"string\",\r\n    \"timestamp\": 1582166299,\r\n    \"version\": \"V1\"\r\n}"
    headers = {
            'Content-Type': 'application/json',
            'X-OCR-SECRET': 'dlZVaEprWEpzY1Nib094b1B5amF1c05wYkJub2tTYXk='
            }
    try:
        #コマンド送信
        response = requests.request("POST", url, headers=headers, data = payload)
    except:
        response=""
    return response.json()["images"][0]
 

path=os.getcwd()+'\\undone'
files = os.listdir(path)
files_file = [f for f in files if os.path.isfile(os.path.join(path, f))]

for i in range(len(files_file)):
    if ".pdf" in files_file[i]:
        # .spyder-py3\Data フォルダーの、 結合されている merge.pdf から  split_0,split_1,...へ 分割
        split_pdf_pages(os.getcwd()+'\\undone\\'+files_file[i])   
        #os.remove("undone\\"+files_file[i])

os.getcwd()
q = os.chdir("undone")
filepath = os.listdir(q)
#os.system("del /q \""+os.getcwd().replace("\\Data","\\CVImage")+"\\*.*\"")
os.chdir("../")
#PDFをⅠページづつに分解してそれぞれPNGに変換する
for file in filepath:
    basename = os.path.basename(file)
    if ".pdf" in basename:
        pdfimages = pdf2image.convert_from_path("undone\\"+file)    
        cvimage = np.asarray(pdfimages[0])
        cvimage = cv2.cvtColor(cvimage, cv2.COLOR_RGB2BGR)
        #cvimage = cv2.resize(cvimage, (1920, 1050))
        temp=str(basename).replace('.pdf','.png')
        imwrite('undone\\'+temp, cvimage)
        #send2trash.send2trash("undone\\"+basename+".pdf")

#全ての画像、jpg,jpeg,tif,bmpをpngに変換
subprocess.Popen("mogrify -format png "+os.getcwd()+"\\undone\\*.jpg", shell=True)
subprocess.Popen("mogrify -format png "+os.getcwd()+"\\undone\\*.jpeg", shell=True)
subprocess.Popen("mogrify -format png "+os.getcwd()+"\\undone\\*.tif", shell=True)
subprocess.Popen("mogrify -format png "+os.getcwd()+"\\undone\\*.bmp", shell=True)

#対象のファイル名を書き出す（ｓｔｒ１にカンマ区切りで格納しておく)
subprocess.call("dir /B "+os.getcwd()+"\\undone\\*.png>"+os.getcwd()+"\\file.txt",shell=True)
str1=""
t=open('file.txt')
f=t.readline()
while f:
    if ""==f:
        continue
    if "tmp.png"==f:
        continue
    str1=str1+f+","
    f=t.readline()
t.close()
str1=str1.replace("\n","")

#tifs変数に配列として、ファイル名を格納
tifs=str1.split(",")
ocr_output=""

#ファイルサイズを減らしてOCRできるように（1000固定をどうするか後程検討）
#塊毎にまとまって出力するように後日修正（LINEに発注）
subprocess.call("cd /d "+os.getcwd()+"\\undone&&Iconvert -geometry mogrify -resize '1000x1000>' *.png",shell=True)
x=0
for i in range(0, len(tifs)-1):
    print("\n\n"+tifs[i])
    if tifs[i]!="":
        images=readImg(tifs[i],i)
        if "field" in str(images):
            shutil.move("undone\\temp"+str(i)+".png", "undone\\"+tifs[i])
            for txt_lines in images["fields"]:
                if x<txt_lines['boundingPoly']['vertices'][0]['x']:
                    f=open("undone\\"+tifs[i]+".txt","a+")
                    f.write(" "+txt_lines["inferText"].encode('cp932', "ignore").decode('cp932'))
                    f.close
                    print(" "+txt_lines["inferText"], end="")
                    x=txt_lines['boundingPoly']['vertices'][0]['x']
                else:
                    #次の行
                    f=open("undone\\"+tifs[i]+".txt","a+")
                    f.write("\n "+txt_lines["inferText"].encode('cp932', "ignore").decode('cp932'))
                    f.close
                    x=txt_lines['boundingPoly']['vertices'][0]['x']

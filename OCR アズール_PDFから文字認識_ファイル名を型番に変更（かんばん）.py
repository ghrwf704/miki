import os
import cv2
import requests
import numpy as np
import pdf2image
import codecs
import re
from datetime import datetime
import time
import socks
import pymraw

def text_to_points(txt):
    tmp = txt.split(",")
    p = [int(tmp[0]), int(tmp[1]), int(tmp[2]), int(tmp[3])]
    return (p[0], p[1]), (p[0]+p[2], p[1]+p[3])

#プロキシ設定
proxies = {
  'https': 'swg-proxy.mikipulley.co.jp:8080',
}
socks.set_default_proxy(socks.SOCKS5, "swg-proxy.mikipulley.co.jp", 8080)

#初期化処理、対象のディレクトリに存在する「txt」[jpg]をあらかじめ削除しておく
os.getcwd()
q = os.chdir("Data")
os.system("del /q \""+os.getcwd()+"\\*.txt\"")

filepath = os.listdir(q)
os.system("del /q \""+os.getcwd().replace("\\Data","\\CVImage")+"\\*.*\"")

#ファイルの数繰り返す
for file in filepath:
    time.sleep(1)
    basename = os.path.basename(file)
    pdfimages = pdf2image.convert_from_path(file)    
    cvimage = np.asarray(pdfimages[0])
    cvimage = cv2.cvtColor(cvimage, cv2.COLOR_RGB2BGR)
     
    #httpでファイルをアップロードする
    temp=str(basename).replace('.pdf','.jpg')
    cv2.imwrite(temp, cvimage)
    
    cv2.waitKey(-1)
    os.system("move \""+os.getcwd()+"\\"+temp+"\" \""+os.getcwd().replace("\\Data","\\CVImage\""))

    local_image = cv2.imread("..\\cvimage\\"+temp)

    #アップロード用のphpファイルに接続
    url = 'http://mikipulley.sakura.ne.jp/microsoft/azuru/ocr/img/upload.php'
    file = {'upfile': open("..\\cvimage\\"+temp, 'rb')}
    res = requests.post(url, files=file)#画像ファイルなどをアップするのに必要
    
    #ocr用のコマンドを作成
    ocr_url = 'https://japaneast.api.cognitive.microsoft.com/vision/v2.0/ocr'
    headers  = {'Ocp-Apim-Subscription-Key': 'd2b5ec46b3874a41877fa066466e5b2d'}
    params   = {'language': 'ja', 'detectOrientation ': 'true'}
    data     = {'url': 'http://mikipulley.sakura.ne.jp/microsoft/azuru/ocr/img/files/'+temp}
    response = requests.post(ocr_url, headers=headers, params=params, json=data, proxies=proxies)
    response.raise_for_status()
    
    #OCRデータ（json）
    ocr_data = response.json()
    
    output = "\n"
    count = 0
    
    #解析しておらず（JSONの分解）　サンプル引用
    for txt_lines in ocr_data['regions']:
        p1, p2 = text_to_points(txt_lines['boundingBox'])
        # cv2.rectangle(local_image, p1, p2, (0, 0, 255), 3)
    
        for txt_words in txt_lines['lines']:
            p1, p2 = text_to_points(txt_words['boundingBox'])
            cv2.rectangle(local_image, p1, p2, (0, 255, 0), 2)
            cv2.putText(local_image, str(count), p1, cv2.FONT_HERSHEY_COMPLEX_SMALL, 2, (0, 0, 0), 2)
    
    
            for txt_word in txt_words['words']:
                p1, p2 = text_to_points(txt_word['boundingBox'])
                cv2.rectangle(local_image, p1, p2, (255, 0, 0), 1)
    
                output += txt_word['text']
            output += '\n'
        output += '\n'
        
    #現在時刻と品番注番からpdfをリネーム
    now = datetime.now()
    datestr = f'{now:%Y}{now:%m}{now:%d}_{now:%H}{now:%M}{now:%S}'
    output = output.translate(str.maketrans({'　': '', ' ': '',}))
    cyuban=""
    output2=""
    hinban = re.search(r'[^0-9][0-9]{9}[^0-9]', output)
    if (hinban is not None):
        hinban=re.search("\d{9}",hinban[0])
        output2=output.replace(hinban[0],"")
        cyuban = re.search(r'[^0-9][0-9]{7}[^0-9]', output2)
        if(cyuban is not None):
            cyuban=re.search("\d{7}",cyuban[0])
            os.system("ren "+temp.replace(".jpg","")+".pdf かん_"+hinban[0]+"_"+cyuban[0]+"_"+datestr+".pdf")
    f = codecs.open(temp+'.txt', 'w', "utf8")
    f.write(output.replace('\n\n','\n'))
    f.close()

 

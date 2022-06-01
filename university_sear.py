import requests as rq
import time
import socks
import datetime
from bs4 import BeautifulSoup
import urllib.parse
import urllib3
from urllib3.exceptions import InsecureRequestWarning
urllib3.disable_warnings(InsecureRequestWarning)


#プロキシ設定
proxies = {
  'https': 'swg-proxy.mikipulley.co.jp:8080',
}
socks.set_default_proxy(socks.SOCKS5, "swg-proxy.mikipulley.co.jp", 8080)

#検索文言設定(エクセルから取り込む)
d=datetime.datetime.today()
d=str(d.year)+str(d.month)+str(d.day)

#HTMLとCSSの書き込み
with open("result_"+d+".html", encoding='utf-8',mode='w') as f:
    header='''
    <html>
    <head><title>検索結果</title></head>
    <style>
    table {
       table-layout: fixed;     /* 固定レイアウト */
       border-collapse: collapse; /* 隣接する枠線を重ねる */
       border: 2px solid green;   /* 外枠：2px,実線,緑色 */
       width: 100%;              /* 表の横幅：350ピクセル */
    }
    .back{
          background-color: cornsilk;
    }
    table td {
       border: 1px solid green; /* 表内側の線：1px,実線,緑色 */
       padding: 3px;            /* セル内側の余白：3ピクセル */
    }
    
    body {
        word-wrap: break-word;
    }
    </style>
    <body>
    <table>
    <tbody>
    <th>タイトル</th><th>URL</th>
    '''
    f.write(header)

#設定用ファイル（エクセル）を読み込む→パターンごとに検索していく
import openpyxl as px
wb = px.load_workbook(u'検索設定.xlsx')
mySheet = wb.worksheets[0]
s=rq.Session()
for row1 in range(2,9999):
    if (mySheet["A"+str(row1)].value) is None:
        break
    else:
        searchWord="intitle:\""+mySheet["A"+str(row1)].value+"\""
        for row2 in range(2,9999):
            if (mySheet["B"+str(row2)].value) is None:
                break
            else:
                searchWord=searchWord+" "+mySheet["B"+str(row2)].value
        with open("result_"+d+".html", encoding='utf-8',mode='a+') as f:
            f.write("<tr><tr><td>　</td><td>　</td>\n")
            f.write("<tr></tr><th class=\"back\">検索キーワード -- "+searchWord+" -- </th><th class=\"back\"> </th>")
        #重複回避用
        flg=0
        #URLをエンコード
        for b in range(1,999,10):
            time.sleep(3)
            html=s.get("https://search.yahoo.co.jp/search?"+urllib.parse.urlencode({'p': searchWord})+"&b="+str(b), timeout=100, verify=False).text
            if "一致するウェブページは見つかりませんでした" in html:
                with open("result_"+d+".html", encoding='utf-8',mode='a+') as f:
                    f.write("<tr><tr><td>検索結果がありません。</td><td>　</td>\n")
                break
            #soupにBeautifulSoupのオブジェクト格納
            soup = BeautifulSoup(html, "html.parser")
            #欲しい情報はli-a配下のhrefとtitle
            bb=soup.find_all("a")
            #次のページへのリンクがあるかを示すflg
            nextFlg=0
            for ii in range(3,len(bb),1):
                href=bb[ii]['href']
                title=bb[ii].text
                if "次へ" in title:
                    nextFlg=1
                if "search.yahoo.co.jp" not in href:
                    time.sleep(2)
                    with open("result_"+d+".html", encoding='utf-8',mode='a+') as f:
                        f.write("<tr><tr><td>"+title+"</td><td><a  target=\"_blank\" href=\""+href+"\">"+href+"</a></td>\n")
                    print(title+"@"+href)
            print(str(b)+"ぺ－ジ目完了\n")
            if nextFlg==1:
                continue
            else:
                #次の検索パターンに移る
                break
 
with open("result_"+d+".html", encoding='utf-8',mode='a+') as f:
    footer='''
    </tbody>
    </table>
    </body>
    </html>
    '''
    f.write(footer)
        

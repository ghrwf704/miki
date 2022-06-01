# webdriver の情報
from selenium import webdriver
# html の タブの情報を取得
from selenium.webdriver.common.by import By
# キーボードを叩いた時に web ブラウザに情報を送信する
from selenium.webdriver.common.keys import Keys
# 次にクリックしたページがどんな状態になっているかチェックする
from selenium.webdriver.support import expected_conditions as EC
# 待機時間を設定
from selenium.webdriver.support.ui import WebDriverWait
# 確認ダイアログ制御
from selenium.webdriver.common.alert import Alert
import keyboard

def printExecute():
    driver.execute_script('document.querySelector("#frm\\:j_id4_dp > img")[0].click();');
            
driver = webdriver.Ie('IEDriverServer.exe')
driver.get("http://192.168.200.28")

#keyboard.wait("alt+p")

printExecute()


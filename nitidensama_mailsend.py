import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.utils import formatdate
from email.mime.application import MIMEApplication
from os.path import basename
from email.mime.text import MIMEText
def addAttach(name):
    with open(name, "rb") as f:
        part = MIMEApplication(
            f.read(),
            Name=basename(name)
        )
    part['Content-Disposition'] = 'attachment; filename="%s"' % basename(name)
    return part    

"""
(1) MIMEMultipartでメッセージを作成
"""
main_text = "井上さん、お疲れ様です。本件ファイルにて送付いたします。\nご確認ください。\n（実際に稼働させる際には少し修正が必要なのでその際にはお声がけください。\n\nよろしくお願いします）"
charset = "utf-8"
msg = MIMEMultipart()
msg["Subject"] = "ファイル送信プログラム完成しました【吉ヶ江】"
msg["From"] = "yoshigae.kei@mikipulley.co.jp" 
msg["To"] = "h_inoue@mikipulley.co.jp"
msg["Cc"] = ""
msg["Bcc"] = "shigemi@mikipulley.co.jp"
msg["Date"] = formatdate(None,True)
body = MIMEText(main_text.encode("utf-8"), 'plain', 'utf-8')
msg.attach(body)

host = "mail.securemx.jp"
nego_combo = ("ssl", 465) # ("通信方式(ssl)", port番号)
context = ssl.create_default_context()
smtpclient = smtplib.SMTP_SSL(host, nego_combo[1], timeout=10, context=context)
smtpclient.set_debuglevel(2) # サーバとの通信のやり取りを出力

"""
(2) サーバーにログイン
"""
USERNAME = "yoshigae.kei@mikipulley.co.jp"
PASSWORD = "czKOFxD2op"
smtpclient.login(USERNAME, PASSWORD)

"""
(3) 添付ファイル追加
"""
msg.attach(addAttach("添付.txt"))
msg.attach(addAttach("mailsend.py"))

"""
(4) メールを送信する
"""
smtpclient.send_message(msg)
smtpclient.quit()

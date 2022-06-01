import requests    #1

def talk_api(message):
    apikey = "DZZCaxheFII0qqzHjc3annnpvife7hLJ"
    talk_url = "https://api.a3rt.recruit-tech.co.jp/talk/v1/smalltalk"    #4

    payload = {"apikey": apikey, "query": message}    #5
    response = requests.post(talk_url, data=payload)

    try:
        #返信する内容を返す
        
        return response.json()["results"][0]["reply"]    #6
    except:
        return "システム管理者いお問い合わせください。"

def main():
     while(True):
         print("あなた：", end="")    #2
         message = input()

         print("BOT：" + talk_api(message))    #3  #7

if __name__ == "__main__":
     main()
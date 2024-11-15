import os
import logging
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError
from dotenv import load_dotenv
import datetime
import glob



def date():
   t_delta = datetime.timedelta(hours=9)                             # 9時間
   JST = datetime.timezone(t_delta, 'JST')                           # UTCから9時間差の「JST」タイムゾーン
   dt = datetime.datetime.now(JST)                                   # タイムゾーン付きでローカルな日付と時刻を取得
   date=dt.strftime("%Y-%m-%d")                                      # 現在日
   return str(date)
 

# デバッグレベルのログを出力します
#logging.basicConfig(level=logging.DEBUG)

load_dotenv(verbose=True)
slack_token = os.getenv("SLACK_API_BOT_TOKEN")
slack_channel = os.getenv("SLACK_CHANNEL")


def file_upload():
           client = WebClient(token=slack_token)
           file_uploads = []
           #dataフォルダ配下のファイルをリスト化
           files = glob.glob(r"./mailbox/*")
   
           #dataフォルダ内にファイルが格納されている場合
           if files != []:
              for file in files:
                 #filepath= folder + file
                 filename =os.path.basename(file)
                 #print(file)
                 file_uploads.append({
                      'file': file,
                      'filename' : filename,
                      'title' : filename
                 })

              new_file = client.files_upload_v2(
                      file_uploads=file_uploads,
                      channel=slack_channel,
                      initial_comment=date()+":bow:"
        
              )
               
              #dataフォルダ内のファイルを削除する。
              for file in files:
                      os.remove(file)
  
if __name__ == "__main__":
    file_upload()

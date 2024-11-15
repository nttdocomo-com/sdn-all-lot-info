# モジュールのインポート.
from ftplib import FTP

#jsonモジュールをインポート
import json

#schedule,timeモジュールをインポートし常時起動を行う
import schedule
import time

#日時モジュールをインポート
import datetime

#OS制御モジュールインポート
import os

#ディレクトリ内のファイル一覧を取得するモジュールをインポート
import glob

#SMTP関連
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# jsonファイルを読み込み
m = open(r'../acc.json', 'r')
js = json.load(m)
ops_eastserver = js['ops_info']['east_server']
ops_userid = js['ops_info']['ops_userid']
ops_password = js['ops_info']['ops_password']


def job_silent():
    #今日の日時を取得
    dt_now = datetime.datetime.now()
    print("起動日時",dt_now)
    
    # FTP接続
    try:
        ftp = FTP(ops_eastserver)
        ftp.set_pasv('true')
        #filepath_nfm= r'共通\仮想化収容NW機器\70.Tool\mieruka\nfm_memcheck' #nem_memcheck用
        filepath_silent= r'共通\仮想化収容NW機器\70.Tool\65.サイレント故障検知ツール\output'
        ftp.login(ops_userid,ops_password)
        ftp.encoding = 'shift_jis'
        #ftp.cwd(filepath_nfm)
        #ftp_list_nfm = ftp.nlst(".")
        #filepath_silent= r'共通\仮想化収容NW機器\70.Tool\65.サイレント故障検知ツール'
        ftp.cwd(filepath_silent)
        ftp_list_silent = ftp.nlst(".")
        print("complete ftp-connection") #for-debug

    except:
        print("failed-to-ftp-connection")
        return

    #2024-10-10 追加 サイレント故障ツールログ取得
    for file in ftp_list_silent:
        if "all_lot" in file:
            #print(file) #for-debug
            #all_lot_infoの取得
            try:
                #C:\temp\nfm_log
                with open(r'./silent-check/' + file, 'wb') as f: #ローカルパス
                    ftp.retrbinary('RETR %s' % file, f.write)
                        #ftp.delete(file)
            except:
                #エラ―解析用      
                import traceback
                with open('error.log', 'a') as f:
                    traceback.print_exc(file=f)
                return
                
            #データ加工するモジュールの実行
            try:
                import arrange_all_lot_info_sdn
                arrange_all_lot_info_sdn.arrange()
                import mailsend_test_silent # type: ignore
                mailsend_test_silent.mail_send_silent()
        
            except:
                #エラ―解析用      
                import traceback
                with open('error.log', 'a') as f:
                    traceback.print_exc(file=f)
                return


        
#毎日9時20に実行
schedule.every().day.at("09:20").do(job_silent)

#デバッグ用
# job_silent()


while True:
    schedule.run_pending()
    #1秒ごとに実行タイミングをチェック
    time.sleep(1)

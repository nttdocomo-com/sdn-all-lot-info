def mail_send_silent():
   import os
   import re
   import glob
   import datetime

   #SMTP関連
   import smtplib
   from email.mime.text import MIMEText
   from email.mime.multipart import MIMEMultipart
   from email.mime.application import MIMEApplication

   print("start")
   try:

      # 送受信先
      #to_addr = "ip-tech-aci-ml@ml.nttdocomo.com"
      to_addr = "kouhei.ichiyama.mt@nttdocomo.com"
      from_addr = "yochoukenchi.kasoukanettowaku.gf@s1.nttdocomo.com"
      body= ""
      
      msg=MIMEMultipart()
      msg['Subject'] = "[TEST]all-lot-info"
      msg['From'] = from_addr
      msg['To'] = to_addr
      msg.attach(MIMEText(body))
      
      #カレントパスの絶対パスを取得
      folder = os.getcwd()
      #print(folder)
      
      #dataフォルダ配下のファイルをリスト化
      files = glob.glob(r"./testbox/*")
   
      #dataフォルダ内にファイルが格納されている場合
      if files != []:
          for file in files:
              filepath= folder + file
              file =os.path.basename(file)
              print(file)
              with open(filepath,"rb") as f:
                 attachment = MIMEApplication(f.read())
                 
              attachment.add_header("Content-Disposition","attachment",filename=file)
              msg.attach(attachment)

          #メール送信
          #with smtplib.SMTP_SSL(host="smtp.ddreams.jp", port=465) as smtp:
          with smtplib.SMTP(host="smtp.ddreams.jp", port=25) as smtp:
               #smtp.login(smtp_userid, smtp_password)         
               smtp.send_message(msg)
               smtp.quit()
                   
      #dataフォルダ内のファイルを削除する。
      # for file in files:
      #     os.remove(file)                

      #今日の日時を取得
      dt_now = datetime.datetime.now()
      print("起動日時",dt_now)
   
   except:
          print("error")
          #エラ―解析用      
          import traceback
          with open("./error/error_" + str(datetime.date.today()) + ".log", 'a') as f:
             traceback.print_exc(file=f)
          return

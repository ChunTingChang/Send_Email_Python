#!/usr/bin/env python
# coding: utf-8


import smtplib
from email.mime.multipart import MIMEMultipart #email內容載體
from email.mime.text import MIMEText #用於製作文字內文
from email.mime.base import MIMEBase #用於承載附檔
from email import encoders #用於附檔編碼
import datetime
import ssl

# 預設本周要寄出上周一至周日的報告，故抓出上周一的日期
today_date = datetime.date.today()
days_to_mon = today_date.weekday()

this_mon = today_date - datetime.timedelta(days = days_to_mon)
last_mon = this_mon - datetime.timedelta(days = 7)


#寄件者使用的Gmail帳戶資訊
gmail_user = 'k830716@gmail.com'
gmail_password = 'secret'
from_address = gmail_user
  
#設定信件內容與收件人資訊
to_address = ['myfriend@gmail.com', 'myfmaily@gmail.com']  
Subject = "Here is the Weekly Report ({})".format(last_mon)
contents = """
Hi my friend,
Attached please find the Weekly Play Report ({}) you requested.

Feel free to reach me if you guys have any suggestion.

Regards,
Angela
""".format(last_mon)

# 設定附件（可設多個）
attachments = ['path\\Play report {}.xlsx'.format(last_mon)]
 
#開始組合信件內容
mail = MIMEMultipart()
mail['From'] = from_address
mail['To'] = ', '.join(to_address)
mail['Subject'] = Subject
#將信件內文加到email中
mail.attach(MIMEText(contents))     
#將附加檔案們加到email中   
for file in attachments:
    with open(file, 'rb') as fp:
        add_file = MIMEBase('application', "octet-stream")
        add_file.set_payload(fp.read())
    encoders.encode_base64(add_file)
    add_file.add_header('Content-Disposition', 'attachment', filename='Play report {}.xlsx'.format(last_mon))
    mail.attach(add_file)
 

# 設定smtp伺服器並寄發信件    
smtpserver = smtplib.SMTP_SSL("smtp.gmail.com", 465)
smtpserver.ehlo()
smtpserver.login(gmail_user, gmail_password)
smtpserver.sendmail(from_address, to_address, mail.as_string())
smtpserver.quit()


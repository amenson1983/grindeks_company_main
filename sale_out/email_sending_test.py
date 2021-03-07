import smtplib
import ssl
import locale
from datetime import datetime

import pandas as pd

now = pd.Timestamp.now()
smtp_server = "smtp.gmail.com"
port = 465
login = "amenson1983@gmail.com"
password = "Chernayamast_16"
sender = "amenson1983@gmail.com"
receivers = "amenson1983@gmail.com"


myDate = datetime.today().date()
myMonth = myDate.strftime('%B')
now_ = str(myDate)[8:10] #split(sep='-')
now_ = int(now_)
the_day_before = now_-1
print(the_day_before)
myDate.replace(1,2,the_day_before)
message = f"""\
Subject: Updated Big table report as of {myDate}
Dear colleagues, good afternoon,
please find BIG TABLE REPORT in attachments updated as of {the_day_before}-th of {myMonth} (including it).
Have a nice day"""

context = ssl.create_default_context()
print('Starting to send')
with smtplib.SMTP_SSL(smtp_server,port,context=context) as server:
    server.login(login,password)
    server.sendmail(sender,receivers,message)
print('Sent')


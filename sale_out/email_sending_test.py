
from datetime import datetime

myDate = datetime.today().date()
myMonth = myDate.strftime('%B')
now_ = str(myDate)[8:10] #split(sep='-')
now_ = int(now_)
the_day_before = now_-1
print(the_day_before)
subject = f"Updated Big table report as of {myDate}"


message = f"""\
Subject: Updated Big table report as of {myDate}
Dear colleagues, good afternoon,
please find BIG TABLE REPORT in attachments updated as of {the_day_before}-th of {myMonth} (including it).
Have a nice day

Best regards 
Andriy Turchyn

JSC "Grindeks" Representative Office in Ukraine
Product Market Research Analytics Specialist
andriy.turchyn@grindeks.ua
M: +380 503576118
www.grindeks.eu
"""

text = f"""\
Уважаемые коллеги, добрый день,\n
в приложении обновленный отчет на {the_day_before}-е число включительно.\n
Хорошего дня!\n
\n
Best regards\n 
Andriy Turchyn\n
\n
\n
\n
**************************************************
\n
JSC "Grindeks" Representative Office in Ukraine\n
Product Market Research Analytics Specialist\n
andriy.turchyn@grindeks.ua\n
M: +380 503576118\n
www.grindeks.eu\n
"""

def save_to_file_in_new_row(arr, filename = "C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\big_table_report_ukraine\\email.txt"):
    with open(filename, "w") as file_write:
          file_write.write(str(arr) + "\n")

save_to_file_in_new_row(text)
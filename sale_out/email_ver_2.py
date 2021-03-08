from email.mime.application import MIMEApplication

import requests
import lxml
from pandas.tests.dtypes.test_missing import now

smtp_server = "smtp.gmail.com"
port = 587
smtp_server_1 = "mail.grindeks.ua"
port_1 = 587
login_1 = "andriy.turchyn@grindeks.ua"
password_1 = "inula685"
login = "amenson1983@gmail.com"
password = "Chernayamast_16"
now_ = str(now)[0:10]


RECEIVER_EMAIL_1 = "amenson1983@gmail.com"
RECEIVER_EMAIL_2 = "turchyna.natali82@gmail.com"
RECEIVER_EMAIL_3 = "andriy.turchyn@grindeks.ua"
RECEIVER_NAME_1 = 'Андрей'
RECEIVER_NAME_2 = 'Натичек, сонечко '
RECEIVER_NAME_3 = 'Андрей'
import smtplib
import ssl
from email.mime.text import MIMEText
from email.utils import formataddr
from email.mime.multipart import MIMEMultipart  # New line
from email.mime.base import MIMEBase  # New line
from email import encoders  # New line

# User configuration
sender_email = "amenson1983@gmail.com"
sender_name = "amenson1983@gmail.com"


receiver_emails = [RECEIVER_EMAIL_1, RECEIVER_EMAIL_2, RECEIVER_EMAIL_3]
receiver_names = [RECEIVER_NAME_1, RECEIVER_NAME_2, RECEIVER_NAME_3]

# Email body
email_html = open('C:\\Users\\Anastasia Siedykh\\PhpstormProjects\\grindex_main_company\\Form.html',encoding="UTF-8")
email_body = email_html.read()

filename = 'C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\big_table_report_ukraine\\big_table_report_2021_new_1.xlsm'
filename_text = 'big_table_report_2021_new_1.xlsm'
for receiver_email, receiver_name in zip(receiver_emails, receiver_names):
        print("Sending the email...")
        # Configurating user's info
        msg = MIMEMultipart()
        msg['To'] = formataddr((receiver_name, receiver_email))
        msg['From'] = formataddr((sender_name, sender_email))
        msg['Subject'] =  receiver_name + f' , данные были обновлены сегодня - {now_}'

        msg.attach(MIMEText(email_body, 'html'))
        part = MIMEApplication(open('C:\\Users\\Anastasia Siedykh\\PhpstormProjects\\grindex_main_company\\image002.png', 'rb').read())
        part.add_header('Content-Disposition', 'attachment', filename='image002.png')
        msg.attach(part)
        part = MIMEApplication(open('C:\\Users\\Anastasia Siedykh\\PhpstormProjects\\grindex_main_company\\logo.png', 'rb').read())
        part.add_header('Content-Disposition', 'attachment', filename='logo.png')
        msg.attach(part)

        try:
            # Open PDF file in binary mode
            with open(filename, "rb") as attachment:
                            part = MIMEBase("application", "octet-stream")
                            part.set_payload(attachment.read())

            # Encode file in ASCII characters to send by email
            encoders.encode_base64(part)

            # Add header as key/value pair to attachment part
            part.add_header(
                    "Content-Disposition",
                    f"attachment; filename= {filename_text}",
            )

            msg.attach(part)
        except Exception as e:
                print(f"Oh no! We didn't found the attachment!n{e}")
                break

        try:
                # Creating a SMTP session | use 587 with TLS, 465 SSL and 25
                server = smtplib.SMTP(smtp_server, port)
                # Encrypts the email
                context = ssl.create_default_context()
                server.starttls(context=context)
                # We log in into our Google account
                server.login(login, password)
                # Sending email from sender, to receiver with the email body
                server.sendmail(sender_email, receiver_email, msg.as_string())
                print('Email sent!')
        except Exception as e:
                print(f'Oh no! Something bad happened!n{e}')
                break
        finally:
                print('Closing the server...')
                server.quit()
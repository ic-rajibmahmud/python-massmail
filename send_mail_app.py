import pandas as pd
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from string import Template
from email.mime.base import MIMEBase
from secrets import mycred
from email import encoders

file = pd.read_excel('addressbook.xlsx', sheet_name='Sheet1')
contacts = pd.DataFrame(file)


sender_mail =  input("Enter your email id: ")
sender_password = input("Please enter password : ")
smtp_server_add = input("Please enter smtp server address : ")
smtp_server_port = input("Please enter smtp server port (465 for gmail) : ")


for i in range (len(contacts)):
    name, email, address = contacts.iloc[i]
    port = smtp_server_port # ssl code 587 or 465
    smtp_server = smtp_server_add
    msg = MIMEMultipart()
    msg['from'] = 'Test Server"'
    msg['to'] = email
    msg['subject'] = f"Hi {name}"
    html = Template(Path('index_git.html').read_text())
    body = html.substitute(Participant = name)

    filename = 'test.pdf'
    attachment = open(filename, 'rb')

    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= " + filename)

    msg.attach(part)

    msg.attach(MIMEText(body,'html'))
    text = msg.as_string()
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(smtp_server, port, context = context) as server:

        server.login(sender_mail, sender_password)
        server.sendmail(sender_mail, msg['to'], text)
    print('Sent to: ', name)
print ("All done boss")

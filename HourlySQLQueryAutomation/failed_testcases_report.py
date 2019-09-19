import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from smtplib import SMTPException


def html_parser(data, subject):
    html= """
    <html>
    <head>
    <style> 
     table, th, td {{ border: 2px solid black; border-collapse: collapse; }}
      th, td {{ padding: 5px; }}
    </style>
    </head>
    <body>
    <p>{0}<br>{1}<p>
    </table>
    </body>
    </html>
    """.format(subject, data[0].to_html())
    return html

def send_email(html_data, subject):
    message = MIMEMultipart("alternative", None, [MIMEText(html_data,'html')])

    sender= 'murthybandaru57@gmail.com'
    receiver= 'murthybandaru57@gmail.com'
    server = 'smtp.gmail.com:587'
    password = 'mostwantedcutelad'
    message['Subject']= subject
    message['From']= sender
    message['To']= receiver
    try:
        server= smtplib.SMTP(server)
        server.ehlo()
        server.starttls()
        server.login(sender, password)
        server.sendmail(sender, receiver.split(','), message.as_string())
        server.quit()
        print ("Successfully sent email")
    except SMTPException:
        print ("Error: unable to send email")

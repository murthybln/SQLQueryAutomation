import smtplib

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from smtplib import SMTPException


def generate_html(data):
    html = """
    <html>
    <head>
    <p>{0}</p>
    <p>This is test email</p>
    </head>
    <body>
    </body>
    </html>
    """.format(data)
    return html


def send_email(html_data, subject, receiver, attachment=None):
    message = MIMEMultipart("alternative", None, [MIMEText(html_data, 'html')])

    sender = 'DevSupport@maverickcap.com'
    server = 'mavsmtp.mcap.off:25'
    message['Subject'] = subject
    message['From'] = sender
    message['To'] = receiver
    pdf_attachment = attachment
    if pdf_attachment:
        with open(pdf_attachment, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())

        encoders.encode_base64(part)

        part.add_header(
            "Content-Disposition",
            "attachment; filename= {filename}".format(filename=pdf_attachment.split('\\')[-1])
        )

        message.attach(part)

    try:
        server = smtplib.SMTP(server)
        server.sendmail(sender, receiver.split(','), message.as_string())
        server.quit()
        print ("Successfully sent email to {0}".format(receiver))
    except SMTPException as ex:
        print ("Error: unable to send email {}".format(ex))


def send_error_email(traceback, subject, receiver, job_info=None, sql_query=None):
    html_str = ""
    html_str += job_info if job_info else ""
    html_str += "<br/>"
    html_str += "<br/>".join(traceback.splitlines())
    html_str += "<br/>"
    html_str += "<br/>"
    html_str += sql_query if sql_query else ''
    html_str += "<br/>"
    message = MIMEMultipart("alternative", None, [MIMEText(html_str, 'html')])

    sender = 'DevSupport@maverickcap.com'
    server = 'mavsmtp.mcap.off:25'
    message['Subject'] = subject
    message['From'] = sender
    message['To'] = receiver
    try:
        server = smtplib.SMTP(server)
        server.sendmail(sender, receiver.split(','), message.as_string())
        server.quit()
        print ("Successfully sent email to {0}".format(receiver))
    except SMTPException:
        print ("Error: unable to send email")
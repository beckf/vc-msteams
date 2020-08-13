from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib


def send_mail_notification(send_from, send_to, subject, text, server=None):
    s = smtplib.SMTP(host=server)
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = send_to
    msg['Subject'] = subject
    msg.attach(MIMEText(text, 'plain'))
    s.sendmail(send_from, send_to, msg.as_string())
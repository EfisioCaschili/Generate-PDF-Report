import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from mimetypes import guess_type
from email.encoders import encode_base64
from getpass import getpass
from smtplib import SMTP
from datetime import datetime, timedelta
import re
import os
from pathlib import Path 
from dotenv import dotenv_values

dirname=os.path.dirname(__file__)
dirname="//192.168.1.125/gbts/Create_PDF_Report/" 

env=dotenv_values(dirname+'env.env')

base_subject = env.get('subject')
distr_list = env.get('distr_list')
distr_name = env.get('distr_name')
sender_name = env.get('sender_name')
sender_email = env.get('sender_email')
smtp_server = env.get('smtp_server')
password = env.get('password_email')
port = env.get('port')

class Email(object):
    def __init__(self, from_, to, subject, message, message_type='plain',
                 attachments=None):
        self.email = MIMEMultipart()
        self.email['From'] = from_
        self.email['To'] = to
        self.email['Subject'] = subject
        self.email['X-Priority'] = '1'  # High priority
        self.email['X-MSMail-Priority'] = 'High'
        self.email['Importance'] = 'High'
        text = MIMEText(message, message_type)
        self.email.attach(text)
    def __str__(self):
        return self.email.as_string()

class EmailConnection(object):
    def __init__(self, server, port, username, password):
        if ':' in server:
            data = server.split(':')
            self.server = data[0]
            self.port = int(data[1])
        else:
            self.server = server
            self.port = port
        self.username = username
        self.password = password
        self.connect()

    def connect(self):
        self.connection = SMTP(self.server, self.port)
        self.connection.ehlo()
        self.connection.starttls()
        self.connection.ehlo()
        self.connection.login(self.username, self.password)

    @staticmethod
    def get_email(email):
        if '<' in email:
            data = email.split('<')
            email = data[1].split('>')[0].strip()
        return email.strip()

    def send(self, message, from_=None, to=None,  cc=None):
        if isinstance(message, str):
            if from_ is None or to is None:
                raise ValueError('You need to specify `from_` and `to`')
            else:
                from_ = EmailConnection.get_email(from_)
                to = EmailConnection.get_email(to)
        else:
            from_ = message.email['From']
            to = message.email['To']
            to=to.split(';')
            toaddrs=to
            message = str(message)
        return self.connection.sendmail(from_, toaddrs, message)

    def close(self):
        self.connection.close()


def sender():
    print('Connecting to server...')
    server = EmailConnection(smtp_server, port, sender_email, password)
    print('Preparing the email...')
    opentext="Logbook SH has been not completed yet.<BR>Please, complete it as soon as possible.<BR>You can find the incosistencies here: //192.168.1.125/gbts/log.txt."
    email = Email(from_='"%s" <%s>' % (sender_name, sender_email),  # you can pass only email
                          #to='"%s" <%s>' % (distr_name, distr_list),  # you can pass only email
                          #to='%s' % (distr_list),  # you can pass only email
                          #cc= '%s' % (cc),  # you can pass only email
                          to=distr_list,
                          subject=base_subject, message=opentext, message_type='html')

    print('Sending...')
    server.send(email)
    print('Email sent...')

if __name__ == '__main__':
    sender()
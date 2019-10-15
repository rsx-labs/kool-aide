import smtplib
import email.utils
from email.mime.text import MIMEText
import getpass

# Prompt the user for connection info

servername = 'webmail.g07.fujitsu.local'
username = 'g07\\aide.notifications'
password = 'Retail.2019!'

# Create the message
msg = MIMEText("""This is a test message""")
sender = 'aide.notifications@ph.fujitsu.com'
recipients = 'm.trilles@ph.fujitsu.com'
msg['Subject'] = "AIDE Test Notification"
# msg['From'] = sender
# msg['To'] = recipients

server = smtplib.SMTP(servername, 25)
try:
    server.set_debuglevel(True)

    # identify ourselves, prompting server for supported features
    server.ehlo()

    # If we can encrypt this session, do it
    if server.has_extn('STARTTLS'):
        server.starttls()
        server.ehlo() # re-identify ourselves over TLS connection

    server.ehlo()
    server.login(username, password)
    server.sendmail('aide.notifications@ph.fujitsu.com', [recipients], msg.as_string())
finally:
    server.quit()

import email
import getpass, imaplib
import os
import sys
import io

detach_dir = '.'
if 'attachments' not in os.listdir(detach_dir):
    os.mkdir('attachments')

userName = 'fbhacker7373@gmail.com'
passwd = 'computer$@'

try:
    imapSession = imaplib.IMAP4_SSL('imap.gmail.com')
    typ, accountDetails = imapSession.login(userName, passwd)


    imapSession.select('INBOX')
    typ, data = imapSession.search(None, 'ALL')

    print("login")
    # Iterating over all emails
    msgId = data[0].split()[-1]

    print(msgId)
    typ, messageParts = imapSession.fetch(msgId, '(RFC822)')
    if typ != 'OK':
        print('Error fetching mail.')
        raise

    emailBody = messageParts[0][1]
    mail = email.message_from_bytes(emailBody)
    for part in mail.walk():
        if part.get_content_maintype() == 'multipart':
            # print part.as_string()
            continue
        if part.get('Content-Disposition') is None:
            # print part.as_string()
            continue
        fileName = part.get_filename()
        print(fileName)
        if bool(fileName):
            print(fileName)
            filePath = os.path.join(detach_dir, 'attachments', fileName)
            if not os.path.isfile(filePath):
                print
                fileName
                fp = open(filePath, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()
    imapSession.close()
    imapSession.logout()
except Exception as ex:
    print(ex)
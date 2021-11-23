# -*- coding: utf-8 -*-
"""
Created on Sun Nov 14 14:34:24 2021

@author: Pierrick

Basé sur le tuto :
    https://www.accomplishedafricanwomen.org/post/email-analysis-using-python-3-part-i-by-ogheneyoma-okobiah
    

"""

import pprint
import imaplib
import email
import getpass
import pandas as pd

### ETAPE 1 : Connexion ###

username =  input("Enter the email address: ")
password = getpass.getpass("Enter password: ")
mail = imaplib.IMAP4_SSL('imap.gmail.com')
mail.login(username, password)

### ETAPE 2 : Recherche ###

mail_lists = mail.list()
#mail.select("INBOX")
mail.select('"Mes dossiers/Immobilier"')

result, numbers = mail.search(None, 'ALL')
uids = numbers[0].split()
uids = [id.decode("utf-8") for id in uids ]
result, messages = mail.fetch(','.join(uids) ,'(RFC822)')

body_list =[]
date_list = []
from_list = [] 
subject_list = []
body = ""
for _, message in messages[::2]:
    email_message = email.message_from_bytes(message)
    email_subject = email.header.decode_header(email_message['Subject'])[0]
    subject_list.append(email_subject[0])
    for part in email_message.walk():
          if part.get_content_type() == "text/plain" :
              body = part.get_payload(decode=True)
              print(body)
              # body_list.append(body)
          else:
              continue
    body_list.append(body)
    body = ""
    date_list.append(email_message.get('date'))
    fromlist = email_message.get('From')
    fromlist = fromlist.split("<")[0].replace('"', '')
    from_list.append(fromlist)

### ETAPE 3 : Export dans un CSV ###

date_list = pd.to_datetime(date_list)
date_list = [item.isoformat(' ')[:-6]for item in date_list]
data = pd.DataFrame(data={'Date':date_list,'Sender':from_list,'Subject':subject_list, 'Body':body_list})
data.to_csv('emails.csv',index=False)
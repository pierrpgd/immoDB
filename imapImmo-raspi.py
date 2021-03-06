# -*- coding: utf-8 -*-
"""
Created on Sun Nov 14 14:34:24 2021

Mis à jour le 30/01/22

@author: Pierrick

Basé sur le tuto :
    https://www.accomplishedafricanwomen.org/post/email-analysis-using-python-3-part-i-by-ogheneyoma-okobiah
    

"""

def SeLogerOld(i):
    
    global nvResultat
    
    url = ''
    prix = ''
    nbPieces = ''
    surface = ''
    ville = ''
    codePostal = ''
    
    body = data.loc[i,'Body']
    
    if "annonces de vente d'appartements à" in data.loc[i,'Subject']:
        nbAnnonces = "multiple"
    else:
        nbAnnonces = "unique"
    
    body = body.replace('\t','')
    elements = pd.DataFrame(columns=['Contenu'])
    elements['Contenu'] = body.split('\r\n')
    
    # Boucle dans les éléments du corps de mail #
    
    if nbAnnonces == "multiple":
        
        for j in range(len(elements)):
            if (j > 3) and (elements.loc[j,'Contenu'] != '') and (elements.loc[j-1,'Contenu'] + elements.loc[j-2,'Contenu'] + elements.loc[j-3,'Contenu'] == '') and (elements.loc[j-4,'Contenu'] != ''):
                url = elements.loc[j,'Contenu'].replace(' ','').replace('-ALI168','')
            if '€' in elements.loc[j,'Contenu']:
                prix = re.sub(r'http\S+', '', elements.loc[j,'Contenu']).replace(' ','').replace('€','')
                prix = unidecode.unidecode(prix).replace(' ','')
            if '²' in elements.loc[j,'Contenu']:
                nbPieces = elements.loc[j,'Contenu'].split(' • ')[1].split(' ')[0]
                surface = elements.loc[j,'Contenu'].split(' • ')[2].split(' ')[0]
                lieu = elements.loc[j+1,'Contenu']
                ville = lieu.split(' (')[0].replace('ème','').replace('er','')
                codePostal = lieu.split(' (')[1].replace(')','')
                
            if (url != '') and (prix != '') and (nbPieces !='') and (surface != '') and (ville != '') and (codePostal !=''):
                annonce = {'Date':date,'Sender':sender,'Code postal':codePostal,'Ville':ville,'Surface (m²)':surface,'Nb pièces':nbPieces,'Prix (€)':prix,'URL':url}
                nvResultat = nvResultat.append(annonce, ignore_index=True)
                
                url = ''
                prix = ''
                nbPieces = ''
                surface = ''
                ville = ''
                codePostal = ''
                
    if nbAnnonces == 'unique':
        
        for j in range(len(elements)):
            if 'Appartement - ' in elements.loc[j,'Contenu']:
                prix = elements.loc[j,'Contenu'].split(' - ')[1]
                surface = elements.loc[j,'Contenu'].split(' - ')[2].replace(' m²','')
                surface = surface.replace(' ','0')
                lieu = elements.loc[j,'Contenu'].split(' - ')[3]
                ville = lieu.split(' (')[0].replace('ème','').replace('er','')
                codePostal = lieu.split(' (')[1].replace(')','')
            if prix + '  : ' in  elements.loc[j,'Contenu']:
                url =  elements.loc[j,'Contenu'].split(': ')[1].replace(' ','')
                prix = prix.replace(' ','').replace('€','')
                prix = unidecode.unidecode(prix).replace(' ','')
    
        if (url != '') and (prix != '') and (surface != '') and (ville != '') and (codePostal !=''):
            annonce = {'Date':date,'Sender':sender,'Code postal':codePostal,'Ville':ville,'Surface (m²)':surface,'Nb pièces':'0','Prix (€)':prix,'URL':url}
            nvResultat = nvResultat.append(annonce, ignore_index=True)
            
            url = ''
            prix = ''
            surface = ''
            ville = ''
            codePostal = ''

def SeLogerNew(i):
    
    global nvResultat
    
    url = ''
    prix = ''
    nbPieces = ''
    surface = ''
    ville = ''
    codePostal = ''
    
    body = data.loc[i,'Body']
    
    if type(body) == bytes:
        body = body.decode('utf-8')
    
    if "annonces de Vente d'appartements à" in data.loc[i,'Subject']:
        nbAnnonces = "multiple"
    else:
        nbAnnonces = "unique"
    
    body = body.replace('\t','')
    elements = pd.DataFrame(columns=['Contenu'])
    
    for h in range(len(unidecode.unidecode(body).split('\r\n'))):
        elements.loc[h,'Contenu'] = unidecode.unidecode(body).split('\r\n')[h]
    #elements['Contenu'] = body.split('\r\n')
    
    # Boucle dans les éléments du corps de mail #
    
    if nbAnnonces == "multiple":
        
        for j in range(len(elements)):
            #print(str(i) + ' ' + str(j))
# =============================================================================
#             if (j > 3) and (elements.loc[j,'Contenu'] != '') and (elements.loc[j-1,'Contenu'] + elements.loc[j-2,'Contenu'] + elements.loc[j-3,'Contenu'] == '') and (elements.loc[j-4,'Contenu'] != ''):
#                 url = elements.loc[j,'Contenu'].replace(' ','').replace('-ALI168','')
#                 print(url)
# =============================================================================
            if 'EUR(http' in elements.loc[j,'Contenu']:
                url = elements.loc[j,'Contenu'].split('(')[1].split(')')[0]
                prix = re.sub(r'http\S+', '', elements.loc[j,'Contenu']).replace(' ','')
                prix = prix.replace(' ','').replace('(','').replace('EUR','')
            if 'm2(http' in elements.loc[j,'Contenu']:
                nbPieces = elements.loc[j,'Contenu'].split(' * ')[1].split(' ')[0]
                surface = elements.loc[j,'Contenu'].split(' * ')[2].split(' ')[0]
                lieu = elements.loc[j+3,'Contenu']
                ville = lieu.split(' (')[0].replace('ème','').replace('er','')
                codePostal = lieu.split(' (')[1].replace(')','')
                
            if (url != '') and (prix != '') and (surface != '') and (ville != '') and (codePostal !=''):
                annonce = {'Date':date,'Sender':sender,'Code postal':codePostal,'Ville':ville,'Surface (m²)':surface,'Nb pièces':nbPieces,'Prix (€)':prix,'URL':url}
                nvResultat = nvResultat.append(annonce, ignore_index=True)
                
                url = ''
                prix = ''
                nbPieces = ''
                surface = ''
                ville = ''
                codePostal = ''
                
    if nbAnnonces == 'unique':
        
        for j in range(len(elements)):
            #print(str(i) + ' ' + str(j))
            if 'Annonce exclusive a' in elements.loc[j,'Contenu']:
                ville = elements.loc[j,'Contenu'].split(' ')
                if len(ville) == 6:
                    ville = elements.loc[j,'Contenu'].split(' ')[3] + ' ' + elements.loc[j,'Contenu'].split(' ')[4]
                else:
                    ville = elements.loc[j,'Contenu'].split(' ')[3]
            if 'EUR(http' in elements.loc[j,'Contenu']:
                url = elements.loc[j,'Contenu'].split('(')[1].split(')')[0]
                prix = re.sub(r'http\S+', '', elements.loc[j,'Contenu']).replace(' ','')
                prix = prix.replace(' ','').replace('(','').replace('EUR','')
                if not prix.isnumeric():
                    prix = ''
            if 'pieces(http' in elements.loc[j,'Contenu']:
                nbPieces = re.sub(r'http\S+', '', elements.loc[j,'Contenu']).replace(' ','').split('*')[0].split('piece')[0]
            if 'm2(http' in elements.loc[j,'Contenu']:
                surface = re.sub(r'http\S+', '', elements.loc[j,'Contenu']).replace(' ','').split('*')[1].replace('m2(','')
                
                
# =============================================================================
#             if 'Appartement - ' in elements.loc[j,'Contenu']:
#                 prix = elements.loc[j,'Contenu'].split(' - ')[1]
#                 surface = elements.loc[j,'Contenu'].split(' - ')[2].replace(' m²','')
#                 surface = surface.replace(' ','0')
#                 lieu = elements.loc[j,'Contenu'].split(' - ')[3]
#                 ville = lieu.split(' (')[0].replace('ème','').replace('er','')
#                 codePostal = lieu.split(' (')[1].replace(')','')
#             if prix + '  : ' in  elements.loc[j,'Contenu']:
#                 url =  elements.loc[j,'Contenu'].split(': ')[1].replace(' ','')
#                 prix = prix.replace(' ','').replace('€','')
#                 prix = unidecode.unidecode(prix).replace(' ','')
# =============================================================================
    
        if (url != '') and (prix != '') and (surface != '') and (ville != ''):
            annonce = {'Date':date,'Sender':sender,'Code postal':'','Ville':ville,'Surface (m²)':surface,'Nb pièces':nbPieces,'Prix (€)':prix,'URL':url}
            nvResultat = nvResultat.append(annonce, ignore_index=True)
            
            url = ''
            prix = ''
            surface = ''
            ville = ''
            codePostal = ''
            nbPieces = ''

def Century21(i):
    
    global nvResultat
    
    url = ''
    prix = ''
    nbPieces = ''
    surface = ''
    ville = ''
    codePostal = ''
    
    body = data.loc[i,'Body']
    elements = pd.DataFrame(columns=['Contenu'])
    elements['Contenu'] = body.split('**')
    
    # Boucle dans les éléments du corps de mail #
    
    for j in range(len(elements)):
        if 'Ref :' in elements.loc[j,'Contenu']:
            lieu = elements.loc[j-4,'Contenu']
            ville = lieu.split(' ')[1]
            codePostal = lieu.split(' ')[2]
            if '6900' in codePostal:
                ville = ville + ' ' + codePostal.replace('6900','')
            surface = elements.loc[j-2,'Contenu'].split(',')[0].replace(' m2','')
            surface = str(round(float(surface)))
            nbPieces = elements.loc[j-2,'Contenu'].split(',')[1].replace(' ','').replace('piece','').replace('s','')
            prix = elements.loc[j+2,'Contenu'].replace(' ','').replace('€','')
        
            annonce = {'Date':date,'Sender':sender,'Code postal':codePostal,'Ville':ville,'Surface (m²)':surface,'Nb pièces':nbPieces,'Prix (€)':prix,'URL':url}
            nvResultat = nvResultat.append(annonce, ignore_index=True)
            
            prix = ''
            nbPieces = ''
            surface = ''
            ville = ''
            codePostal = ''
            
def PAP(i):
    
    global nvResultat
    
    url = ''
    prix = ''
    nbPieces = ''
    surface = ''
    ville = ''
    codePostal = ''
    
    body = data.loc[i,'Body']
    elements = pd.DataFrame(columns=['Contenu'])
    if type(body) is str:
        elements['Contenu'] = body.split('\r\n')
    
    # Boucle dans les éléments du corps de mail #
    
    for j in range(len(elements)):
        if 'Une annonce correspondant à votre recherche' in elements.loc[j,'Contenu']:
            nbPieces = elements.loc[j+2,'Contenu'].split(' ')[2]
            ville = elements.loc[j+3,'Contenu'].replace("e",'').replace("r",'')
            surface = elements.loc[j+4,'Contenu'].replace(' m²','')
            prix = elements.loc[j+5,'Contenu'].replace(" EUR",'').replace('.','')
            
            annonce = {'Date':date,'Sender':sender,'Code postal':codePostal,'Ville':ville,'Surface (m²)':surface,'Nb pièces':nbPieces,'Prix (€)':prix,'URL':url}
            nvResultat = nvResultat.append(annonce, ignore_index=True)
            
            prix = ''
            nbPieces = ''
            surface = ''
            ville = ''

import imaplib
import email
import pandas as pd
import numpy as np
import re
import unidecode
import os

from tabulate import tabulate
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import mimetypes

### ETAPE 0 : Paramètres ###

dossier = "/home/pi/git-projects/Im-mail/"
fichier = 'annonces.csv'
cheminFichier = dossier + fichier

resultat_old = pd.read_csv(cheminFichier,sep=';',decimal=',')
derDate = resultat_old['Date'].max() #.split(' ')[0]

### ETAPE 1 : Connexion ###

username =  'pierrick.pagaud@gmail.com'
#password = getpass.getpass("Enter password for " + username + ": ")
password = 'ystwmjplhtvhrvya'
mail = imaplib.IMAP4_SSL('imap.gmail.com')
mail.login(username, password)

### ETAPE 2 : Recherche ###

mail_lists = mail.list()
#mail.select("INBOX")
mail.select('"Mes dossiers/Immobilier"')

result, numbers = mail.search(None, '(SINCE "' + pd.to_datetime(derDate).strftime("%d-%b-%Y") + '")') #'(SINCE "01-Jan-2010")'
uids = numbers[0].split()
uids = [id.decode("utf-8") for id in uids ]
result, messages = mail.fetch(','.join(uids) ,'(RFC822)')

### ETAPE 3 : Déchiffrage ###

body_list =[]
date_list = []
from_list = [] 
subject_list = []
body = ""
for _, message in messages[::2]:
    email_message = email.message_from_bytes(message)
    email_subject = email.header.decode_header(email_message['Subject'])[0]
    decode_type = email_subject[1]
    if decode_type:
        subject_list.append(email_subject[0].decode(decode_type))
    else:
        subject_list.append(email_subject[0])
    for part in email_message.walk():
          if part.get_content_type() == "text/plain" :
              body = part.get_payload(decode=True)
              if decode_type:
                  body = body.decode(decode_type)
                  #body = re.sub(r'http\S+', '', body)
          else:
              continue
    body_list.append(body)
    body = ""
    date_list.append(email_message.get('date'))
    fromlist = email_message.get('From')
    fromlist = fromlist.split("<")[0].replace('"', '').replace(" ","")
    from_list.append(fromlist)

date_list = pd.to_datetime(date_list)
date_list = [item.isoformat(' ')[:-6]for item in date_list]

while date_list[0] <= derDate:
    del date_list[0]
    del from_list[0]
    del subject_list[0]
    del body_list[0]
    
data = pd.DataFrame(data={'Date':date_list,'Sender':from_list,'Subject':subject_list, 'Body':body_list, 'Lieu':np.nan})

### ETAPE 4 : Décodage des corps de mail ###

resultat = pd.DataFrame(columns={'Date','Sender','Code postal','Ville','Prix m²','Surface (m²)','Nb pièces','Prix (€)','URL'})
nvResultat = pd.DataFrame(columns={'Date','Sender','Code postal','Ville','Prix m²','Surface (m²)','Nb pièces','Prix (€)','URL'})

# Boucle dans la liste de mails #

for i in range(len(data)):
    
    date = data.loc[i,'Date']
    sender = data.loc[i,'Sender']
    nvResultat = nvResultat.drop(nvResultat.index)
    
    # Traitement des mails #

    if sender == 'SeLoger':
        if date < '2021-12-13 12:23:09':
        
            SeLogerOld(i)
        
        else:
            
            SeLogerNew(i)
        
    if sender == 'CENTURY21':
        
        Century21(i)
        
# =============================================================================
#     if sender == 'AVendreALouer.fr':
#         
#     if sender == 'Logic-Immo':
# =============================================================================
        
    if sender == 'PAP.fr':
        
        PAP(i)
        
# =============================================================================
#     if sender == 'ParuVendu.fr':
# =============================================================================
    
    resultat = resultat.append(nvResultat, ignore_index=True)

### Sauvegarde des nouveaux résultats pour corps de mail ###

resultat = resultat[['Date','Sender','Code postal','Ville','Prix m²','Surface (m²)','Nb pièces','Prix (€)','URL']]
resultat = resultat.astype({"Surface (m²)":int,'Prix (€)':int})
resultat.loc[:,'Prix m²'] = round(resultat.loc[:,'Prix (€)'] / resultat.loc[:,'Surface (m²)'], 2)

data = resultat.values.tolist()

for j in data:
    del j[0]
    if int(j[6]) > 200001:
        del j[:]

data.insert(0, resultat.columns.tolist())
del data[0][0]

### ETAPE 5 : Formatage ###

resultat = resultat_old.append(resultat, ignore_index=True)
# =============================================================================
# resultat = resultat.astype({"Surface (m²)":int,'Prix (€)':int})
# resultat.loc[:,'Prix m²'] = round(resultat.loc[:,'Prix (€)'] / resultat.loc[:,'Surface (m²)'], 2)
# =============================================================================
#resultat = resultat.astype({'Prix m²':int})
resultat = resultat[['Date','Sender','Code postal','Ville','Prix m²','Surface (m²)','Nb pièces','Prix (€)','URL']]



### ETAPE 6 : Export dans un CSV ###

#data.to_csv('emails.csv',index=False,sep=";",decimal=',')
resultat.to_csv(cheminFichier,index=False,sep=";",decimal=',')

### ETAPE 7 : Envoi par mail ###

mimetypes.init()

text = """
Hello, Friend.

Here is your data:

{table}

Regards,

Me"""

html = """
<html><body><p>Hello, Friend.</p>
<p>Here is your data:</p>
{table}
<p>Regards,</p>
<p>Me</p>
</body></html>
"""

text = text.format(table=tabulate(data, headers="firstrow", tablefmt="grid"))
html = html.format(table=tabulate(data, headers="firstrow", tablefmt="html"))

#The mail addresses and password
sender_address = 'pierrick.pagaud@gmail.com'
sender_pass = 'Sandreas39*'
receiver_address = sender_address
#Setup the MIME
message = MIMEMultipart(
    "alternative", None, [MIMEText(text), MIMEText(html,'html')])
message['From'] = sender_address
message['To'] = receiver_address
message['Subject'] = 'Immap-Immo / Annonces du jour'
#The subject line
#The body and the attachments for the mail
#message.attach([MIMEText(text, 'plain'), MIMEText(html,'html')])
attach_file_name = 'annonces.csv'
#attach_file_name = cheminFichier
attach_file = open(cheminFichier, 'rb') # Open the file as binary mode

mimetype, _ = mimetypes.guess_type(cheminFichier)
if mimetype is None:
    mimetype = 'application/octet-stream'
type_, _, subtype = mimetype.partition('/')
payload = MIMEBase(type_, subtype, Name = attach_file_name)

#payload = MIMEBase('application', 'octate-stream')
payload.set_payload((attach_file).read())
encoders.encode_base64(payload) #encode the attachment
#add payload header with filename
payload.add_header('Content-Decomposition', 'attachment', filename=attach_file_name)
message.attach(payload)
#Create SMTP session for sending the mail
session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
session.starttls() #enable security
session.login(sender_address, password) #login with mail_id and password
text = message.as_string()
session.sendmail(sender_address, receiver_address, text)
session.quit()
print('Mail Sent')

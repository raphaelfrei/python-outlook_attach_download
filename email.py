from asyncio.windows_events import NULL
import win32com.client
import os
from datetime import datetime, timedelta

outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")

for account in mapi.Accounts:
    print(account.DeliveryStore.DisplayName)
    
inbox = mapi.GetDefaultFolder(6) #.Folders["Pasta"]
# 3  Deleted Items
# 4  Outbox
# 5  Sent Items
# 6  Inbox

#Define o diretório onde salva
outputDir = os.getcwd()

received_dt = NULL

backup = os.path.join(outputDir, 'backup.txt')

with open(backup, 'r') as f:
    received_dt = f.read()

if received_dt == "":
    received_dt = (datetime.now() - timedelta(days = 1) - timedelta(minutes = 10)).strftime('%m/%d/%Y %H:%M %p')

#Apenas faz o download dos emails deste remetente
email_sender = 'abc@gmail.com'

print(received_dt)

messages = inbox.Items
messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")

def DownloadAttachment():
    try:
        for message in list(messages):
            if message.SenderEmailAddress == email_sender and message.unread:
                try:                    
                    for attachment in message.Attachments:
                        attachment.SaveASFile(os.path.join(outputDir, attachment.FileName))
                        print(f"attachment {attachment.FileName} from {message.sender} saved")

                except Exception as e:
                    print("Error when saving the attachment:" + str(e)) 
            
    except Exception as e:
        print("Error when processing emails messages:" + str(e))

DownloadAttachment()
    

with open(backup, 'w') as f:
    received_dt = (datetime.now() - timedelta(minutes = 10)).strftime('%m/%d/%Y %H:%M %p')
    f.write(received_dt)
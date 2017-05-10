import win32com.client
import pythoncom
import os
import email
import re
import json
#import urllib2

count=1

class HandlerClass(object):

    def OnNewMailEx(self, receivedItemsIDs):
        for Id in receivedItemsIDs.split(","):
            mail = outlook.Session.GetItemFromID(Id)

            print("Subject:  "+ mail.Subject)
            Body = mail.Body

            #Body=str(mail.Body.encode('ascii', 'ignore'))
            #Actual_Body=''
            #for letter in Body:
                #if letter!='\r':
                #Actual_Body+=letter

            print("Body:  "+Body)
            global count
            if mail.Attachments:
                while (os.path.isdir('C:\\Users\\rnchris\\Desktop\\Attachments'+str(count))):
                    count+=1
                os.makedirs('C:\\Users\\rnchris\\Desktop\\Attachments'+str(count))
                #os.chdir('C:\\Users\\rnchris\\Desktop\\Attachments')
                for att in mail.Attachments:
                    attachmentType=att.FileName
                    attachmentType=attachmentType.split('.')[1]
                    att.SaveAsFile('C:\\Users\\rnchris\\Desktop\\Attachments'+str(count)+'\\Download'+str(count)+'.'+attachmentType)
                    count+=1
                    print('Worked!')
            #subject = mail.subject
            #command = re.search(r"%(.*?)%", subject).group(1)
            #print(command)


outlook = win32com.client.DispatchWithEvents("Outlook.Application", HandlerClass)

#outlook2=win32com.client.Dispatch("Outlook.Application").GetNameSpace('MAPI')
#checkAndDownLoadAttachments(outlook2)
pythoncom.PumpMessages()
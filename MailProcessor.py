import win32com.client
import pythoncom
import email
import re

class Handler_Class(object):

    def OnNewMailEx(self, receivedItemsIDs):
        for Id in receivedItemsIDs.split(","):
            mail = outlook.Session.GetItemFromID(Id)
            print("Subject:  "+ mail.Subject)
            Body = mail.Body
            #Body=str(mail.Body.encode('ascii', 'ignore'))
            Actual_Body=''
            for letter in Body:
                if letter!='\r':
                    Actual_Body+=letter



            print("Body:  "+Actual_Body)
            #subject = mail.subject
            #command = re.search(r"%(.*?)%", subject).group(1)
            #print(command)


outlook = win32com.client.DispatchWithEvents("Outlook.Application", Handler_Class)

pythoncom.PumpMessages()
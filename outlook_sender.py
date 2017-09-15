import win32com.client
from fuzzywuzzy import process
import datetime
import sys
import pickle
import os

pickl_loc = 'C:\\Users\\' + os.environ.get('USERNAME') + r'\Scripts\Data\senders.p'

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(5)
senders = {}
 
def get_greetime():
    currentTime = datetime.datetime.now()
    if currentTime.hour < 12:
        return 'Morning'
    elif 12 <= currentTime.hour < 17:
        return 'Afternoon'
    else:
        return 'Evening'
 
def email(recipient):
    body = "Good %s %s,<br><br><br><br>Regards,<br>YOUR NAME<br><font size='2'>Desk: YOUR PHONE<br>YOUR ROLE</font>"%(get_greetime(),recipient.split(',')[-1].strip())
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = ' '.join(sys.argv[2:]).title()
    newMail.HTMLBody = body
    newMail.To = recipient
    newMail.Display()
 
def add_to_dict(p):
    if p not in senders:
        senders[p] = 1
    else:
        senders[p] += 1
 
def update_send_list():
    for i in inbox.items:
        try:
            if len(i.To.split(';'))>1:
                for p in i.To.split(';'):
                    add_to_dict(p.strip())
            else:
                add_to_dict(i.To.strip())    
        except AttributeError:
            pass
        
    pickle.dump(senders, open(pickl_loc, 'wb'))
 
if __name__ == "__main__": 
    search_name = sys.argv[1]
    senders = pickle.load(open(pickl_loc,'rb'))
    all_matches = process.extract(search_name, senders.keys())
    good_matches = [x[0] for x in all_matches if x[1]>60 ]
    email(max(good_matches, key=senders.get))
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys("{DOWN}", 0)
    shell.SendKeys("{DOWN}", 0)
    update_send_list()


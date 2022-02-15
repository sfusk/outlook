import win32com.client, sys, datetime, re

outlook = win32com.client.Dispatch("Outlook.Application")
item = outlook.CreateItem(1)

if (len(sys.argv) == 3):
    if (re.match('^\d+$', sys.argv[2]) == None):
        print('number format error')
        exit()
    else:
        item.start = datetime.datetime.today() + datetime.timedelta(days=int(sys.argv[2]))
else:
    item.start = datetime.datetime.today()

item.subject = sys.argv[1]
item.allDayEvent = True
item.reminderSet = False
item.Save()
print(item.start.strftime('%Y/%m/%d') + ":" + item.subject + " is added")

import win32com.client, sys
import datetime

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
calender = outlook.GetDefaultFolder(9).items
items = sorted(calender, key=lambda calendar:calendar.start)
now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9)))
weekday_liset = ['月', '火', '水', '木', '金', '土', '日']
keyword = sys.argv[1]
extract_items = []

for item in items:
    if keyword in item.subject and item.start >= now:
	        extract_items.append(item)

for extract_item in extract_items:
    print(extract_item.subject + "：・" + extract_item.start.strftime("%Y/%m/%d(") + weekday_liset[extract_item.start.weekday()] + extract_item.start.strftime(") %H:%M") + "～" + extract_item.end.strftime("%H:%M"))

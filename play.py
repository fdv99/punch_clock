#from datetime import date, timedelta, datetime
import datetime
from datetime import date, timedelta
import time

lastMonthYear = (datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).year
lastMonth = (datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).month
currentMonthYear = datetime.datetime.today().year
currentMonth = datetime.datetime.today().month
nextMonthYear = (datetime.datetime.now().replace(day=28) + datetime.timedelta(days=4)).year
nextMonth = (datetime.datetime.now().replace(day=28) + datetime.timedelta(days=4)).month
day = datetime.datetime.today().day
date1 = (time.strftime("%Y-%m-%d"))
print(date1)
print(type(date1))
date2 = datetime.date.today()
print(date2)
print(type(date2))

if int(day) > 25:
    d1 = date(currentMonthYear, currentMonth, 26)  # start date
    d2 = date(nextMonthYear, nextMonth, 25)  # end date
    delta = d2 - d1         # timedelta
    for i in range(delta.days + 1):
        print(d1 + timedelta(i))
        #ws.cell(row=i + 3, column=1, value=(d1 + timedelta(i)))
else:
    d1 = date(lastMonthYear, lastMonth, 26)  # start date
    d2 = date(currentMonthYear, currentMonth, 25)  # end date
    delta = d2 - d1         # timedelta
    for i in range(delta.days + 1):
        print(d1 + timedelta(i))
        #ws.cell(row=i + 3, column=1, value=(d1 + timedelta(i)))

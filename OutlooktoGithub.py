from win32com.client import Dispatch
from tabulate import tabulate
import datetime
import pdb

from icalendar import Calendar, Event
from icalendar import vDatetime
from datetime import datetime as dt
import tempfile, os,time
import pytz
from git import Repo

OUTLOOK_FORMAT = '%m/%d/%Y %H:%M'
outlook = Dispatch("Outlook.Application")

ns = outlook.GetNamespace("MAPI")

appointments = ns.GetDefaultFolder(9).Items 

begin = datetime.date.today()
end = begin + datetime.timedelta(days=30);
restriction = "[Start] >= '" + begin.strftime("%m/%d/%Y") + "' AND [End] <= '" +end.strftime("%m/%d/%Y") + "'"

appointments.IncludeRecurrences = True
appointments.Sort("[Start]")
restrictedItems = appointments.Restrict(restriction)

cal = Calendar()
cal.add('prodid', '-//Microsoft Corporation//Outlook 16.0 MIMEDIR//EN')
cal.add('version', '2.0')
cal.add('METHOD', 'PUBLISH')
cal.add('X-MS-OLK-FORCEINSPECTOROPEN', 'TRUE')

# directory ="D:\Me_Data\py\Calendar"
directory =r"D:\NeilAuto\cal"
f = open(os.path.join(directory, 'Cneil.ics'), 'wb')

for restrictedItem in restrictedItems:
    event = Event()
    event.add('CLASS','PUBLIC')
    event.add('DTEND',restrictedItem.End)
    event.add('DTSTART',restrictedItem.Start)
    event.add('summary',restrictedItem.Subject)
    event.add('DESCRIPTION', restrictedItem.Body)
    event.add('LOCATION',restrictedItem.Location)
    event.add('X-MICROSOFT-CDO-BUSYSTATUS','BUSY')
    cal.add_component(event)
f.write(cal.to_ical())
f.close()
print("文件保存OK!")

# mydir =r"D:\Me_Data\OTH\PortableGit\mingw64\bin"
# 给 MYDIR赋值(临时创建的环境变量)
# os.environ["PATH"] = mydir+";"+os.environ["PATH"]
# 

path1= r"D:\NeilAuto\cal"
repo = Repo(path1)
index = repo.index
index.add(['Cneil.ics'])
index.commit(str(dt.now()))
remote = repo.remote()
remote.push()
print("Successful push to MyGit!")

time.sleep(3)
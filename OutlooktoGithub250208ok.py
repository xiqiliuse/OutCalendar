
from win32com.client import Dispatch
from icalendar import Calendar, Event
from pywinauto import Application
from multiprocessing import Process
# from datetime import datetime as dt
from git import Repo
import datetime
import os
import time

def get_outlook_window():
    app = Application(backend="uia").connect(path="OUTLOOK.EXE")
    outlook_main = app.window(title_re=".*Outlook.*")
    # outlook_main.print_control_identifiers()
    waittime=20
    try:
        fw_CheckBox = outlook_main.child_window(title="允许访问(A)", auto_id="4771", control_type="CheckBox").wait("visible", timeout=waittime)
        allow_button = outlook_main.child_window(title="允许", auto_id="4774", control_type="Button").wait("visible", timeout=waittime)
        # print("找到允许按钮")
        # 获取 CheckBox 的状态
        state = fw_CheckBox.get_toggle_state()
        print(f"CheckBox 状态: {state}")  # 0 - 未选中, 1 - 选中, 2 - 不确定
        
        # 根据状态进行操作
        if state == 0:
            fw_CheckBox.click()
            print("CheckBox 已选中")
        else:
            print("CheckBox 已取消选中")

        if allow_button.is_enabled():
            allow_button.click()
            allow_button.click_input()
            print("允许按钮已点击")
        else:
            print("允许按钮不可点击")
        return 
    except Exception as e:
        print(f"未找到允许按钮: {e}")
        return None
def downloadoutlook():
    # OUTLOOK_FORMAT = '%m/%d/%Y %H:%M'
    outlook = Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    appointments = ns.GetDefaultFolder(9).Items 
    begin = datetime.date.today()
    end = begin + datetime.timedelta(days=30) 
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
    global foldername
    f = open(os.path.join(directory, 'Cneilnew.ics'), 'wb')
    time.sleep(1)
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
    return
if __name__ == "__main__":
    # foldername='Cneilnew.ics'

    p1 = Process(target=get_outlook_window)
    p2 = Process(target=downloadoutlook)
    p1.start()
    p2.start()
    p1.join()
    p2.join()

    repo = Repo(r"D:\NeilAuto\cal")
    index = repo.index
    index.add(['Cneilnew.ics'])
    index.commit(str(datetime.datetime.now()))
    remote = repo.remote()
    remote.push()
    print("Successful push to MyGit!")

    print("所有功能执行完毕！")
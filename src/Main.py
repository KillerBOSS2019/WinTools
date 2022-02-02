from typing import Type
import TouchPortalAPI
from numpy import clip
import pycaw
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import threading
import pyautogui
from time import sleep
from TouchPortalAPI import TYPES
import subprocess
import json
import sys
import pythoncom
# import extract_icon
import ctypes
import psutil
import win32process
import pygetwindow
import os


### Gitago Imports
import mss
import mss.tools
from io import BytesIO
from PIL import Image
import win32clipboard
import win32ui
import win32gui
import win32con 
import time
import schedule
import dateutil.relativedelta
from datetime import datetime
from wintoast import ToastNotifier
import wintoast
from winotify import Notification, audio


### 
# Origin Launcher Window = OriginWebHelperService
# Steam Launcher = Steam vguiPopupWindow
# Epic Launcher = Epic Games Launcher
# Discord = Discord Chrome_WidgetWin_1   ???   MIGHT NOT NEED THIS  shows other stuff before.


TPClient = TouchPortalAPI.Client('Windows-Tools')




def win_toast(atitle="", amsg="", aduration="short", icon=""):
    ### setting the base notification stuff
    if not os.path.exists(rf"{icon}") or icon == "": icon = os.path.join(os.getcwd(),"src\icon.png")
    print(icon)
    toast = Notification(app_id="WinTools",
                         title=atitle,
                         msg=amsg,
                         icon=rf"{icon}",
                         duration = aduration.lower(),
                         )
    
    if False:
        toast.add_actions(label=button, 
                        link=alink)

    if True:
        audioDic = {
            "Default": audio.Default,
            "IM": audio.IM,
            "Mail": audio.Mail,
            "Reminder": audio.Mail,
            "SMS": audio.SMS,
            "LoopingAlarm1": audio.LoopingAlarm,
            "LoopingAlarm2": audio.LoopingAlarm2,
            "LoopingAlarm3": audio.LoopingAlarm3,
            "LoopingAlarm4": audio.LoopingAlarm4,
            "LoopingAlarm5": audio.LoopingAlarm6,
            "LoopingAlarm6": audio.LoopingAlarm8,
            "LoopingAlarm7": audio.LoopingAlarm9,
            "LoopingAlarm8": audio.LoopingAlarm10,
            "LoopingCall1": audio.LoopingCall,
            "LoopingCall2": audio.LoopingCall,
            "LoopingCall3": audio.LoopingCall,
            "LoopingCall4": audio.LoopingCall,
            "LoopingCall5": audio.LoopingCall,
            "LoopingCall6": audio.LoopingCall,
            "LoopingCall7": audio.LoopingCall,
            "LoopingCall8": audio.LoopingCall,
            "LoopingCall9": audio.LoopingCall,
            "LoopingCall10": audio.LoopingCall,
            "Silent": audio.Silent,
        }
        #if theaudio =="Default":
        toast.set_audio(audioDic["Default"], loop=False)
            
            
    toast.build().show()
    
    


def run_continuously(interval=1):
    """Continuously run, while executing pending jobs at each
    elapsed time interval.
    @return cease_continuous_run: threading. Event which can
    be set to cease continuous run. Please note that it is
    *intended behavior that run_continuously() does not run
    missed jobs*. For example, if you've registered a job that
    should run every minute and you set a continuous run
    interval of one hour then your job won't be run 60 times
    at each interval but only once.
    """
    cease_continuous_run = threading.Event()

    class ScheduleThread(threading.Thread):
        @classmethod
        def run(cls):
            while not cease_continuous_run.is_set():
                schedule.run_pending()
                time.sleep(interval)

    continuous_thread = ScheduleThread()
    continuous_thread.start()
    return cease_continuous_run


def magnifier(action):
    if action == "Zoom In":
        pyautogui.hotkey('win', '=')
    if action == "Zoom Out":
        pyautogui.hotkey('win', '-')
    if action == "Exit":
        pyautogui.hotkey('win', 'escape')
        
def winextra(action):
  # if action == "Keep Active, Minimize All (Toggle)":
  #     pyautogui.hotkey('win', 'home')
  #     
  # if action == "Minimize All (Toggle)":
  #         pyautogui.hotkey('win', 'd')
  #         
    if action == "Emoji":
        pyautogui.hotkey('win', '.')
        
    if action == "Keyboard":
        pyautogui.hotkey('win', 'ctrl', "o")
        

def get_app_icon():
    import win32api
    active_path = getActiveExecutablePath()
    print(active_path)
 
    if active_path == "C:\Windows\System32\ApplicationFrameHost.exe" or None:
       pass
    else:
       ico_x = win32api.GetSystemMetrics(win32con.SM_CXICON)
       ico_y = win32api.GetSystemMetrics(win32con.SM_CYICON)
       try:
           large, small = win32gui.ExtractIconEx(getActiveExecutablePath(),0)
           win32gui.DestroyIcon(small[0])

           hdc = win32ui.CreateDCFromHandle( win32gui.GetDC(0) )
           hbmp = win32ui.CreateBitmap()
           hbmp.CreateCompatibleBitmap( hdc, ico_x, ico_x )
           hdc = hdc.CreateCompatibleDC()

           hdc.SelectObject( hbmp )
           hdc.DrawIcon( (0,0), large[0] )

           hbmp.SaveBitmapFile( hdc, r'C:\Users\dbcoo\AppData\Roaming\TouchPortal\plugins\WinTools\tmp\newwwwicon.bmp')  
           time.sleep(2)
       except IndexError as err:
            print(err)
            time.sleep(5)
            #continue

### gotta get this to go to icon, then have to get it to nt update less its new...
##   schedule.every(1).seconds.do(get_app_icon)

from pyvda import AppView, get_apps_by_z_order, VirtualDesktop, get_virtual_desktops
def virtual_desktop(target_desktop=None, move=False, pinned=False):
    number_of_active_desktops = len(get_virtual_desktops())
    print(f"There are {number_of_active_desktops} active desktops")

    current_desktop = VirtualDesktop.current()
    
    if target_desktop == "Next":
        if move:
            target_desktop2 = current_desktop.number + 1
            current_window = AppView.current()
            target_desktop = VirtualDesktop(target_desktop2)
            current_window.move(target_desktop) 
            print(f"did it move to {target_desktop2} ?")
            if pinned:
                AppView.current().pin()
        
        elif not move:
            target_desktop2 = current_desktop.number + 1
            if target_desktop2 <= number_of_active_desktops:
                VirtualDesktop(target_desktop2).go()
            else:
                print("too many")
        
    elif target_desktop == "Previous":
        if move:
            target_desktop2 = current_desktop.number - 1
            current_window = AppView.current()
            target_desktop = VirtualDesktop(target_desktop2)
            current_window.move(target_desktop)
            if pinned:
                AppView.current().pin()
        
        elif not move:
            target_desktop2 = current_desktop.number - 1
            if target_desktop2 > 0:
                VirtualDesktop(target_desktop2).go()
            else:
                pass
            
    elif target_desktop is not "Previous" or "Next":  ## else if 0, 1, 2 3, etc.. anything but next or previous
        print("THE TARGETED DESKTOP IS", target_desktop)
        if move:
            target_desktop2 = int(target_desktop)
            print("The Target desktop secondary thing is", target_desktop2)
            current_window = AppView.current()
            target_desktop = VirtualDesktop(int(target_desktop2))
            current_window.move(target_desktop)
         # current_window = AppView.current()
         # target_desktop2 = VirtualDesktop(target_desktop)
         # current_window.move(target_desktop)
         # if pinned:
         #     AppView.current().pin()
            print(" hmmm")
        
      # elif not move:
      #     target_desktop2 = current_desktop.number - 1
      #     if target_desktop2 > 0:
      #         VirtualDesktop(target_desktop2).go()
      #     else:
      #         pass
      # 
        
    #print(f"Current desktop is number {current_desktop.number}")
   ##if move:
   ##    if target_desktop =="Next":
   ##        target_desktop2 = current_desktop.number + 1
   ##        if target_desktop2 <= number_of_active_desktops:
   ##                current_window = AppView.current()
   ##                target_desktop = VirtualDesktop(move_to)
   ##                current_window.move(target_desktop)
   ##                
   ##    elif target_desktop == "Previous":
   ##        target_desktop2 = current_desktop.number - 1
   ##        if target_desktop2 > 0:
   ##            current_window = AppView.current()
   ##            target_desktop = VirtualDesktop(target_desktop2)
   ##            current_window.move(target_desktop)       
   ##            
   ##    current_window = AppView.current()
   ##    target_desktop = VirtualDesktop(move_to)
   ##    current_window.move(target_desktop)
   ##    print(f"Moved window {current_window.hwnd} to {target_desktop.number}")
        
    ### what does this do?
    #print("Pinning the current window")
    #AppView.current().pin()

def vd_check():
    vdlist=[]
    virtual_desk_count = len(get_virtual_desktops())
    vdlist.append("Next")
    vdlist.append("Previous")
    for i in range (virtual_desk_count):
        vdlist.append(str(i))
    print(vdlist)
    TPClient.choiceUpdate("KillerBOSS.TP.Plugins.virtualdesktop.actionchoice", vdlist)
    print("choice updated?")



def get_size(bytes, suffix="B"):
    """
    Scale bytes to its proper format
    e.g:
        1253656 => '1.20MB'
        1253656678 => '1.17GB'
    """
    factor = 1024
    for unit in ["", "K", "M", "G", "T", "P"]:
        if bytes < factor:
            return f"{bytes:.2f}{unit}{suffix}"
        bytes /= factor
        

def getDriveName(driveletter):
    return subprocess.check_output(["cmd","/c vol "+driveletter]).decode().split("\r\n")[0]

## NOT TO BE CONFUSED WITH UPLOAD / DOWNLOAD SPEED
def network_usage():
    ### this gets ran by disk_usage
    adict = {}
    net_io = psutil.net_io_counters()
    adict['sent'] = get_size(net_io.bytes_sent)
    adict['received'] = get_size(net_io.bytes_recv)
    return adict


def disk_usage(drives=False):
    # Disk Information
    # get all disk partitions
    try:
        partitions = psutil.disk_partitions()

        for partition in partitions:
            the_partition = partition.device.split(":")
            driveletter = (the_partition[0])       
            if not drives:
               #print(f"=== Device: {partition.device} ===")
               #print("NOT DRIVES")
                # print(f"  Mountpoint: {partition.mountpoint}")
                #print(f"  File system type: {partition.fstype}")
                the_partition = partition.device.split(":")
                drive_name = getDriveName(the_partition[:1][0]+":")

                drive_name_replaced = drive_name.replace(f"Volume in drive {the_partition[:1][0]} is", "")
                if drive_name.endswith("has no label."):
                    drive_name_replaced = partition.mountpoint
                try:
                    partition_usage = psutil.disk_usage(partition.mountpoint)
                except PermissionError:
                            # this can be catched due to the disk that
                            # isn't ready
                    continue
                freespace = get_size(partition_usage.free).replace("GB","")
                usedspace = get_size(partition_usage.used).replace("GB","")


                TPClient.createStateMany([
                {
                    "id": f'KillerBOSS.TP.Plugins.Windows.drive.letter_{driveletter}',
                    "desc": f"{driveletter} Drive: Name",
                    "value": ""
                },
                {
                    "id": f'KillerBOSS.TP.Plugins.Windows.drive.size_{driveletter}',
                    "desc": f"{driveletter} Drive: Total Space",
                    "value": ""
                },
                {
                    "id": f'KillerBOSS.TP.Plugins.Windows.drive.free_{driveletter}',
                    "desc": f"{driveletter} Drive: Free Space",
                    "value": ""
                },
                {
                    "id": f'KillerBOSS.TP.Plugins.Windows.drive.percent_{driveletter}',
                    "desc": f"{driveletter} Drive: Used",
                    "value": ""
                },
                {
                    "id": f'KillerBOSS.TP.Plugins.Windows.drive.percent_{driveletter}',
                    "desc": f"{driveletter} Drive: Percentage",
                    "value": ""
                },
                ])

                percentage = 100 - partition_usage.percent
                str_percent = str(percentage)
                str_percent = str_percent[0:4]

                TPClient.stateUpdateMany([
                {
                    "id": f'KillerBOSS.TP.Plugins.Windows.drive.letter_{driveletter}',
                    "value": drive_name_replaced
                },
                {
                    "id": f'KillerBOSS.TP.Plugins.Windows.drive.size_{driveletter}',
                    "value": get_size(partition_usage.total)
                },
                {
                    "id": f'KillerBOSS.TP.Plugins.Windows.drive.free_{driveletter}',
                    "value": freespace
                },
                {
                    "id": f'KillerBOSS.TP.Plugins.Windows.drive.used_{driveletter}',
                    "value": usedspace
                },
                {
                    "id": f'KillerBOSS.TP.Plugins.Windows.drive.percent_{driveletter}',
                    "value": str_percent
                },
                ])
    except:
        pass           
     
     ### Total Read/Write since boot       
    # get IO statistics since boot
    disk_io = psutil.disk_io_counters()
    print(f"Total read: {get_size(disk_io.read_bytes)}")
    print(f"Total write: {get_size(disk_io.write_bytes)}")
   
    network = network_usage()
    print(f"Total Bytes Sent: {network['received']}")
    print(f"Total Bytes Received: {network['sent']}")
    TPClient.stateUpdateMany([
    {
        "id": f'KillerBOSS.TP.Plugins.windows.network.sent',
        "value": network['sent']
    },
    {
        "id": f'KillerBOSS.TP.Plugins.windows.network.received',
        "value": network['received']
    },
    {
        "id": f'KillerBOSS.TP.Plugins.windows.disk.read',
        "value": get_size(disk_io.read_bytes)
    },
    {
        "id": f'KillerBOSS.TP.Plugins.windows.disk.write',
        "value": get_size(disk_io.write_bytes)
    },
    ])
## we could let user limit the drives we get details from... but should we bother?
#disk_usage(drives=["C"])


def pc_uptime():
    rd = time_booted()    
    hours = rd.hours
    minutes = rd.minutes
    seconds = rd.seconds

    while True:
        seconds = int(seconds)
        minutes = int(minutes)
        hours = int(hours)
        if seconds == 59:
            seconds =  seconds = 0
            minutes = minutes + 1
        elif seconds < 59:
            seconds = seconds + 1   
        if minutes == 60:
           minutes = minutes = 0
           hours = hours + 1
           
 ### makin it look "pretty"      
        if minutes <10:
            minutes = "0" + str(minutes)
        if seconds <10:
            seconds = "0" + str(seconds)
        time.sleep(1)
        
        pc_live_time = (f"{hours}:{minutes}:{seconds}")
        TPClient.stateUpdate("KillerBOSS.TP.Plugins.Windows.livetime", str(pc_live_time))
        print(pc_live_time)
    

### Find when PC was booted
previous_time = ""
def time_booted():
    global previous_time
    boot_time_timestamp = psutil.boot_time()
    # bt = datetime.fromtimestamp(boot_time_timestamp)
    current = time.time()
    dt1 = datetime.fromtimestamp(boot_time_timestamp)
    dt2 = datetime.fromtimestamp(current) 
    rd = dateutil.relativedelta.relativedelta (dt2, dt1)
    return rd


###   ### Find when PC was booted
###   previous_time = ""
###   def time_booted():
###       global previous_time
###       boot_time_timestamp = psutil.boot_time()
###       bt = datetime.fromtimestamp(boot_time_timestamp)
###       #print(f"Boot Time: {bt.year}/{bt.month}/{bt.day} {bt.hour}:{bt.minute}")
###       
###       ## how to get unix timestamp from datetime instead of time module?
###       current = time.time()
###       dt1 = datetime.fromtimestamp(boot_time_timestamp)
###       dt2 = datetime.fromtimestamp(current) 
###       rd = dateutil.relativedelta.relativedelta (dt2, dt1)
###       #return rd
###       
###       
###       
###       if rd.minutes <10:
###           rd.minutes = "0" + str(rd.minutes)
###       if rd.seconds <10:
###           rd.seconds = "0" + str(rd.seconds)
###       
###       if rd.seconds == previous_time:
###           pass
###       else:
###           previous_time = rd.seconds
###           print(f"PC LIVE TIME: {rd.hours}:{rd.minutes}:{rd.seconds}")  
###           pc_live_time = (f"{rd.hours}:{rd.minutes}:{rd.seconds}")
###           #TPClient.createState("KillerBOSS.TP.Plugins.Windows.livetime", "Windows Live Time", "")
###           TPClient.stateUpdate("KillerBOSS.TP.Plugins.Windows.livetime", str(pc_live_time))
###           return(f"{rd.hours}:{rd.minutes}:{rd.seconds}")






def check_process(process_name, shortcut ="", focus=True, focus_type="Restore"):
    exist = False
    processes = []
    win32gui.EnumWindows(lambda x, _: processes.append(x), None)
    
    for hwnd in processes:
        window_name = win32gui.GetWindowText(hwnd)
        class_name = win32gui.GetClassName(hwnd)
        if process_name.lower() in window_name.lower():
            exist = True
            print(window_name, class_name, hwnd)
            
            
            if focus:
                print("attempting to focus window")
                #SW_SHOWMAXIMIZED,  SW_RESTORE, SW_SHOWNOACTIVATE, SW_SHOWNORMAL, Minimize 
                #### How to SHOW but not bring to front? certain times windows wont capture cause they arent in some sort of focus...
                if focus_type == "Normal":                 ### difference between SHOWNORMAL and NORMAL  ??  or SHOW_OPENWINDOW ?
                    win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL) 
                    win32gui.SetForegroundWindow(hwnd)
                if focus_type == "Maximized":
                    win32gui.ShowWindow(hwnd, win32con.SW_SHOWMAXIMIZED)  
                    win32gui.SetForegroundWindow(hwnd)
                if focus_type == "Restore":
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)  
                    win32gui.SetForegroundWindow(hwnd)
                if focus_type == "Minimized":
                    win32gui.ShowWindow(hwnd, win32con.SW_SHOWMINIMIZED)  
                    win32gui.SetForegroundWindow(hwnd)
                break
    if not exist:
        try:
            print("load via shortcut")
            os.system('"' + shortcut + '"')
        except:
            pass
    
#check_process("Discord", shortcut=r"C:\Users\dbcoo\AppData\Local\Discord\Update.exe --processStart Discord.exe", focus=True)


def get_windows():
    results = []
    def winEnumHandler(hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetWindowText(hwnd):
    
                results.append(win32gui.GetWindowText(hwnd))
                
    win32gui.EnumWindows(winEnumHandler, None)
    return results

old_results = []
def get_windows_update():
    global windows_active, old_results
    windows_active = get_windows()
    print("triggered get_windows")
    if len(old_results) is not len(windows_active):
        # windows_active = get_windows()
        print("Previous Count:", len(old_results), "New Count:", len(windows_active))
        old_results = windows_active
        TPClient.choiceUpdate("KillerBOSS.TP.Plugins.screencapture.window_name", windows_active)
        TPClient.stateUpdate("KillerBOSS.TP.Plugins.Windows.activeCOUNT", str(len(windows_active)))
    else:
        windows_active = get_windows()
        old_results = windows_active


def copy_im_to_clipboard(image):
    bio = BytesIO()
    image.save(bio, 'BMP')
    data = bio.getvalue()[14:] # removing some headers
    bio.close()
    send_to_clipboard(win32clipboard.CF_DIB, data)


def send_to_clipboard(clip_type, data):
    if clip_type == "text":
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(data)
        win32clipboard.CloseClipboard()
    else:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(clip_type, data)
        win32clipboard.CloseClipboard()

## potentially not needed anymore
def file_to_bytes(filepath):
    ## Take image into bytes and onto clipboard
    image = Image.open(filepath)
    output = BytesIO()
    image.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()
    
    ### Sending to Clipboard
    send_to_clipboard(win32clipboard.CF_DIB, data)
    ### Deleting Temp File
    os.remove(filepath)
    print("Temp Image Deleted")


monitor_count_old = ""
def check_number_of_monitors():
    print("triggered get num of monitors")
    global monitor_count_old
    with mss.mss() as sct:
        monitor_count = (len(sct.monitors))
        
    monitor_list = []
    if str(monitor_count_old) == str(monitor_count):
        pass
    elif str(monitor_count_old) is not str(monitor_count):
        monitor_count_old = monitor_count
        for monitor_number in range(monitor_count):
            if monitor_number == 0:
                monitor_list.append(str(monitor_number))
            else:
                monitor_list.append(str(monitor_number))
                
        TPClient.choiceUpdate("KillerBOSS.TP.Plugins.screencapture.monitors", monitor_list)
        ### would love to be able to capture the monitor name but cannot find the method.. need to use user32.dll ??  - https://discord.com/channels/@me/786771528381104178/934685222577004545
        return monitor_count
        

###screenshot window without bringing it to foreground 
def screenshot_window(capture_type, window_title=None, clipboard=False, save_location=None):
    from ctypes import windll
    hwnd = win32gui.FindWindow(None, window_title)
    try:
        left, top, right, bot = win32gui.GetClientRect(hwnd)
        #left, top, right, bot = win32gui.GetWindowRect(hwnd)
        w = right - left
        h = bot - top
        
        hwndDC = win32gui.GetWindowDC(hwnd)
        mfcDC  = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()
        
        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)
        saveDC.SelectObject(saveBitMap)
        
        # Change the line below depending on whether you want the whole window
        # or just the client area as shown above. 
                              # 1, 2, 3 all give different results   ( 3 seems to work for everything)
        result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), capture_type)
        bmpinfo = saveBitMap.GetInfo()
        bmpstr = saveBitMap.GetBitmapBits(True)
        
        im = Image.frombuffer(
            'RGB',
            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
            bmpstr, 'raw', 'BGRX', 0, 1)
        
        win32gui.DeleteObject(saveBitMap.GetHandle())
        saveDC.DeleteDC()
        mfcDC.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwndDC)
        
        if result == 1:
            #PrintWindow Succeeded
            if clipboard == True:
                copy_im_to_clipboard(im)
                print("Copied to Clipboard")
            elif clipboard == False:
                im.save(save_location+".png")
                ### very bad and ugly compression
               ### im.save(save_location+"_Compressed_.png", 
               ###     "JPEG", 
               ###     optimize = True, 
               ###     quality = 10)
                print("Saved to Folder")
    except:
        pass


def screenshot_monitor(monitor_number, filename="", clipboard = False):    
    with mss.mss() as sct:
        try:
            mon = sct.monitors[monitor_number]  
            # Capturing Entire Monitor
            monitor = {
                "top": mon["top"],
                "left": mon["left"],
                "width": mon["width"],
                "height": mon["height"],
                "mon": monitor_number,
            }
            
            # Grab the Image
            sct_img = sct.grab(monitor)     
            
            if clipboard == True:
                if monitor_number==0:
                    print("capturing all, using temp file")  ## having to use temp file to capture all screens successfully?
                    mss.tools.to_png(sct_img.rgb, sct_img.size, output="temp.png")
                    file_to_bytes("temp.png") ## Saved to Clipboard
                    
                 ##   Instead of making a temp file we get it direct from raw to clipboard
                elif monitor_number != 0:
                    img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
                    copy_im_to_clipboard(img)

            if clipboard == False:
                mss.tools.to_png(sct_img.rgb, sct_img.size, output=filename + ".png")
                print("Image saved -> "+ filename+ ".png" )

        except IndexError:
            print("This Monitor does not exist")

#screenshot_monitor(3, "test_file", clipboard = True)



### HOW TO SET THESE WITH


#### Need a SETTING so user can set custom refresh increments, and or turn them off all together.
### tried hard to get into settings but it gave me non stop issues.. i give up for now..

## Computer Up-Time Check
##   schedule.every(1).seconds.do(time_booted)
##   ## Gets Active windows and updates stuff if they are different
##   schedule.every(30).seconds.do(get_windows_update)
##   ## Virtual Desktop Check
##   schedule.every(5).minutes.do(vd_check)
##   ## Check Number of Monitors
##   schedule.every(5).minutes.do(check_number_of_monitors)
##   ## Disk Usage Check
##   schedule.every(55).seconds.do(disk_usage)



# Start the background thread  this makes the schedules run
stop_run_continuously = run_continuously()


#### END OF SCREEN CAPTURE STUFF ####

class AudioController(object):
    def __init__(self, process_name):
        self.process_name = process_name
        self.volume = self.process_volume()

    def mute(self):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                interface.SetMute(1, None)
                print(self.process_name, 'has been muted.')  # debug

    def unmute(self):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                interface.SetMute(0, None)
                print(self.process_name, 'has been unmuted.')  # debug

    def process_volume(self):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                #print('Volume:', interface.GetMasterVolume())  # debug
                return interface.GetMasterVolume()

    def set_volume(self, decibels):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                # only set volume in the range 0.0 to 1.0
                self.volume = min(1.0, max(0.0, decibels))
                interface.SetMasterVolume(self.volume, None)

    def decrease_volume(self, decibels):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                # 0.0 is the min value, reduce by decibels
                self.volume = max(0.0, self.volume-decibels)
                interface.SetMasterVolume(self.volume, None)

    def increase_volume(self, decibels):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                # 1.0 is the max value, raise by decibels
                self.volume = min(1.0, self.volume+decibels)
                interface.SetMasterVolume(self.volume, None)
                
def AudioDeviceCmdlets(command, output=True):
    process = subprocess.Popen(["powershell", "-Command", "Import-Module .\AudioDeviceCmdlets.dll;", command],stdout=subprocess.PIPE, shell=True)
    proc_stdout = process.communicate()[0]
    if output:
        proc_stdout = proc_stdout[proc_stdout.decode("utf-8", "ignore").index("["):-1]
        return json.loads(proc_stdout) 

def getActiveExecutablePath():
    hWnd = ctypes.windll.user32.GetForegroundWindow()
    if hWnd == 0:
        return None # Note that this function doesn't use GetLastError().
    else:
        _, pid = win32process.GetWindowThreadProcessId(hWnd)
        return psutil.Process(pid).exe()
    
    

# Setup TouchPortal connection
TPClient = TouchPortalAPI.Client('Windows-Tools')

running = False
updateXY = True

def muteAndUnMute(process, value):
    if value == "Mute":
        value = 1
    elif value == "Unmute":
        value = 0
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session.SimpleAudioVolume
        if session.Process and session.Process.name() == process:
            volume.SetMute(value, None)

def volumeChanger(process, action, value):
    if action == "Set":
        AudioController(str(process)).set_volume((int(value)*0.01))
    elif action == "Increase":
        AudioController(str(process)).increase_volume((int(value)*0.01))
    elif action == "Decrease":
        AudioController(str(process)).decrease_volume((int(value)*0.01))

def setMasterVolume(Vol):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    scalarVolume = int(Vol) / 100
    volume.SetMasterVolumeLevelScalar(scalarVolume, None)

def getMasterVolume() -> int:
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    return int(round(volume.GetMasterVolumeLevelScalar() * 100))

def AdvancedMouseFunction(x, y, delay, look):
    if look == 0:
        look = None
    if look == "None":
        try:
            pyautogui.moveTo(x, y, delay)
        except:
            pass
    elif look == "Start slow, end fast":
        try:
            pyautogui.moveTo(x, y, delay, pyautogui.easeInQuad)
        except:
            pass
    elif look == "Start fast, end slow":
        try:
            pyautogui.moveTo(x, y, delay, pyautogui.easeOutQuad)
        except:
            pass
    elif look == "Start and end fast, slow in middle":
        try:
            pyautogui.moveTo(x, y, delay, pyautogui.easeInOutQuad)
        except:
            pass
    elif look == "bounce at the end":
        try:
            pyautogui.moveTo(x, y, delay, pyautogui.easeInBounce)
        except:
            pass
    elif look == "rubber band at the end":
        try:
            pyautogui.moveTo(x, y, delay, pyautogui.easeInElastic)
        except:
            pass

old_volume_list = []
monitor_count_old = 0
global_states = []
global Timer
counter = 0
def updateStates():
    global old_volume_list, global_states, Timer, counter, monitor_count_old
    Timer = threading.Timer(0.4, updateStates)
    Timer.start()
    if running:
        current_audio_source = ["Master Volume", "Current app"]
        can_audio_run = True
        try:
            pythoncom.CoInitialize()
            sessions = AudioUtilities.GetAllSessions()
        except Exception as e:
            can_audio_run = False
            pass
        if can_audio_run:
            for x in sessions:
                try:
                    current_audio_source.append(x.Process.name())
                except AttributeError:
                    pass
            if old_volume_list != current_audio_source:
            #print(current_audio_source)
                TPClient.choiceUpdate('KillerBOSS.TP.Plugins.VolumeMixer.Increase/DecreaseVolume.process', current_audio_source[2:])
                TPClient.choiceUpdate('KillerBOSS.TP.Plugins.VolumeMixer.Mute/Unmute.process', current_audio_source[2:])
                TPClient.choiceUpdate("KillerBOSS.TP.Plugins.VolumeMixer.slidercontrol", current_audio_source)
                old_volume_list = current_audio_source

                for eachprocess in current_audio_source[2:]:
                    if eachprocess not in global_states:
                        TPClient.createState(f'KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{eachprocess}', f'{eachprocess} Volume', "0")
                        global_states.append(eachprocess)
                        print(f'creating states for {eachprocess}')

                for x in global_states:
                    if x not in current_audio_source:
                        try:
                            TPClient.removeState(f'KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{x}')
                            print(f'Removing {x}')
                        except Exception:
                            print("exception at 765")
                            pass
                    
            for x in global_states:
                try:
                    appVolume = str(int(AudioController(x).process_volume()*100))
                    TPClient.stateUpdate(f'KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{x}', appVolume)
                    TPClient.send(
                    {
                        "type":"connectorUpdate",
                        "connectorId":f"pc_Windows-Tools_KillerBOSS.TP.Plugins.VolumeMixer.connectors.APPcontrol|KillerBOSS.TP.Plugins.VolumeMixer.slidercontrol={x}",
                        "value": appVolume
                    }
                )
                except TypeError:
                    pass 
        counter = counter + 1
        if counter >= 5:

            
            
            TPClient.stateUpdate("KillerBOSS.TP.Plugins.Application.currentFocusedAPP", pygetwindow.getActiveWindowTitle())
            activeWindow = getActiveExecutablePath()
            if activeWindow != None:
                activeWindow = os.path.basename(getActiveExecutablePath())
            TPClient.send(
                    {
                        "type":"connectorUpdate",
                        "connectorId":"pc_Windows-Tools_KillerBOSS.TP.Plugins.VolumeMixer.connectors.APPcontrol|KillerBOSS.TP.Plugins.VolumeMixer.slidercontrol=Master Volume",
                        "value": str(getMasterVolume())
                    }
                )
            try:
                TPClient.send(
                    {
                        "type":"connectorUpdate",
                        "connectorId":"pc_Windows-Tools_KillerBOSS.TP.Plugins.VolumeMixer.connectors.APPcontrol|KillerBOSS.TP.Plugins.VolumeMixer.slidercontrol=Current app",
                        "value": "0" if activeWindow == None else str(int(AudioController(activeWindow).process_volume()*100))
                    }
                )
            except TypeError:
                TPClient.send(
                    {
                        "type":"connectorUpdate",
                        "connectorId":"pc_Windows-Tools_KillerBOSS.TP.Plugins.VolumeMixer.connectors.APPcontrol|KillerBOSS.TP.Plugins.VolumeMixer.slidercontrol=Current app",
                        "value": "0"
                    }
                )
        if counter >= 34:
            counter = 0
            try:
                output = AudioDeviceCmdlets('Get-AudioDevice -List | ConvertTo-Json')
                for x in output:
                    if x['Type'] == "Playback" and x['Default'] == True:
                        TPClient.stateUpdate('KillerBOSS.TP.Plugins.Sound.CurrentOutputDevice', x['Name'])
                    elif x['Type'] == "Recording" and x['Default'] == True:
                        TPClient.stateUpdate('KillerBOSS.TP.Plugins.Sound.CurrentInputDevice', x['Name'])
            except json.JSONDecodeError as err:
                print("Audio Device Decode to Json Error: ", err)
                pass     
                   
        
        # try:
        #     currentActiveWindowIco = extract_icon.extractIco(getActiveExecutablePath())
        #     TPClient.stateUpdate("KillerBOSS.TP.Plugins.Application.CurrentProgramIco", currentActiveWindowIco)
        # except:
        #     pass
        # print(currentActiveWindowIco)

        # Advanced Mouse
        if updateXY:
            TPClient.stateUpdate('KillerBOSS.TP.Plugins.AdvanceMouse.MousePos.X', str(pyautogui.position()[0]))
            TPClient.stateUpdate('KillerBOSS.TP.Plugins.AdvanceMouse.MousePos.Y', str(pyautogui.position()[1]))
    else:
        Timer.cancel()
        print('canceled the Timer')
        
running = False
updateStates()
@TPClient.on('info')
def onStart(data):
    if settings := data.get('settings'):
        handleSettings(settings, False)

    th_uptime = threading.Thread(target=pc_uptime)
    th_uptime.start()
    global running
    running = True
    updateStates()


settings = {}
def handleSettings(settings, on_connect=False):
    newsettings = { list(settings[i])[0] : list(settings[i].values())[0] for i in range(len(settings)) }
    global stop_run_continuously
    
    
    ###Stopping Scheduled Tasks and Clearing List
    stop_run_continuously.set()
    schedule.clear()
    time.sleep(2)
    
    ### looping thru and checking if zero, if so we dont reschedule it
    for setting in newsettings:
        interval = float(newsettings[setting])

        if setting == "Update Interval: Hard Drive":
            if interval > 0:
                schedule.every(interval).seconds.do(disk_usage)
                print(f"{setting} is now {interval}")
            else:
                print(f"{setting} is TURNED OFF")

        if settings == "Update Interval: Network'":
            if interval > 0:
                #schedule.every(interval).seconds.do(vd_check)
                print(f"{setting} is now {interval}")
            else:
                print(f"{setting} is TURNED OFF")


        if setting == "Update Interval: Active Monitors":
            if interval > 0:
                schedule.every(interval).seconds.do(check_number_of_monitors)
                print(f"{setting} is now {interval}")
            else:
                print(f"{setting} is TURNED OFF")

        if setting == "Update Interval: Active Windows":
            if interval > 0:
                schedule.every(interval).seconds.do(get_windows_update)
                print(f"{setting} is now {interval}")
            else:
                print(f"{setting} is TURNED OFF")

# Start the background threads again
    stop_run_continuously = run_continuously()
    return settings


# Settings handler
@TPClient.on(TouchPortalAPI.TYPES.onSettingUpdate)
def onSettingUpdate(data):
    if (settings := data.get('values')):
        handleSettings(settings, False)


    
@TPClient.on(TouchPortalAPI.TYPES.onAction)
def Actions(data):
    print(data)
    if data['actionId'] == 'KillerBOSS.TP.Plugins.VolumeMixer.Mute/Unmute':
        if data['data'][0]['value'] is not '':
            muteAndUnMute(data['data'][0]['value'], data['data'][1]['value'])
    if data['actionId'] == 'KillerBOSS.TP.Plugins.VolumeMixer.Increase/DecreaseVolume':
        volumeChanger(data['data'][0]['value'], data['data'][1]['value'], data['data'][2]['value'])
    
    if data['actionId'] == 'KillerBOSS.TP.Plugins.VolumeMixer.SetmasterVolume':
        setMasterVolume(data['data'][0]['value'])

    if data['actionId'] == 'KillerBOSS.TP.Plugins.AdvanceMouse.HoldDownToggle':
        if data['data'][0]['value'] == 'Down':
            pyautogui.mouseDown(button=(data['data'][1]['value']).lower())
        elif data['data'][0]['value'] == "Up":
            pyautogui.mouseUp(button=(data['data'][1]['value']).lower())
    if data['actionId'] == "KillerBOSS.TP.Plugins.AdvanceMouse.teleport":
        AdvancedMouseFunction(int(data['data'][0]['value']),int(data['data'][1]['value']), int(data['data'][3]['value']), data['data'][2]['value'])
    if data['actionId'] == "KillerBOSS.TP.Plugins.AdvanceMouse.MouseClick":
        pyautogui.click(clicks=int(data['data'][0]['value']), button=(data['data'][2]['value']).lower(), interval=float(data['data'][1]['value']))
    if data['actionId'] == 'KillerBOSS.TP.Plugins.AdvanceMouse.Function':
        try:
            pyautogui.move(int(data['data'][0]['value']), int(data['data'][1]['value']))
        except Exception:
            pass
    if data['actionId'] == 'KillerBOSS.TP.Plugins.ChangeAudioOutput' and data['data'][0]['value'] != "Pick One": 
        print('powershell is running')
        AudioDeviceCmdlets(f"(Get-AudioDevice -list | Where-Object Name -like (\'{data['data'][1]['value']}') | Set-AudioDevice).Name", output=False)
    if data['actionId'] == "KillerBOSS.TP.Plugins.AdvanceMouse.StopMouseUpdate":
        global updateXY
        if data['data'][0]['value'] == "OFF":
            updateXY = False
        else:
            updateXY = True
            
        
    if data['actionId'] == "KillerBOSS.TP.Plugins.window.current":
        if data['data'][0]['value'] == "Clipboard":
            print("clip board stuff")
            current_window = pygetwindow.getActiveWindowTitle()
            print(current_window)
            screenshot_window(capture_type=3, window_title=current_window, clipboard=True)
        elif data['data'][0]['value'] == "File":
            current_window = pygetwindow.getActiveWindowTitle()
            afile_name = (data['data'][1]['value']) +"/" +(data['data'][2]['value']) 
            screenshot_window(capture_type=3, window_title=current_window, clipboard=False, save_location=afile_name)
            
               ###using wildcard to FILE
    if data['actionId'] == "KillerBOSS.TP.Plugins.screencapture.window.file.wildcard":
        global windows_active
        windows_active = get_windows()
        for thing in windows_active:
            if data['data'][0]['value'].lower() in thing.lower():
                print("We found", thing)
                if data['data'][4]['value'] == "Clipboard":
                    print("cliboard mmk")
                    screenshot_window(capture_type=int(data['data'][1]['value']), window_title=thing, clipboard=True)
                    
                elif data['data'][4]['value'] == "File":
                    print("File stuf")
                    afile_name = (data['data'][2]['value']) +"/" +(data['data'][3]['value']) 
  
                    screenshot_window(capture_type=int(data['data'][1]['value']), window_title=thing, clipboard=False, save_location=afile_name)
                break
             
    if data['actionId'] == "KillerBOSS.TP.Plugins.screencapture.full.file":   
        if data['data'][1]['value'] == "Clipboard":
            screenshot_monitor(monitor_number=int(data['data'][0]['value']), clipboard=True)
        elif data['data'][1]['value'] == "File":
            try:
                afile_name = (data['data'][2]['value']) +"/" +(data['data'][3]['value'])    
                screenshot_monitor(monitor_number=int(data['data'][0]['value']), filename=afile_name, clipboard=False)
            except:
                pass
        
    if data['actionId'] == "KillerBOSS.TP.Plugins.screencapture.window.file":
        if (data['data'][0]['value']):
            if data['data'][4]['value'] == "Clipboard":
                screenshot_window(capture_type=int(data['data'][1]['value']), window_title=data['data'][0]['value'], clipboard=True)
            if data['data'][4]['value'] == "File":
                afile_name = (data['data'][2]['value']) +"/" +(data['data'][3]['value'])    
                screenshot_window(capture_type=int(data['data'][1]['value']), window_title=data['data'][0]['value'], clipboard=False, save_location=afile_name)
        
        

    if data['actionId'] == "KillerBOSS.TP.Plugins.screencapture.processcheck":
        app_to_check = data['data'][0]['value'].lower()
        focus = data['data'][1]['value']
        afocus_type = data['data'][2]['value']
        shortcut_to_open = ""
        if data['data'][3]['value']:
            shortcut_to_open = data['data'][3]['value']
        if focus == "Focus":
            focus_check = True
        else:
            focus_check = False
        
        check_process(app_to_check, shortcut_to_open, focus=focus_check, focus_type=afocus_type)
        
        
    if data['actionId'] == "KillerBOSS.TP.Plugins.capture.clipboard":
        send_to_clipboard("text", data['data'][0]['value'] )
        
        
    if data['actionId'] == "KillerBOSS.TP.Plugins.virtualdesktop.actions":
        choice = data['data'][0]['value']
        virtual_desktop(target_desktop=choice)
        
        
    if data['actionId'] == "KillerBOSS.TP.Plugins.virtualdesktop.actions.move_window":
        choice = data['data'][0]['value']
        if data['data'][1]['value'] == "False":
            virtual_desktop(move=True, target_desktop=choice)
        if data['data'][1]['value'] == "True":
            virtual_desktop(move=True, target_desktop=choice, pinned=True)
        
    if data['actionId'] == "KillerBOSS.TP.Plugins.magnifier.actions":
        print(data['data'][0]['value'])
        magnifier(data['data'][0]['value'])
        
    if data['actionId'] == "KillerBOSS.TP.Plugins.toast.create":
        win_toast(atitle=data['data'][0]['value'], amsg=data['data'][1]['value'], aduration=data['data'][2]['value'], icon=data['data'][5]['value'])
        
    if data['actionId'] == "KillerBOSS.TP.Plugins.winextra.emojipanel":
        winextra("Emoji")
        
    if data['actionId'] == "KillerBOSS.TP.Plugins.winextra.keyboard":
        winextra("Keyboard")
            
        
@TPClient.on(TouchPortalAPI.TYPES.onHold_down)
def heldingButton(data):
    print(data)
    while True:
        if TPClient.isActionBeingHeld('KillerBOSS.TP.Plugins.AdvanceMouse.MouseClick'):
            pyautogui.click(clicks=int(data['data'][0]['value']), button=(data['data'][2]['value']).lower(), interval=float(data['data'][1]['value']))
        elif TPClient.isActionBeingHeld('KillerBOSS.TP.Plugins.AdvanceMouse.Function'):
            try:
                pyautogui.move(int(data['data'][0]['value']), int(data['data'][1]['value']))
            except Exception:
                pass
        elif TPClient.isActionBeingHeld('KillerBOSS.TP.Plugins.VolumeMixer.Increase/DecreaseVolume'):
            if data['data'][1]['value'] == "Decrease":
                volumeChanger(data['data'][0]['value'], 'Decrease', data['data'][2]['value'])
                sleep(0.05)
            elif data['data'][1]['value'] == "Increase":
                volumeChanger(data['data'][0]['value'], 'Increase', data['data'][2]['value'])
                sleep(0.05)
        else:
            break

def updateDeviceOutput(options):
    output = AudioDeviceCmdlets('Get-AudioDevice -List | ConvertTo-Json')
    outPutDevice = []
    inputDevice = []
    for x in output:
        if x['Type'] == "Playback":
            outPutDevice.append(x['Name'])
        elif x['Type'] == "Recording":
            inputDevice.append(x['Name'])
    if options == "Output":
        TPClient.choiceUpdate('KillerBOSS.TP.Plugins.ChangeAudioOutput.Device', outPutDevice)
        print('updating Output', outPutDevice)
    elif options == "Input":
        TPClient.choiceUpdate('KillerBOSS.TP.Plugins.ChangeAudioOutput.Device', inputDevice)
        print('updating input', outPutDevice)

@TPClient.on(TYPES.onListChange)
def listChangeAction(data):
    print(data)
    if data['actionId'] == 'KillerBOSS.TP.Plugins.ChangeAudioOutput':
        try:
            updateDeviceOutput(data['value'])
        except KeyError:
            pass
        
    if data['actionId'] == 'KillerBOSS.TP.Plugins.screencapture.window.clipboard':

        pass

@TPClient.on(TYPES.onConnectorChange)
def connectors(data):
    print(data)
    import os
    if data['connectorId'] == "KillerBOSS.TP.Plugins.VolumeMixer.connectors.APPcontrol":
        if data['data'][0]['value'] == "Master Volume" :
            setMasterVolume(data['value'])
        elif data['data'][0]['value'] == "Current app":
            activeWindow = getActiveExecutablePath()
            if activeWindow != "":
                volumeChanger(os.path.basename(activeWindow), "Set", data['value'])
        else:
            try:
                volumeChanger(data['data'][0]['value'], "Set", data['value'])
            except:
                pass

@TPClient.on('closePlugin')
def Shutdown(data):
    global running, Timer
    running = False
    stop_run_continuously.set()
    schedule.clear()
    TPClient.disconnect()
    Timer.cancel()
    sys.exit()
TPClient.connect()

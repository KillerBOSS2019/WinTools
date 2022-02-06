import ctypes
import json
import os
import subprocess
import time
from ctypes import POINTER, cast
from datetime import datetime
from io import BytesIO

import dateutil.relativedelta
import psutil
import pyautogui
import pyttsx3
import win32clipboard
import win32con
import win32gui
import win32process
import win32ui
import win32api
from comtypes import CLSCTX_ALL
from PIL import Image
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from winotify import Notification, audio
import re

#####################################################
#                                                   #
#             Audio Stuff...                        #          
#                                                   #
######################################################

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


#####################################################
#                                                   #
#             Mouse func                            #          
#                                                   #
######################################################

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


#####################################################
#                                                   #
#             Windows Notification                  #          
#                                                   #
######################################################

def win_toast(atitle="", amsg="", buttonText = "", buttonlink = "", sound = "", aduration="short", icon=""):
    ### setting the base notification stuff
    if not os.path.exists(rf"{icon}") or icon == "": icon = os.path.join(os.getcwd(),"src\icon.png")
    print(icon)
    toast = Notification(app_id="WinTools",
                         title=atitle,
                         msg=amsg,
                         icon=icon,
                         duration = aduration.lower(),
                         )
    
    if buttonText != "" and buttonlink != "":
        toast.add_actions(label=buttonText, 
                        link=buttonlink)

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
    toast.set_audio(audioDic[sound], loop=False)
            
    toast.build()
    toast.show()

#####################################################
#                                                   #
#             Clipboard stuff                       #          
#                                                   #
######################################################

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
    except Exception as e:
        print("error screenshot" + e )

#####################################################
#                                                   #
#             Other                                 #          
#                                                   #
######################################################



import sounddevice as sd
import audio2numpy as a2n
### Getting all audio outputs available for use with TTS
def getAllOutput_TTS2():
    audio_dict = {}
    for outputDevice in sd.query_hostapis(0)['devices']:
        if sd.query_devices(outputDevice)['max_output_channels'] > 0:
            audio_dict[sd.query_devices(device=outputDevice)['name']] = sd.query_devices(device=outputDevice)['default_samplerate']
    return audio_dict



def rotate_display(display_num, rotate_choice):
    ### display num is gonna need add or subtractin to match other things..  cause 0 is monitor 1 here, and in other spots 0 is ALL monitors..
    ## so display 0 would actually be display 1 in "settings"
    display_num = display_num -1
    rotation_val=""
    if (rotate_choice != None):
        if (rotate_choice == "180"):
            rotation_val=win32con.DMDO_180
        elif(rotate_choice == "90"):
            rotation_val=win32con.DMDO_270
        elif (rotate_choice == "270"):   
            rotation_val=win32con.DMDO_90
        else:
            rotation_val=win32con.DMDO_DEFAULT

    device = win32api.EnumDisplayDevices(None,display_num)
    dm = win32api.EnumDisplaySettings(device.DeviceName,win32con.ENUM_CURRENT_SETTINGS)
    if((dm.DisplayOrientation + rotation_val)%2==1):
        dm.PelsWidth, dm.PelsHeight = dm.PelsHeight, dm.PelsWidth   
    dm.DisplayOrientation = rotation_val

    win32api.ChangeDisplaySettingsEx(device.DeviceName,dm)
    
    


def magnifier(action):
    if action == "Zoom In":
        pyautogui.hotkey('win', '=')
    if action == "Zoom Out":
        pyautogui.hotkey('win', '-')
    if action == "Exit":
        pyautogui.hotkey('win', 'escape')
    
    
    ### used mostly for checking numlock status
def get_key_state(key):
    import ctypes
    hllDll = ctypes.WinDLL ("User32.dll")
    if key == "NUM LOCK":
        return hllDll.GetKeyState(0x90)
    if key == "CAPS LOCK":
        return hllDll.GetKeyState(0x14)

def move_win_button(direction):
    check = get_key_state('NUM LOCK')
    if not check:
        pyautogui.hotkey('win', 'shift', direction)
    elif check:
        pyautogui.press('numlock')
        pyautogui.hotkey('win', 'shift', direction)
        pyautogui.press('numlock')
    
    
    
def win_shutdown(time, cancel=False):
    ## should we create a shutdown timer/countdown ??  we can warn the user 5 minutes before shutting down so they can cancel
    time= pyautogui.prompt(text='💻 How many MINUTES do you want to wait before shutting down?', title='Shutdown PC?', default='')
    if time:
        try:
            time = int(time) * 60    ## multiplying by 60 to get minutes
            os.system(f"shutdown -s -t {time}")
        except:
            pass
    else:
        os.system(f"shutdown -a")
    




from subprocess import PIPE, run

def out(command):
    result = run(command, stdout=PIPE, stderr=PIPE, universal_newlines=True, shell=True)
    return result.stdout

def get_powerplans(currentcheck=False):
    import re
    pplans={}
    for powerplan in out("powercfg -List").split("\n")[2:]:
        if "Scheme" in powerplan.split():
            ParsedData = powerplan.split(":")[1].split()
            the_data = (ParsedData[0])
            plan_name = (" ".join(ParsedData[1:]))
        
    
            if "*" in plan_name:
                print("this is current")
                ##Update TP State to Show Current Power Plan.
                
                ### Removing all special chars except spaces and *
                regex = re.compile('[^a-zA-Z]')
                #First parameter is the replacement, second parameter is your input string
                plan_name = regex.sub('', plan_name)
                pplans[plan_name]=the_data
                if currentcheck==True:
                    return plan_name
            else:
                regex = re.compile('[^a-zA-Z]')
                plan_name = regex.sub('', plan_name)
                pplans[plan_name]=the_data
        
    return pplans
    

################### THIS NEEDS IMPLEMENTED ###################
### Should we let the user 'schedule' a ping that goes every X seconds to give results?
### Should we just have it on press only so they can make a custom repeat rate, or spam issues with that maybe?
def ping_ip(the_ip):

    ping_result = subprocess.check_output('ping -n 3 ' + the_ip,  universal_newlines=True)

    ping_pattern = re.compile("time[<, =, >](?P<ms>\d+)ms")
    ttl_pattern = re.compile("TTL[<, =, >](?P<ms>\d+)")
    ip_pattern = re.compile("(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})")


    pings = re.findall(ping_pattern, ping_result)
    ttl =  re.findall(ttl_pattern, ping_result)
    pinged_ip = ip_pattern.search(ping_result)[0]

    for i in range(0, len(pings)):
        pings[i] = int(pings[i])

    for i in range(0, len(ttl)):
        ttl[i] = int(ttl[i])

    ip_dict = {
        "The IP": pinged_ip,
        "Average": f'{round(sum(pings) / len(pings))}ms',
        "TTL": f'{round(sum(ttl) / len(ttl))}'
    }

    print(f"The average ping for {the_ip} was {round(sum(pings) / len(pings), 2)}ms")
    print(f"The average TTL for {the_ip} was {round(sum(ttl) / len(ttl))}")
    
    return ip_dict




def change_pplan(choice):
    the_thing=(get_powerplans())
    out(f"powercfg.exe /S {the_thing[choice]}")
    
    
        
def winextra(action):
    
  # if action == "Keep Active, Minimize All (Toggle)":
  #     pyautogui.hotkey('win', 'home')
  #    
  # if action == "Minimize All (Toggle)":
  #     pyautogui.hotkey('win', 'd')
#         
  #  if action =="Clipboard History":
  #      pyautogui.hotkey('win', 'v')
        
    if action == "Emoji":
        pyautogui.hotkey('win', '.')
        
    if action == "Keyboard":
        pyautogui.hotkey('win', 'ctrl', "o")
        
    

def getActiveExecutablePath():
    hWnd = ctypes.windll.user32.GetForegroundWindow()
    if hWnd == 0:
        return None # Note that this function doesn't use GetLastError().
    else:
        _, pid = win32process.GetWindowThreadProcessId(hWnd)
        return psutil.Process(pid).exe()

def get_app_icon():
    
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

from pyvda import AppView, VirtualDesktop, get_virtual_desktops


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
            
    elif target_desktop != "Previous" or "Next":  ## else if 0, 1, 2 3, etc.. anything but next or previous
        print("THE TARGETED DESKTOP IS", target_desktop)
        if move:
            target_desktop2 = int(target_desktop)
            print("The Target desktop secondary thing is", target_desktop2)
            current_window = AppView.current()
            target_desktop = VirtualDesktop(int(target_desktop2))
            current_window.move(target_desktop)

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
    net_io = psutil.net_io_counters(nowrap=False)
    adict['sent'] = get_size(net_io.bytes_sent)
    adict['received'] = get_size(net_io.bytes_recv)
    return adict

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
    hours = rd.hours
    minutes = rd.minutes
    seconds = rd.seconds
    seconds = int(seconds)
    minutes = int(minutes)
    hours = int(hours)
    if minutes <10:
        minutes = "0" + str(minutes)
    if seconds <10:
        seconds = "0" + str(seconds)
    time.sleep(1)
    
    return f"{rd.days}:{hours}:{minutes}:{seconds}"

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
        except Exception as e:
            print("error loading shortcut " + e)

def get_windows():
    results = []
    def winEnumHandler(hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetWindowText(hwnd):
    
                results.append(win32gui.GetWindowText(hwnd))
                
    win32gui.EnumWindows(winEnumHandler, None)
    return results

def AudioDeviceCmdlets(command, output=True):
    process = subprocess.Popen(["powershell", "-Command", "Import-Module .\AudioDeviceCmdlets.dll;", command],stdout=subprocess.PIPE, shell=True)
    proc_stdout = process.communicate()[0]
    if output:
        proc_stdout = proc_stdout[proc_stdout.decode("utf-8", "ignore").index("["):-1]
        return json.loads(proc_stdout) 



import sounddevice as sd
import audio2numpy as a2n
import pyttsx3
def TextToSpeech(message, voicesChoics, volume=100, rate=100, output="Default"):
    engine = pyttsx3.init()
    voices = engine.getProperty('voices')
    engine.setProperty("volume", volume/100)
    engine.setProperty("rate", rate)
    engine.setProperty('voice', voices[1].id if voicesChoics == "Female" else voices[0].id)

    if output =="Default":
        try:
            engine.say(message)
            engine.runAndWait()
            engine.stop()
        except Exception as e:
            print("TTS Error", e)
    else:
        try:
            appdata = os.getenv('APPDATA')
            engine.save_to_file(message, rf"{appdata}/TouchPortal/Plugins/WinTools/speech.wav")
            engine.runAndWait()
            engine.stop()
            device = getAllOutput_TTS2()  ### can make this list returned a global + save that will pull it once and thats it?  instead of every time this action is called..which could be troublesome..
            sd.default.samplerate = device[output]
            sd.default.device = output +", MME"
            x,sr=a2n.audio_from_file(rf"{appdata}/TouchPortal/Plugins/WinTools/speech.wav")
            sd.play(x, sr, blocking=True)
        except Exception as e:
            print("test", e)
            
        
        
def activate_windows_setting(choice=False):
    
    settings ={
        "SYSTEM:  Display": 'ms-settings:display',
        "SYSTEM:  Advanced Display": 'ms-settings:display-advanced',
        "SYSTEM:  Night Light": 'ms-settings:nightlight',
        "SYSTEM:  Sound": 'ms-settings:sound',
        "SYSTEM:  Manage Sound Devices": 'ms-settings:sound-devices',
        "SYSTEM:  Manage App/Device Volume": 'ms-settings:apps-volume',
        "SYSTEM:  App Volume & Device Preferences": 'ms-settings:apps-volume',
        "SYSTEM:  Notifcations & Actions": 'ms-settings:notifications',
        "SYSTEM:  Power & Sleep": 'ms-settings:powersleep',
        "SYSTEM:  Battery": 'ms-settings:batterysaver',
        "SYSTEM:  Battery Usage Details": 'ms-settings:batterysaver-usagedetails',
        "SYSTEM:  Default Save Locations": 'ms-settings:savelocations',
        "SYSTEM:  Multi-Tasking": 'ms-settings:multitasking',
        "SYSTEM:  Sign-in Options": 'ms-settings:signinoptions',     ## this was 'ACCOUNTS'
        "SYSTEM:  Date & Time": 'ms-settings:dateandtime',             ## this was 'TIME & LANGUAGE*
        "SYSTEM:  Time Region": 'ms-settings:regionformatting',        ## this was 'TIME & LANGUAGE*
        "SYSTEM:  Settings Home Page": 'ms-settings:',
        "NETWORK:  Ethernet": 'ms-settings:network-ethernet',
        "NETWORK:  Wi-Fi": 'ms-settings:network-wifi',
        "PERSONALIZATION:  Background": 'ms-settings:personalization-background',
        "PERSONALIZATION:  Colors": 'ms-settings:personalization-colors',
        "PERSONALIZATION:  Lock Screen": 'ms-settings:lockscreen',
        "PERSONALIZATION:  Themes": 'ms-settings:themes',
        "PERSONALIZATION:  Start Folders": 'ms-settings:personalization-start-places',
        "APPS:  Apps & Features": 'ms-settings:appsfeatures',
        "APPS:  Manage Startup Apps": 'ms-settings:startupapps',
        "APPS:  Manage Default Apps": 'ms-settings:defaultapps',
        "APPS:  Manage Optional Features": 'ms-settings:optionalfeatures',
        "GAMING:  Game Bar": 'ms-settings:gaming-gamebar',
        "GAMING:  Game DVR": 'ms-settings:gaming-gamedvr',
        "GAMING:  Game Mode": 'ms-settings:gaming-gamemode',
        "GAMING:  XBOX Networking": 'ms-settings:gaming-xboxnetworking',
        "PRIVACY:  Activity History": 'ms-settings:privacy-activityhistory',
        "PRIVACY:  Webcam": 'ms-settings:privacy-webcam',
        "PRIVACY:  Microphone": 'ms-settings:privacy-microphone',
        "PRIVACY:  Background Apps": 'ms-settings:privacy-backgroundapps',
        r"UPDATE & SECURITY:  Windows Update": 'ms-settings:windowsupdate',
        r"UPDATE & SECURITY:  Windows Recovery": 'ms-settings:recovery',
        r"UPDATE & SECURITY:  Update history": 'ms-settings:windowsupdate-history',
        r"UPDATE & SECURITY:  Restart Options": 'ms-settings:windowsupdate-restartoptions',
        r"UPDATE & SECURITY:  Delivery Optimization": 'ms-settings:delivery-optimization',
        r"UPDATE & SECURITY:  Windows Security": 'ms-settings:windowsdefender',
        r"UPDATE & SECURITY:  Windows Defender": 'windowsdefender:',
        r"UPDATE & SECURITY:  For Developers": 'ms-settings:developers',
    }
    if not choice:
        settings_list =[]
        for thing in settings:
            settings_list.append(thing)
        return settings_list
    else:
        os.system(f'explorer "{settings[choice]}"')
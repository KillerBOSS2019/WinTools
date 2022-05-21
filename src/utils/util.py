import ctypes
import json
import os
import sys
#from threading import Thread
import threading
import time
import requests 
from ctypes import POINTER, cast, windll
from datetime import datetime
from io import BytesIO
from subprocess import PIPE, run
import subprocess
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
from pycaw.magic import MagicManager, MagicSession   ### These are used in main.py
from pycaw.constants import AudioSessionState   ### These are used in main.py
from winotify import Notification, audio, Registry, PY_EXE, Notifier
from pyvda import AppView, VirtualDesktop, get_virtual_desktops
from ctypes import windll, create_unicode_buffer, c_wchar_p, sizeof  # drive details
from string import ascii_uppercase   # used for getting drive details
import pyttsx3
import re
import pywintypes
import pythoncom
import winreg
import base64
import sounddevice as sd
import audio2numpy as a2n
import numpy as np
from PIL import Image
from win32com.client import GetObject
import mss
import TouchPortalAPI

TPClient = TouchPortalAPI.Client('Windows-Tools')

"""Current Working Directory"""
cwd = os.path.join(os.getcwd())

#####################################################
#                                                   #
#             Audio Stuff...                        #          
#                                                   #
######################################################

class AudioController(object):
    def __init__(self, process_name):
        self.process_name = process_name
        self.volume = self.process_volume()

    def process_volume(self):
        pythoncom.CoInitialize() ### 3rd coinitilize...
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                #print('Volume:', interface.GetMasterVolume())  # debug
                return interface.GetMasterVolume()

    def set_volume(self, decibels):
        sessions = AudioUtilities.GetAllSessions()
        print(self.process_name)
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


######################################################
##                                                   #
##             Windows Notification                  #          
##                                                   #
#######################################################
app_id = "WinTools-Notify"
app_path = os.path.abspath(__file__)
r = Registry(app_id, PY_EXE, app_path, force_override=True)
notifier = Notifier(r)

@notifier.register_callback
def clear_notify():
    notifier.clear()


def win_toast(atitle="", amsg="", buttonText = "", buttonlink = "", sound = "", aduration="short", icon=""):
    ### setting the base notification stuff
    if not os.path.exists(rf"{icon}") or icon == "": icon = os.path.join(os.getcwd(),"src\icon.png")

    toast = notifier.create_notification(
                         title=atitle,
                         msg=amsg,
                         icon=icon,
                         duration = aduration.lower(),
                         )
    
    ### we can allow multiple links
    if buttonText != "" and buttonlink != "":
        toast.add_actions(label=buttonText, 
                        launch=buttonlink)
      # toast.add_actions("Open Github", "https://github.com/versa-syahptr/winotify")
      # toast.add_actions("Quit app", "Kewl")
      # toast.add_actions("spam", "sweet")

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

    
    

def get_monitors2():
    objWMI = GetObject('winmgmts:\\\\.\\root\\WMI').InstancesOf('WmiMonitorID')
    count = 0
    monitor_list = []
    for obj in objWMI:
        count = count + 1
       # print("######  Monitor " +str(count) + " ########")
        if obj.Active != None:
            monitor_is_active = str(obj.Active)
        if obj.InstanceName != None:
           pass
        if obj.ManufacturerName != None:
            monitor_manufacturer = (bytes(obj.ManufacturerName)).decode()
            split_manufacturer_name = monitor_manufacturer.split("\x00")
        if obj.UserFriendlyName != None:
            monitor_name = (bytes(obj.UserFriendlyName)).decode()
            split_monitor_name = monitor_name.split("\x00")
            
        the_end = (str(count) +": "+ str((split_monitor_name[0]))+"("+str(split_manufacturer_name[0])+")")
        print(the_end)
        monitor_list.append(the_end)

    return monitor_list
    



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

def get_clipboard_data():
    try:
        win32clipboard.OpenClipboard()
        data = win32clipboard.GetClipboardData()
        win32clipboard.CloseClipboard()
        print("HELLO?")
        return data
    except TypeError as err:
        print("BROKEN")
        return "Invalid Clipboard Data"


## Image to Bytes
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
                print("Saved to Folder")
    except Exception as e:
        print("error screenshot" + e )

#####################################################
#                                                   #
#             Other                                 #          
#                                                   #
######################################################


def getFrame_base64(frame_image):
 #  # Get frame (only rgb - smaller size)
 #  frame_rgb     = mss.mss().grab(mss.mss().monitors[2]).rgb 

 #  # Convert it from bytes to resize
 #  frame_image   = Image.frombytes("RGB", (1920, 1080), frame_rgb, "raw", "RGB") 
    
    ### TEMP SAVING IMAGE TO BUFFER THEN TO BASE 64
    buffer = BytesIO()
    frame_image.save(buffer, format='PNG')
    #### Save it to File #####
    b64_str = base64.standard_b64encode(buffer.getvalue())
    
    frame_image.close()
    return b64_str



"""We have to make this only fire one time when the function is"""
def resize_image(im, height, width):
    if im.size[0] == height and im.size[1] == width: 
        return im
    else:
        print("We are actually resizing")
        size = (height,width)
        im.load()
        bands = list(im.split())
        a = np.asarray(bands[-1])
        a.flags.writeable = True
        a[a != 0] = 1
        bands[-1] = Image.fromarray(a)
        bands = [b.resize(size, Image.LINEAR) for b in bands]
        a = np.asarray(bands[-1])
        a.flags.writeable = True
        a[a != 0] = 255
        bands[-1] = Image.fromarray(a)
        im = Image.merge('RGBA', bands)
        return im
 
def bgra_to_rgba(sct_img):
    img = Image.frombytes('RGB', sct_img.size, sct_img.rgb,'raw')
    img = img.convert("RGBA")
    return img

def capture_around_mouse(height, width, livecap=False, overlay=False):
    m_position = pyautogui.position()
    if livecap:
        if overlay:
            monitor_number=1
            with mss.mss() as sct:
                screenshot_size = [height, width]
                monitor = {
                "top": m_position.y - screenshot_size[0] // 2,
                "left": m_position.x - screenshot_size[1] // 2,
                "width": screenshot_size[0],
                "height": screenshot_size[1],
                "mon": monitor_number,
                }
                sct_img = sct.grab(monitor)
                
              #  que = queue.Queue()
              #  thr = threading.Thread(target = lambda q, arg : q.put(bgra_to_rgba(sct_img)), args = (que, 2))
               # thr = threading.Thread(target = lambda q, arg : q.put(dosomething(arg)), args = (que, 2))
              #  thr.start()
              #  thr.join()
              #  while not que.empty():
              #      img=que.get()
              #      print("Whats this though?")
                
                finalresult = ""
                img = bgra_to_rgba(sct_img)
                try:
                    finalresult = Image.alpha_composite(img, overlay)
                except ValueError as err:
                    global cap_live
                    cap_live = False
                    print("Value Error", err)
                    
                if finalresult:
                    return finalresult
                else:
                    return "NO RESULT"
        
        elif not overlay:
            monitor_number=1
            with mss.mss() as sct:
                screenshot_size = [height, width]
                monitor = {
                "top": m_position.y - screenshot_size[0] // 2,
                "left": m_position.x - screenshot_size[1] // 2,
                "width": screenshot_size[0],
                "height": screenshot_size[1],
                "mon": monitor_number,
                }
                sct_img = sct.grab(monitor)
                img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
                return img
            
    #""" Else if not live cap"""       
    else:
        print("PUSH TO HOLD CAPTURE ACTIVE")
        monitor_number=1
        with mss.mss() as sct:
            screenshot_size = [height, width]
            monitor = {
              "top": m_position.y - screenshot_size[0] // 2,
              "left": m_position.x - screenshot_size[1] // 2,
              "width": screenshot_size[0],
              "height": screenshot_size[1],
              "mon": monitor_number,
            }
            sct_img = sct.grab(monitor)
            img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")  ##   Instead of making a temp file we get it direct from raw to clipboard
            return img

def get_cursor_choices():
    path = cwd+"\mouse_overlays"
    dirs = os.listdir( path )
    cursor_list=[]
    for item in dirs:
      #  if item.endswith(".png"):
        #    cursor_list.append(item.split(".")[0])
        cursor_list.append(item)
    return cursor_list


def getAllOutput_TTS2():
    audio_dict = {}
    for outputDevice in sd.query_hostapis(0)['devices']:
        if sd.query_devices(outputDevice)['max_output_channels'] > 0:
            audio_dict[sd.query_devices(device=outputDevice)['name']] = sd.query_devices(device=outputDevice)['default_samplerate']
    return audio_dict



def rotate_display(display_num, rotate_choice):
    display_num = display_num.split(":")[0]
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
    
    


def magnifier(action, amount=None):
    if action == "Zoom In":
        mag_level(amount)
    if action == "Zoom Out":
        mag_level(amount)
    if action == "Dock":
        mag_mode(1)
    if action == "Full Screen":
        mag_mode(2)
    if action == "Lens":
        mag_mode(3)
    if action == "Invert Colors":
        mag_invert(switch="On")
       # pyautogui.hotkey('ctrl', 'alt', 'i')
    if action == "Exit":
        try:
            out("""wmic process where "name='magnify.exe'" delete""")
        except:
            pass

    
    
    ### used mostly for checking numlock status
def get_key_state(key):
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
    
    ## IF Blank then we show the GUI
    if time == "":
        time_set= pyautogui.prompt(text='💻 How many MINUTES do you want to wait before shutting down?', title='Shutdown PC?', default='')
        if time_set == "0":
            os.system(f"shutdown -a")
            pyautogui.alert("❗ ABORTED SYSTEM SHUTDOWN ❗")
            time="DONE"
        if type(time) == int:
            if int(time_set) > 0:
                print(f"Shutdown in {time_set} minutes")
                time_set = int(time_set) * 60
                os.system(f"shutdown -s -t {time}")
                time="DONE"
                
    if time == "NOW":
        os.system(f"shutdown -s -t 0")
        
    if type(time) == int:
    
        if int(time) >0:
            try:
                print("Shutdown soon?")
                time = int(time) * 60 
                os.system(f"shutdown -s -t {time}")
            except:
                pass
        elif int(time) == None or 0:
            os.system(f"shutdown -a")
            pyautogui.alert("❗ ABORTED SYSTEM SHUTDOWN ❗")
        


def out(command):
    systemencoding = windll.kernel32.GetConsoleOutputCP()
    systemencoding= f"cp{systemencoding}"
    output = subprocess.run(command, stdout=subprocess.PIPE, shell=True)
    result = str(output.stdout.decode(systemencoding))
    return result

def get_powerplans():
    pplans={}
    current = None
    for powerplan in out("powercfg -List").split("\n"):
        if ":" in powerplan:
            ParsedData = powerplan.split(":")[1].split()
            the_data = (ParsedData[0])
            plan_name = (" ".join(ParsedData[1:]))

            if "*" in plan_name:
                plan_name = plan_name[plan_name.find("(") + 1: plan_name.find(")")]
                pplans[plan_name]=the_data
                current = plan_name
            else:
                plan_name = plan_name[plan_name.find("(") + 1: plan_name.find(")")]
                pplans[plan_name]=the_data
    return (pplans, current)


def get_ip_details(choice):
	if choice == "choice1":
		choice = "https://ip.seeip.org/geoip/"
	if choice == "choice2":
		choice = "http://ipinfo.io/json"
	r = requests.get(choice)
	r= json.loads(r.text)
	list = ['ip', 'country',  'region', 'city', 'organization', 'timezone', 'postal']
	ip_dict = {}
	for item in r:
		for thing in list:
			if item in thing:
				ip_dict[thing] = r[item]
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

###   def get_app_icon():
###       
###       active_path = getActiveExecutablePath()
###       print(active_path)
###    
###       if active_path == "C:\Windows\System32\ApplicationFrameHost.exe" or None:
###          pass
###       else:
###          ico_x = win32api.GetSystemMetrics(win32con.SM_CXICON)
###          ico_y = win32api.GetSystemMetrics(win32con.SM_CYICON)
###          try:
###              large, small = win32gui.ExtractIconEx(getActiveExecutablePath(),0)
###              win32gui.DestroyIcon(small[0])
###   
###              hdc = win32ui.CreateDCFromHandle( win32gui.GetDC(0) )
###              hbmp = win32ui.CreateBitmap()
###              hbmp.CreateCompatibleBitmap( hdc, ico_x, ico_x )
###              hdc = hdc.CreateCompatibleDC()
###   
###              hdc.SelectObject( hbmp )
###              hdc.DrawIcon( (0,0), large[0] )
###   
###              hbmp.SaveBitmapFile( hdc, r'C:\Users\dbcoo\AppData\Roaming\TouchPortal\plugins\WinTools\tmp\newwwwicon.bmp')  
###              time.sleep(2)
###          except IndexError as err:
###               print(err)
###               time.sleep(5)
def get_app_icon(active_path):
   # active_path = getActiveExecutablePath()
    if active_path != "C:\Windows\System32\ApplicationFrameHost.exe" or None:
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
            hbmp.SaveBitmapFile( hdc, cwd+"/tempicon.bmp") 

            img = Image.open(cwd+"/tempicon.bmp")
            #img=getFrame_base64(img).decode()
            TPClient.stateUpdate(stateId="KillerBOSS.TP.Plugins.Application.currentFocusedAPP.icon", stateValue=getFrame_base64(img).decode())
            time.sleep(2)
        except IndexError as err:
            print(err)
            time.sleep(5)



def current_vd():
    current_d = VirtualDesktop.current()
    TPClient.stateUpdate(f"KillerBOSS.TP.Plugins.virtualdesktop.current_vd", f"[{current_d.number}] {current_d.name}")
    return current_d

def vd_pinn_app(pinned=True):
    if pinned:
        AppView.current()
        AppView.current().pin()
    if not pinned:
        current_window = AppView.current()
        target_desktop = VirtualDesktop(current_vd().number)
        current_window.move(target_desktop)
        
    

def virtual_desktop(target_desktop=None, move=False, pinned=False):
    number_of_active_desktops = len(get_virtual_desktops())
    """Retrieving VD Number from name"""
    target_desktop = int(target_desktop[target_desktop.find("[") + 1: target_desktop.find("]")])

    if target_desktop == "Next":
        if move:
            current_window = AppView.current()
            target_desktop = VirtualDesktop(target_desktop)
            current_window.move(target_desktop) 
            if pinned:
                AppView.current().pin()
        elif not move:
            if target_desktop <= number_of_active_desktops:
                VirtualDesktop(target_desktop).go()
            else:
                print("too many")
        
    elif target_desktop == "Previous":
        if move:
            AppView.current().move(VirtualDesktop(target_desktop)) 
            if pinned:
                AppView.current().pin()
                
        elif not move:
            if target_desktop > 0:
                if pinned:
                    AppView.current().pin()
                VirtualDesktop(target_desktop).go()
            else:
                pass
            
    elif target_desktop != "Previous" or "Next":  
        if move:
            if pinned:
                AppView.current().move(VirtualDesktop(target_desktop)) 
                vd_pinn_app()
                VirtualDesktop(target_desktop).go()
            else:
                AppView.current().move(VirtualDesktop(target_desktop)) 
        elif not move:
            if target_desktop > 0:
                VirtualDesktop(target_desktop).go()
                if pinned:
                    vd_pinn_app()
    current_vd()
    
def rename_vd(name, number=None):
    number_of_active_desktops = len(get_virtual_desktops())
    new_vd = VirtualDesktop(number_of_active_desktops)
    VirtualDesktop.rename(new_vd, name)
    

def create_vd(name=None):
    if name:
        VirtualDesktop.create()
        rename_vd(name)
    else:
        VirtualDesktop.create()
        


def remove_vd(remove, fallbacknum=None):
    if fallbacknum:
        try:
            fall_back =VirtualDesktop(fallbacknum)
            remove_vd = VirtualDesktop(remove) 
            VirtualDesktop.remove(remove_vd, fall_back)
        except ValueError as err:
            return err
    else:
        try:
            remove_vd=VirtualDesktop(remove)
            VirtualDesktop.remove(remove_vd)
        except ValueError as err:
            return err
    current_vd()



def get_vd_name(number):
    name = VirtualDesktop(number).name
    if not name:
        name = f"Desktop {number}"
    return name

#### END OF VD STUFF ####


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


def get_win_drive_names2():
    volumeNameBuffer = create_unicode_buffer(1024)
    fileSystemNameBuffer = create_unicode_buffer(1024)
    serial_number = None
    max_component_length = None
    file_system_flags = None
    drive_names = {}
    
    #  Get the drive letters, then use the letters to get the drive names
    bitmask = (bin(windll.kernel32.GetLogicalDrives())[2:])[::-1]  # strip off leading 0b and reverse
    drive_letters = [ascii_uppercase[i] + ':/' for i, v in enumerate(bitmask) if v == '1']
    for d in drive_letters:
        rc = windll.kernel32.GetVolumeInformationW(c_wchar_p(d), volumeNameBuffer, sizeof(volumeNameBuffer),
                                                   serial_number, max_component_length, file_system_flags,
                                                   fileSystemNameBuffer, sizeof(fileSystemNameBuffer))
        if rc:
            if volumeNameBuffer.value:
                drive_names[d[:2].upper()] = volumeNameBuffer.value
            else:
                drive_names[d[:2].upper()] = "No Name"

    return drive_names




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
def time_booted(booted_time_timestamp):
    global previous_time

    current = time.time()
    dt1 = datetime.fromtimestamp(booted_time_timestamp)
    dt2 = datetime.fromtimestamp(current) 
    rd = dateutil.relativedelta.relativedelta (dt2, dt1)
    hours = rd.hours
    minutes = rd.minutes
    seconds = rd.seconds

    if int(minutes) <10:
        minutes = "0" + str(minutes)
    if int(seconds) <10:
        seconds = "0" + str(seconds)
    if rd.days >0:
        if int(hours) <10:
            hours = "0" + str(hours)
            print(f"{rd.days}:{hours}:{minutes}:{seconds}")
            return f"{rd.days}:{hours}:{minutes}:{seconds}"
        else:
            print(f"{rd.days}:{hours}:{minutes}:{seconds}")
            return f"{rd.days}:{hours}:{minutes}:{seconds}"
    else:
        print(f"{hours}:{minutes}:{seconds}")
        return f"{hours}:{minutes}:{seconds}"



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
    ### add in another coinitilize???  
    
    systemencoding = windll.kernel32.GetConsoleOutputCP()
    systemencoding= f"cp{systemencoding}"
    process = subprocess.Popen(["powershell", "-Command", "Import-Module .\AudioDeviceCmdlets.dll;", command],stdout=subprocess.PIPE, shell=True, encoding=systemencoding)
    proc_stdout = process.communicate()[0]
    if output:
        proc_stdout = proc_stdout[proc_stdout.index("["):-1]
        return json.loads(proc_stdout) 




def getAllVoices():
    engine = pyttsx3.init()
    return engine.getProperty("voices")

def TextToSpeech(message, voicesChoics, volume=100, rate=100, output="Default"):
    engine = pyttsx3.init()
    voices = engine.getProperty('voices')
    engine.setProperty("volume", volume/100)
    engine.setProperty("rate", rate)
    for voice in getAllVoices():
        if voice.name == voicesChoics:
            engine.setProperty('voice', voice.id)
            print("Using", voice.name, "voices", voice)

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
            print("EXCEPTION ERROR(Line:880)", e)

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
        
        
"""      DISPLAY CHANGE STUFF       """
#DISPLAY_DEVICE.StateFlags
DISPLAY_DEVICE_ACTIVE = 0x1
DISPLAY_DEVICE_MULTI_DRIVER = 0x2
DISPLAY_DEVICE_PRIMARY_DEVICE = 0x4
DISPLAY_DEVICE_MIRRORING_DRIVER = 0x8
DISPLAY_DEVICE_VGA_COMPATIBLE = 0x10
DISPLAY_DEVICE_REMOVABLE = 0x20
DISPLAY_DEVICE_DISCONNECT = 0x2000000
DISPLAY_DEVICE_REMOTE = 0x4000000
DISPLAY_DEVICE_MODESPRUNED = 0x8000000

#EnumDisplaySettingsEx.iModeNum
ENUM_CURRENT_SETTINGS = -1
ENUM_REGISTRY_SETTINGS = -2

#ChangeDisplaySettingsEx.dwflags
CDS_NONE                 = 0x00000000
CDS_UPDATEREGISTRY       = 0x00000001
CDS_TEST                 = 0x00000002
CDS_FULLSCREEN           = 0x00000004
CDS_GLOBAL               = 0x00000008
CDS_SET_PRIMARY          = 0x00000010
CDS_VIDEOPARAMETERS      = 0x00000020
CDS_ENABLE_UNSAFE_MODES  = 0x00000100
CDS_DISABLE_UNSAFE_MODES = 0x00000200
CDS_RESET                = 0x40000000
CDS_RESET_EX             = 0x20000000
CDS_NORESET              = 0x10000000


primary_device = None
primary_settings = None

def change_primary(monitornum):
    
    monitornum = monitornum.split(":")[0]
    primary = rf"\\.\DISPLAY{monitornum}"
    print("After splitting", monitornum)
    
    # Find the settings of the new primary display
    i = 0
    while True:
        try:
            device = win32api.EnumDisplayDevices(None, i)
        except pywintypes.error:
            break
        
        if device.DeviceName == primary:
            primary_device = device
            primary_settings = win32api.EnumDisplaySettingsEx(device.DeviceName, ENUM_CURRENT_SETTINGS, 0)
            break
            
        i += 1
            
    # Update all the positions of the displays relative to the new primary display
    i = 0
    while True:
        try:
            device = win32api.EnumDisplayDevices(None, i)
        except pywintypes.error:
            break
        
        if device.StateFlags & DISPLAY_DEVICE_ACTIVE != 0:
           # print (device.DeviceName)
            settings = win32api.EnumDisplaySettingsEx(device.DeviceName, ENUM_CURRENT_SETTINGS, 0)
           # print (settings.PelsWidth, "x", settings.PelsHeight, "@", settings.Position_x, ",", settings.Position_y)
            
            settings.Position_x -= primary_settings.Position_x
            settings.Position_y -= primary_settings.Position_y
            
            # CDS_UPDATEREGISTRY | CDS_NORESET = Don't make the changes yet until we've updated all the displays
            if device.DeviceName == primary:
                win32api.ChangeDisplaySettingsEx(device.DeviceName, settings, CDS_SET_PRIMARY | CDS_UPDATEREGISTRY | CDS_NORESET)
            else:
                win32api.ChangeDisplaySettingsEx(device.DeviceName, settings, CDS_UPDATEREGISTRY | CDS_NORESET)
            
        i += 1
            
    ## Update the displays with the registry settings
    win32api.ChangeDisplaySettingsEx(None, None)


"""MAGNIFIER STUFF"""


class WindowsRegistry:
    """Class WindowsRegistry is using for easy manipulating Windows registry.


    Methods
    -------
    query_value(full_path : str)
        Check value for existing.

    get_value(full_path : str)
        Get value's data.

    set_value(full_path : str, value : str, value_type='REG_SZ' : str)
        Create a new value with data or set data to an existing value.

    delete_value(full_path : str)
        Delete an existing value.

    query_key(full_path : str)
        Check key for existing.

    delete_key(full_path : str)
        Delete an existing key(only without subkeys).


    Examples:
        WindowsRegistry.set_value('HKCU/Software/Microsoft/Windows/CurrentVersion/Run', 'Program', r'"c:\Dir1\program.exe"')
        WindowsRegistry.delete_value('HKEY_CURRENT_USER/Software/Microsoft/Windows/CurrentVersion/Run/Program')
    """
    @staticmethod
    def __parse_data(full_path):
        full_path = re.sub(r'/', r'\\', full_path)
        hive = re.sub(r'\\.*$', '', full_path)
        if not hive:
            raise ValueError('Invalid \'full_path\' param.')
        if len(hive) <= 4:
            if hive == 'HKLM':
                hive = 'HKEY_LOCAL_MACHINE'
            elif hive == 'HKCU':
                hive = 'HKEY_CURRENT_USER'
            elif hive == 'HKCR':
                hive = 'HKEY_CLASSES_ROOT'
            elif hive == 'HKU':
                hive = 'HKEY_USERS'
        reg_key = re.sub(r'^[A-Z_]*\\', '', full_path)
        reg_key = re.sub(r'\\[^\\]+$', '', reg_key)
        reg_value = re.sub(r'^.*\\', '', full_path)

        return hive, reg_key, reg_value

    @staticmethod
    def query_value(full_path):
        value_list = WindowsRegistry.__parse_data(full_path)
        try:
            opened_key = winreg.OpenKey(getattr(winreg, value_list[0]), value_list[1], 0, winreg.KEY_READ)
            winreg.QueryValueEx(opened_key, value_list[2])
            winreg.CloseKey(opened_key)
            return True
        except WindowsError:
            return False

    @staticmethod
    def get_value(full_path):
        value_list = WindowsRegistry.__parse_data(full_path)
        try:
            opened_key = winreg.OpenKey(getattr(winreg, value_list[0]), value_list[1], 0, winreg.KEY_READ)
            value_of_value, value_type = winreg.QueryValueEx(opened_key, value_list[2])
            winreg.CloseKey(opened_key)
            return value_of_value
        except WindowsError:
            return None

    @staticmethod
    def set_value(full_path, value, value_type='REG_SZ'):
        value_list = WindowsRegistry.__parse_data(full_path)
        try:
            winreg.CreateKey(getattr(winreg, value_list[0]), value_list[1])
            opened_key = winreg.OpenKey(getattr(winreg, value_list[0]), value_list[1], 0, winreg.KEY_WRITE)
            winreg.SetValueEx(opened_key, value_list[2], 0, getattr(winreg, value_type), value)
            winreg.CloseKey(opened_key)
            return True
        except WindowsError:
            return False
        
    @staticmethod
    def query_key(full_path):
        value_list = WindowsRegistry.__parse_data(full_path)
        try:
            opened_key = winreg.OpenKey(getattr(winreg, value_list[0]), value_list[1] + r'\\' + value_list[2], 0, winreg.KEY_READ)
            winreg.CloseKey(opened_key)
            return True
        except WindowsError:
            return False


  #  @staticmethod
  #  def delete_value(full_path):
  #      value_list = WindowsRegistry.__parse_data(full_path)
  #      try:
  #          opened_key = winreg.OpenKey(getattr(winreg, value_list[0]), value_list[1], 0, winreg.KEY_WRITE)
  #          winreg.DeleteValue(opened_key, value_list[2])
  #          winreg.CloseKey(opened_key)
  #          return True
  #      except WindowsError:
  #          return False


  # @staticmethod
  # def delete_key(full_path):
  #     value_list = WindowsRegistry.__parse_data(full_path)
  #     try:
  #         winreg.DeleteKey(getattr(winreg, value_list[0]), value_list[1] + r'\\' + value_list[2])
  #         return True
  #     except WindowsError:
  #         return False
        
        
 
magpath = r"HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\\"
is_on=WindowsRegistry.get_value(magpath+"\RunningState")

def mag_on():
    """Check if Magnify is on
    - If magnify is off, this turns it on. 
       Always returns True
    """
    is_on=WindowsRegistry.get_value(magpath+"\RunningState")
    if is_on:
        return True
    if not is_on:
        out('magnify')
        return True
    
    
def mag_mode(mode=3):
    """ 
    -   1 = Docked
    -   2 = Full Screen
    -   3 = Lens
    """
    exists = WindowsRegistry.query_value(magpath + "MagnificationMode")
    if exists:
        WindowsRegistry.set_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\MagnificationMode", mode, value_type='REG_DWORD')



def mag_invert():
    """ Magnifier Invert
    - Toggles On / Off
    """
    exists = WindowsRegistry.query_value(magpath + "Invert")
    if exists:
        current = WindowsRegistry.get_value(magpath + "Invert")
        if current == 1:
            WindowsRegistry.set_value(r"HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\Invert", 0, value_type='REG_DWORD')
        if current ==0:
            WindowsRegistry.set_value(r"HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\Invert", 1, value_type='REG_DWORD')

            
def mag_increments(amount):
    """Setting Zoom Increments - MAX 400%, MIN 5%"""
    exists = WindowsRegistry.query_value(magpath + "ZoomIncrement")
    if exists:
        if amount:
            WindowsRegistry.set_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\ZoomIncrement", amount, value_type='REG_DWORD')
    
def mag_level(amount, onhold=False):
    """Set Zoom Levels 
    - 1600 is MAX (overriden to 1000)
    - Amount = Zoom Set
    - On Hold Amount = Increment Zoom
    """
    print("we in..", amount)
    themin=20
    themax=1000
    current = WindowsRegistry.get_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\Magnification")
    exists = WindowsRegistry.query_value(magpath + "Magnification")
    if not onhold:
         if exists:
             if amount:
                    if amount >themax:
                        mag_on()
                        WindowsRegistry.set_value(r"HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\Magnification", themax, value_type='REG_DWORD')
                    else:
                         mag_on()
                         WindowsRegistry.set_value(r"HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\Magnification", amount, value_type='REG_DWORD')
    if onhold:
        if exists:
            if current < themin:
                 print("less than 20")
                 WindowsRegistry.set_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\Magnification", 20, value_type='REG_DWORD')
            elif current > themax:
                try:
                    WindowsRegistry.set_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\Magnification", 1600, value_type='REG_DWORD')
                except:
                    pass
            else:
                try:
                    WindowsRegistry.set_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\Magnification", current + amount, value_type='REG_DWORD')
                except:
                    pass
            
def text_smoothing(switch):
    """Text Smoothing
    - Switch = On / Off
    """
    exists = WindowsRegistry.query_value(magpath + "UseBitmapSmoothing")
    if exists:
        if switch =="On":
            WindowsRegistry.set_value(r"HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\UseBitmapSmoothing", 1, value_type='REG_DWORD')
        elif switch == "Off":
            WindowsRegistry.set_value(r"HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\UseBitmapSmoothing", 0, value_type='REG_DWORD')


def magnifer_dimensions(x=None, y=None, onhold=False):
    """Can set a MIN/MAX Value to avoid this
    -  onhold = int value increment
    """
    exists = WindowsRegistry.query_value(magpath + "LensWidth")
    themax=95
    themin=20
    if not onhold:
        if x:
            if exists:
                WindowsRegistry.set_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\LensWidth", x, value_type='REG_DWORD')
        if y:
            if exists:
                WindowsRegistry.set_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\LensHeight", y, value_type='REG_DWORD')
                
    if onhold:
        if exists:
            if x:
                 current = WindowsRegistry.get_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\LensWidth")
                 if current < themin:
                     print("less than 20")
                     WindowsRegistry.set_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\LensWidth", 20, value_type='REG_DWORD')
                 elif current > themax:
                     WindowsRegistry.set_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\LensWidth", 95, value_type='REG_DWORD')
                 else:
                     WindowsRegistry.set_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\LensWidth", current + onhold, value_type='REG_DWORD')
            if y:
                 current = WindowsRegistry.get_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\LensHeight")
                 if current < themin:
                     print("less than 20")
                     WindowsRegistry.set_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\LensHeight", 20, value_type='REG_DWORD')
                 elif current > themax:
                     WindowsRegistry.set_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\LensHeight", 95, value_type='REG_DWORD')
                 else:
                     WindowsRegistry.set_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\LensHeight", current + onhold, value_type='REG_DWORD')







""" CLIPBOARD LISTENER STARTS IN MAIN.py onstart func"""
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Union, List, Optional

class Clipboard:
    @dataclass
    class Clip:
        type: str
        value: Union[str, List[Path]]

    def __init__(
            self,
            trigger_at_start: bool = False,
            on_text: Callable[[str], None] = None,
            on_update: Callable[[Clip], None] = None,
            on_files: Callable[[str], None] = None,
    ):
        self._trigger_at_start = trigger_at_start
        self._on_update = on_update
        self._on_files = on_files
        self._on_text = on_text

    def _create_window(self) -> int:
        """
        Create a window for listening to messages
        :return: window hwnd
        """
        wc = win32gui.WNDCLASS()
        wc.lpfnWndProc = self._process_message
        wc.lpszClassName = self.__class__.__name__
        wc.hInstance = win32api.GetModuleHandle(None)
        class_atom = win32gui.RegisterClass(wc)
        return win32gui.CreateWindow(class_atom, self.__class__.__name__, 0, 0, 0, 0, 0, 0, 0, wc.hInstance, None)

    def _process_message(self, hwnd: int, msg: int, wparam: int, lparam: int):
        WM_CLIPBOARDUPDATE = 0x031D
        if msg == WM_CLIPBOARDUPDATE:
            self._process_clip()
        return 0

    def _process_clip(self):
        clip = self.read_clipboard()
        if not clip:
            return

        if self._on_update:
            self._on_update(clip)
        if clip.type == 'text' and self._on_text:
            self._on_text(clip.value)
        elif clip.type == 'files' and self._on_text:
            self._on_files(clip.value)
    
    @staticmethod
    def read_clipboard() -> Optional[Clip]:
        try:
            win32clipboard.OpenClipboard()
            def get_formatted(fmt):
                if win32clipboard.IsClipboardFormatAvailable(fmt):
                    return win32clipboard.GetClipboardData(fmt)
                return None
            if files := get_formatted(win32con.CF_HDROP):
                return Clipboard.Clip('files', [Path(f) for f in files])
            elif text := get_formatted(win32con.CF_UNICODETEXT):
                TPClient.stateUpdate("KillerBOSS.TP.Plugins.capture.clipboard.contents", "TRUE")
                TPClient.stateUpdate("KillerBOSS.TP.Plugins.capture.clipboard.contents", Clipboard.Clip('text', text).value)
            elif text_bytes := get_formatted(win32con.CF_TEXT):
                TPClient.stateUpdate("KillerBOSS.TP.Plugins.capture.clipboard.contents", "TRUE")
                TPClient.stateUpdate("KillerBOSS.TP.Plugins.capture.clipboard.contents", Clipboard.Clip('text', text_bytes.decode()).value)
            elif bitmap_handle := get_formatted(win32con.CF_BITMAP):
                # TODO: handle screenshots
                pass
            return None
        finally:
            try:
                win32clipboard.CloseClipboard()
            except:
                pass

    def listen(self):
        if self._trigger_at_start:
            self._process_clip()

        def runner():
            hwnd = self._create_window()
            ctypes.windll.user32.AddClipboardFormatListener(hwnd)
            win32gui.PumpMessages()

        th = threading.Thread(target=runner, daemon=True)
        th.start()
        while th.is_alive():
            th.join(0.25)
            
            

def clipboard_listener():
   clipboard = Clipboard(on_update=print, trigger_at_start=False)
   clipboard.listen()
   



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

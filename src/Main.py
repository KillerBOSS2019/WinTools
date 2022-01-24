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
import win32con  ### needed to show window without issues

### 
# Origin Launcher Window = OriginWebHelperService
# Steam Launcher = Steam vguiPopupWindow
# Epic Launcher = Epic Games Launcher
# Discord = Discord Chrome_WidgetWin_1   ???   MIGHT NOT NEED THIS  shows other stuff before.

def check_process(process_name, shortcut ="", focus=True):
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
                win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)  ##updated n working
                win32gui.SetForegroundWindow(hwnd)
               #### break so it doesnt keep on triggering random process'
                break
                #print(f"{window_name:20.20} {class_name:20.20}")
    if not exist:
        print("load via shortcut")
        os.system('"' + shortcut + '"')
    
#check_process("Discord", shortcut=r"C:\Users\dbcoo\AppData\Local\Discord\Update.exe --processStart Discord.exe", focus=True)


def get_windows():
    results = []

    def winEnumHandler(hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetWindowText(hwnd):
                # print(win32gui.GetWindowText( hwnd ))
                results.append(win32gui.GetWindowText(hwnd))

    win32gui.EnumWindows(winEnumHandler, None)
    return results

#print(get_windows())


def copy_im_to_clipboard(image):
    bio = BytesIO()
    image.save(bio, 'BMP')
    data = bio.getvalue()[14:] # removing some headers
    bio.close()

    send_to_clipboard(win32clipboard.CF_DIB, data)


def send_to_clipboard(clip_type, data):
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


def check_number_of_monitors():
    with mss.mss() as sct:
        monitor_count = (len(sct.monitors))
        return monitor_count

### Call back on_exist not being used            
def on_exists(fname):
    # type: (str) -> None
    if os.path.isfile(fname):
        newfile = fname + ".old"
        print("{} -> {}".format(fname, newfile))
        os.rename(fname, newfile)
        

###screenshot window without bringing it to foreground 
def screenshot_window(capture_type, window_title=None, clipboard=False, save_location=None):
    from ctypes import windll
    hwnd = win32gui.FindWindow(None, window_title)
    
    ##  Change the line below depending on whether you want the whole window
    ##  May be needed in future?
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
    # or just the client area. 
                          # 1, 2, 3 all give different results
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
            
#screenshot_window(capture_type=3, window_title="Calculator", clipboard=False, save_location="testing2.png")

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
                    print("capturing all")  ## having to use temp file to capture all screens successfully?
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
        
old_results = []
old_volume_list = []
monitor_count_old = 0
global_states = []
global Timer
counter = 0
def updateStates():
    global old_volume_list, global_states, Timer, counter, monitor_count_old, old_results
    Timer = threading.Timer(0.4, updateStates)
    Timer.start()
    if running:
        
#        #### I feel like this should loop every 30 seconds at minimum
#        ### Updating Monitor States
#        monitor_list = []
#        monitor_count = check_number_of_monitors()
#        if str(monitor_count_old) == str(monitor_count):
#            pass
#        elif str(monitor_count_old) is not str(monitor_count):
#            monitor_count_old = monitor_count
#            for monitor_number in range(monitor_count):
#                if monitor_number == 0:
#                    monitor_list.append(str(monitor_number))
#                else:
#                    monitor_list.append(str(monitor_number))
#            TPClient.choiceUpdate("KillerBOSS.TP.Plugins.screencapture.monitors", monitor_list)
#
#        ### would like this to only check every 3-5 seconds for new windows
#        ### Getting Active Windows and updating choice state only when windows change.
#        global windows_active
#        windows_active = get_windows()
#          
#        if len(old_results) is not len(windows_active):
#            windows_active = get_windows()
#            print("Previous Count:", len(old_results), "New Count:", len(windows_active))
#            old_results = windows_active
#            TPClient.choiceUpdate("KillerBOSS.TP.Plugins.screencapture.window_name", windows_active)
#            TPClient.stateUpdate("KillerBOSS.TP.Plugins.Windows.activeCOUNT", str(len(windows_active)))
#        else:
#            windows_active = get_windows()
#            old_results = windows_active
#        ##end of changes
        
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
            output = AudioDeviceCmdlets('Get-AudioDevice -List | ConvertTo-Json')
            for x in output:
                if x['Type'] == "Playback" and x['Default'] == True:
                    TPClient.stateUpdate('KillerBOSS.TP.Plugins.Sound.CurrentOutputDevice', x['Name'])
                elif x['Type'] == "Recording" and x['Default'] == True:
                    TPClient.stateUpdate('KillerBOSS.TP.Plugins.Sound.CurrentInputDevice', x['Name'])
                    
                    ###
        
            #### I feel like this should loop every 30 seconds at minimum
            ### Updating Monitor States
            monitor_list = []
            monitor_count = check_number_of_monitors()
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

            ### would like this to only check every 3-5 seconds for new windows
            ### Getting Active Windows and updating choice state only when windows change.
            global windows_active
            windows_active = get_windows()

            if len(old_results) is not len(windows_active):
                windows_active = get_windows()
                print("Previous Count:", len(old_results), "New Count:", len(windows_active))
                old_results = windows_active
                TPClient.choiceUpdate("KillerBOSS.TP.Plugins.screencapture.window_name", windows_active)
                TPClient.stateUpdate("KillerBOSS.TP.Plugins.Windows.activeCOUNT", str(len(windows_active)))
            else:
                windows_active = get_windows()
                old_results = windows_active
        ##end of changes
        
        
        
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
    global running
    running = True
    print(data)
    updateStates()

    
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
            
    if data['actionId'] == "KillerBOSS.TP.Plugins.screencapture.full.clipboard":
        if data['data'][0]['value'] == "ALL DISPLAYS":
            screenshot_monitor(monitor_number=0, clipboard=True)
            screenshot_monitor(0, "test_file", clipboard = True)
            print("FUCKIN EH")
        else:
            shortened = monitor_number=data['data'][0]['value'].replace("Display #", "")
            screenshot_monitor(monitor_number=int(shortened),clipboard=True)
            print("else...")
        
    if data['actionId'] == "KillerBOSS.TP.Plugins.screencapture.window.clipboard":
        screenshot_window(capture_type=int(data['data'][1]['value']), window_title=data['data'][0]['value'], clipboard=True)
        
        
        ###using wildcard to clipboard
    if data['actionId'] == "KillerBOSS.TP.Plugins.screencapture.window.clipboard.wildcard":
        for thing in windows_active:
            if data['data'][0]['value'].lower() in thing.lower():
                print("We found", thing)
                screenshot_window(capture_type=int(data['data'][1]['value']), window_title=thing, clipboard=True)
                break
        
               ###using wildcard to FILE
    if data['actionId'] == "KillerBOSS.TP.Plugins.screencapture.window.file.wildcard":
        for thing in windows_active:
            if data['data'][0]['value'].lower() in thing.lower():
                print("We found", thing)
                afile_name = (data['data'][2]['value']) +"/" +(data['data'][3]['value']) 
                print(afile_name)    
                screenshot_window(capture_type=int(data['data'][1]['value']), window_title=thing, clipboard=False, save_location=afile_name)
                break
             
    if data['actionId'] == "KillerBOSS.TP.Plugins.screencapture.full.file":
        afile_name = (data['data'][1]['value']) +"/" +(data['data'][2]['value'])        
        screenshot_monitor(monitor_number=int(data['data'][0]['value']), filename=afile_name, clipboard=False)
        
    if data['actionId'] == "KillerBOSS.TP.Plugins.screencapture.window.file":
        afile_name = (data['data'][2]['value']) +"/" +(data['data'][3]['value']) 
        print(afile_name)       
        screenshot_window(capture_type=int(data['data'][1]['value']), window_title=data['data'][0]['value'], clipboard=False, save_location=afile_name)
        
        
            ###using wildcard to FILE
    if data['actionId'] == "KillerBOSS.TP.Plugins.screencapture.processcheck":
        app_to_check = data['data'][0]['value'].lower()
        focus = data['data'][1]['value']
        focus_type = data['data'][2]['value']
        shortcut_to_open = data['data'][3]['value']
        if focus == "Focus":
            focus_check = True
        else:
            focus_check = False
        
        check_process(app_to_check, shortcut_to_open, focus=focus_check)
        
        
    
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
    TPClient.disconnect()
    running = False
    Timer.cancel()
    sys.exit()
TPClient.connect()

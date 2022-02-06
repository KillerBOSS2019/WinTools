import json
import os
import sys
import threading
import time
from time import sleep

### Gitago Imports
import mss
import mss.tools
import psutil
import pyautogui
import pygetwindow
import pythoncom
import schedule
import TouchPortalAPI
from PIL import Image
from pycaw.pycaw import AudioUtilities
from TouchPortalAPI import TYPES
from screeninfo import get_monitors

from utils.util import *


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
        
        
def timebooted_loop():
    TPClient.stateUpdate("KillerBOSS.TP.Plugins.Windows.livetime", str(time_booted()))

def vd_check():
    vdlist=[]
    virtual_desk_count = len(get_virtual_desktops())
    vdlist.append("Next")
    vdlist.append("Previous")
    for i in range (virtual_desk_count):
        vdlist.append(str(i))
    TPClient.choiceUpdate("KillerBOSS.TP.Plugins.virtualdesktop.actionchoice", vdlist)
        


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
                except PermissionError as e:
                            # this can be catched due to the disk that
                            # isn't ready
                    print("Permission error " + e)
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
    except Exception as e:
        print("Disk usage", e)
     
     ### Total Read/Write since boot       
    # get IO statistics since boot
    try:
        disk_io = psutil.disk_io_counters()
        # print(f"Total read: {get_size(disk_io.read_bytes)}")
        # print(f"Total write: {get_size(disk_io.write_bytes)}")
    
        network = network_usage()
        # print(f"Total Bytes Sent: {network['received']}")
        # print(f"Total Bytes Received: {network['sent']}")
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
    except Exception as e:
        print("error disk usage n stuff " + e)

old_results = []
def get_windows_update():
    global windows_active, old_results
    windows_active = get_windows()
    if len(old_results) is not len(windows_active):
        # windows_active = get_windows()
        print("Previous Count:", len(old_results), "New Count:", len(windows_active))
        old_results = windows_active
        TPClient.choiceUpdate("KillerBOSS.TP.Plugins.screencapture.window_name", windows_active)
        TPClient.stateUpdate("KillerBOSS.TP.Plugins.Windows.activeCOUNT", str(len(windows_active)))
    else:
        windows_active = get_windows()
        old_results = windows_active


monitor_count_old = ""
def check_number_of_monitors():
    global monitor_count_old
    mon_length = len(get_monitors())   ### Wonder if triggering this each time to get length of monitors is better / less resources than using get_monitors2 ?     this uses screeninfo module
    
    if monitor_count_old != mon_length:
        list_monitor_full = get_monitors2()
        TPClient.choiceUpdate("KillerBOSS.TP.Plugins.winsettings.monchoice", list_monitor_full)
        TPClient.choiceUpdate("KillerBOSS.TP.Plugins.winsettings.primary_monitor_choice", list_monitor_full)
        list_monitor_full.insert(0, "0: ALL MONITORS")
        TPClient.choiceUpdate("KillerBOSS.TP.Plugins.screencapture.monitors", list_monitor_full)  #  KillerBOSS.TP.Plugins.screencapture.full.file 
        monitor_count_old = mon_length
        
        





def screenshot_monitor(monitor_number, filename="", clipboard = False):   
    monitor_number = int(monitor_number.split(":")[0])
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
stop_run_continuously = run_continuously()

# Setup TouchPortal connection
TPClient = TouchPortalAPI.Client('Windows-Tools')

TTSThread = threading.Thread(target=TextToSpeech)
running = False
updateXY = True

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
            print("error on Pythoncom", e)
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
                        except Exception as e:
                            print("exception at 293", e)
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
        if counter%2 == 0:
            """
            Updates every 5 loops 
            """
            #TPClient.stateUpdate("KillerBOSS.TP.Plugins.Windows.livetime", str(time_booted()))
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
        if counter >= 35:
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
        
    ### making this trigger every 1 seconds forever...
    ##happening in handleSettings intead...
    #schedule.every(1).seconds.do(timebooted_loop)
     
    ### Getting Power Plan Details and Updating States and Choices
    pplans = get_powerplans()
    pplans_list =[]
    for item in pplans:
        pplans_list.append(item)
    TPClient.createState("KillerBOSS.TP.Plugins.winsettings.powerplan_current", "Current Power Plan", get_powerplans(currentcheck=True))
    TPClient.choiceUpdate("KillerBOSS.TP.Plugins.winsettings.powerplan_choice", pplans_list)
    
    ### Updating Choices for Windows Settings options from util.py
    TPClient.choiceUpdate("KillerBOSS.TP.Plugins.winsettings.choice", activate_windows_setting())
    
    try:
        import requests
        pub_ip = requests.get('https://checkip.amazonaws.com').text.strip()  
        TPClient.createState("KillerBOSS.TP.Plugins.winsettings.winsettings.publicIP", "Your Public IP", pub_ip)
    except:
        pass
    
    the_devices = getAllOutput_TTS2()
    tts_outputs = []
    for item in the_devices:
        tts_outputs.append(item)
        
    TPClient.choiceUpdate("KillerBOSS.TP.Plugins.TextToSpeech.output", tts_outputs)
        
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

    for scheduleFunc in [(disk_usage, "Update Interval: Hard Drive"),
                         (vd_check, "Update Interval: Network Up/Down(INCOMPLETE)"),
                         (check_number_of_monitors, "Update Interval: Active Monitors"),
                         (get_windows_update, "Update Interval: Active Windows")]:
        interval = float(newsettings[scheduleFunc[1]])
        
        if int(newsettings[scheduleFunc[1]]) > 0:
            schedule.every(interval).seconds.do(scheduleFunc[0])
            print(f"{scheduleFunc[1]} is now {interval}")
        else:
            print(f"{scheduleFunc[1]} is TURNED OFF")
            
    schedule.every(1).seconds.do(timebooted_loop)
    stop_run_continuously = run_continuously()
    return settings


# Settings handler
@TPClient.on(TouchPortalAPI.TYPES.onSettingUpdate)
def onSettingUpdate(data):
    if (settings := data.get('values')):
        handleSettings(settings, False)


    
@TPClient.on(TouchPortalAPI.TYPES.onAction)
def Actions(data):
    global TTSThread
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
            
## Screencap Window Current
    if data['actionId'] == "KillerBOSS.TP.Plugins.window.current":
        if data['data'][0]['value'] == "Clipboard":
            current_window = pygetwindow.getActiveWindowTitle()
            screenshot_window(capture_type=3, window_title=current_window, clipboard=True)
            
        elif data['data'][0]['value'] == "File":
            current_window = pygetwindow.getActiveWindowTitle()
            afile_name = (data['data'][1]['value']) +"/" +(data['data'][2]['value']) 
            screenshot_window(capture_type=3, window_title=current_window, clipboard=False, save_location=afile_name)
            
##Screen Cap Window Wildcard to FILE
    if data['actionId'] == "KillerBOSS.TP.Plugins.screencapture.window.file.wildcard":
        global windows_active
        windows_active = get_windows()
        for thing in windows_active:
            if data['data'][0]['value'].lower() in thing.lower():
                print("We found", thing)
                if data['data'][4]['value'] == "Clipboard":
                    screenshot_window(capture_type=int(data['data'][1]['value']), window_title=thing, clipboard=True)
                    
                elif data['data'][4]['value'] == "File":
                    print("File stuf")
                    afile_name = (data['data'][2]['value']) +"/" +(data['data'][3]['value']) 
                    screenshot_window(capture_type=int(data['data'][1]['value']), window_title=thing, clipboard=False, save_location=afile_name)
                break
             
    if data['actionId'] == "KillerBOSS.TP.Plugins.screencapture.full.file":   
        if data['data'][1]['value'] == "Clipboard":
            try:
                screenshot_monitor(monitor_number=(data['data'][0]['value']), clipboard=True)
            except:
                pass
        elif data['data'][1]['value'] == "File":
            try:
                afile_name = (data['data'][2]['value']) +"/" +(data['data'][3]['value'])    
                screenshot_monitor(monitor_number=(data['data'][0]['value']), filename=afile_name, clipboard=False)
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
        
    if data['actionId'] == "KillerBOSS.TP.Plugins.winsettings.rotate_display":
        if data['data'][0]['value'] != "Pick a Monitor":
            rotate_display(int(data['data'][0]['value']), data['data'][1]['value'])
        
        
    if data['actionId'] == "KillerBOSS.TP.Plugins.winsettings.shutdown":
        win_shutdown(data['data'][0]['value'])
        
    if data['actionId'] == "KillerBOSS.TP.Plugins.winsettings.primary_monitor":
        change_primary(data['data'][0]['value'])
        
        
    if data['actionId'] == "KillerBOSS.TP.Plugins.winsettings.powerplan":
        change_pplan(data['data'][0]['value'])
        TPClient.stateUpdate("KillerBOSS.TP.Plugins.winsettings.powerplan_current", get_powerplans(currentcheck=True))
        
        
    if data['actionId'] == "KillerBOSS.TP.Plugins.winsettings.move_window":
        move_win_button(data['data'][0]['value'])
        
    if data['actionId'] == "KillerBOSS.TP.Plugins.virtualdesktop.actions.move_window":
        choice = data['data'][0]['value']
        if data['data'][1]['value'] == "False":
            virtual_desktop(move=True, target_desktop=choice)
        if data['data'][1]['value'] == "True":
            virtual_desktop(move=True, target_desktop=choice, pinned=True)
        
    if data['actionId'] == "KillerBOSS.TP.Plugins.magnifier.actions":
        magnifier(data['data'][0]['value'])
        
    if data['actionId'] == "KillerBOSS.TP.Plugins.toast.create":
        win_toast(atitle=data['data'][0]['value'], amsg=data['data'][1]['value'], aduration=data['data'][2]['value'], icon=data['data'][5]['value'], buttonText=data['data'][3]['value'], buttonlink=data['data'][4]['value'], sound=data['data'][6]['value'])
        
    if data['actionId'] == "KillerBOSS.TP.Plugins.winextra.emojipanel":
        winextra("Emoji")
        
    if data['actionId'] == "KillerBOSS.TP.Plugins.winextra.keyboard":
        winextra("Keyboard")
        
    if data['actionId'] == "KillerBOSS.TP.Plugins.winsettings.action":
        activate_windows_setting(data['data'][0]['value'])

    if data['actionId'] == "KillerBOSS.TP.Plugins.TextToSpeech.speak":
        #print(data['data'][3]['value'])
        TTSThread = threading.Thread(target=TextToSpeech, args=(data['data'][1]['value'], data['data'][0]['value'], int(data['data'][2]['value']), int(data['data'][3]['value']), data['data'][4]['value']))
        TTSThread.setDaemon(True)
        TTSThread.start()
            
    if data['actionId'] == "KillerBOSS.TP.Plugins.TextToSpeech.stop":
        if TTSThread.is_alive():
            sd.stop()
            pass
        



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

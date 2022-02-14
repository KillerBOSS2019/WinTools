from cmath import e
from gc import callbacks
import json
import os
import sys
import threading
import time
from time import sleep
from utils.util import *

### Gitago Imports
import mss
import mss.tools
import psutil
import pyautogui
import pygetwindow
import schedule
import TouchPortalAPI
from PIL import Image
from pycaw.pycaw import AudioUtilities
from TouchPortalAPI import TYPES
from screeninfo import get_monitors
from pycaw.magic import MagicManager, MagicSession  # isort:skip
from pycaw.constants import AudioSessionState  # isort:skip



##########################################################################
"""                   THE START OF WIN CALLBACK THINGS               """ 
global_states = []
first_time=True
class AppRow(MagicSession):
    """
    Row for each app in mixer.
    handles refreshing the gui if session is changed external.
    handles user input and changing session volume/mute.
    """
    
    def __init__(self):
        global global_states, first_time
        super().__init__(volume_callback=self.update_volume,
                         mute_callback=self.update_mute,
                         state_callback=self.update_state)
        if first_time:
            global_states.append("Master Volume")
            global_states.append("Current app")
            first_time = False
        
        # ______________ DISPLAY NAME ______________
        self.app_name = self.magic_root_session.app_exec
        print(f":: new session: {self.app_name}")
        
        # set initial:
        self.update_mute(self.mute)
        # set initial:
        self.update_state(self.state)
        self.update_volume(self.volume)
        

    def update_state(self, new_state):
        """
        when status changed
        (see callback -> AudioSessionEvents -> OnStateChanged)
        """
        if new_state == AudioSessionState.Inactive:
            # AudioSessionStateInactive
            """Sesssion Has Expired"""
            print(f"{self.app_name} not active")
    
        elif new_state == AudioSessionState.Active:
            """Check if created, if not it will create"""
            print(f"{self.app_name} is an Active Session")
    
        elif new_state == AudioSessionState.Expired:
            """Removing Expired States"""
            check_states(self.app_name, remove=True)

    
    def update_volume(self, new_volume):
        """
        when volume is changed externally
        (see callback -> AudioSessionEvents -> OnSimpleVolumeChanged )
        """
        new_volume2 = str(new_volume)  #### THIS NEEDS FIXED PROPER - I WAS BAMBOOZLED
        if new_volume2[0:3] == "1.0":
            TPClient.stateUpdate(f'KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{self.app_name}', new_volume2[2:4])
            TPClient.send({
                "type":"connectorUpdate",
                "connectorId":f"pc_Windows-Tools_KillerBOSS.TP.Plugins.VolumeMixer.connectors.APPcontrol|KillerBOSS.TP.Plugins.VolumeMixer.slidercontrol={self.app_name}",
                "value": 100
            })
        else:
            TPClient.stateUpdate(f'KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{self.app_name}', new_volume2[2:4])
            TPClient.send({
                "type":"connectorUpdate",
                "connectorId":f"pc_Windows-Tools_KillerBOSS.TP.Plugins.VolumeMixer.connectors.APPcontrol|KillerBOSS.TP.Plugins.VolumeMixer.slidercontrol={self.app_name}",
                "value": new_volume2[2:4]
            })
        print("[{}] 🔈:{}".format(self.app_name.split(".")[0], new_volume))

        
    def update_mute(self, muted):
        """ when mute state is changed by user or through other app """
       # app_name = self.app_name.split(".")[0]
        check_states(self.app_name)
        if muted:
            print(f"{self.app_name} is unmuted")
            TPClient.stateUpdate(f"KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{self.app_name}.muteState", "Muted")
        else:
            print(f"{self.app_name} is muted")
            TPClient.stateUpdate(f"KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{self.app_name}.muteState", "Un-Muted")
        


"""                     THE END OF CALLBACK                            """
##########################################################################

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
        



def check_states(app_name, remove=False):
    global global_states
    if app_name not in global_states:
        TPClient.createStateMany([
                {
                    "id": f'KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{app_name}.muteState',
                    "desc": f"{app_name} Mute State",
                    "value": ""
                },
                {
                    "id": f'KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{app_name}',
                    "desc": f"{app_name} Volume",
                    "value": ""
                }
                ])
        global_states.append(app_name)

        """UPDATING CHOICES WITH GLOBALS"""
        TPClient.choiceUpdate('KillerBOSS.TP.Plugins.VolumeMixer.Increase/DecreaseVolume.process', global_states)
        TPClient.choiceUpdate('KillerBOSS.TP.Plugins.VolumeMixer.Mute/Unmute.process', global_states)
        TPClient.choiceUpdate("KillerBOSS.TP.Plugins.VolumeMixer.slidercontrol", global_states)
        return True
    
    
    if remove:
        """     Deleting Volume States As Needed     """
        TPClient.removeStateMany([f"KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{app_name}.muteState", f"KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{app_name}"])
        global_states.remove(app_name)
        
        TPClient.choiceUpdate('KillerBOSS.TP.Plugins.VolumeMixer.Increase/DecreaseVolume.process', global_states)
        TPClient.choiceUpdate('KillerBOSS.TP.Plugins.VolumeMixer.Mute/Unmute.process', global_states)
        TPClient.choiceUpdate("KillerBOSS.TP.Plugins.VolumeMixer.slidercontrol", global_states)
        print(f":: closed session: {app_name}")



def timebooted_loop():
    """This is on threaded loop instead of schedule so its more consistant"""
    while True:
        TPClient.stateUpdate("KillerBOSS.TP.Plugins.Windows.livetime", str(time_booted()))
        time.sleep(1)

def vd_check():
    vdlist=[]
    virtual_desk_count = len(get_virtual_desktops())
    vdlist.append("Next")
    vdlist.append("Previous")
    for i in range (virtual_desk_count):
        vdlist.append(str(i))
    TPClient.choiceUpdate("KillerBOSS.TP.Plugins.virtualdesktop.actionchoice", vdlist)
        
        
        ## should we bother checking old IP info to new to see if its different before we update states?
def get_ip_loop():
    try:
        pub_ip =get_ip_details("choice1")  
        for item in pub_ip:
            TPClient.stateUpdate(f"KillerBOSS.TP.Plugins.publicip.{item}", pub_ip[item])
    except:
        pass
    
    
global count
count = 0
def disk_usage():
    global count
    thedrivenames = get_win_drive_names2()
    partitions = psutil.disk_partitions()
    drives = False
    count = count + 1
    print(f"================[  {count}   ]==============")
    try:
        for partition in partitions:
            the_partition = partition.device.split(":")
            driveletter = (the_partition[0])
            if not drives:
                if partition.mountpoint[:2] in thedrivenames:
                    the_drive_name = thedrivenames[partition.mountpoint[:2]]
                    partition_usage = psutil.disk_usage(partition.mountpoint)
                    percentage = 100 - partition_usage.percent
                    str_percent = str(percentage)
                    str_percent = str_percent[0:4]

                    freespace = get_size(partition_usage.free).replace("GB","")
                    usedspace = get_size(partition_usage.used).replace("GB","")

                    driveletter = partition.mountpoint[:1]
                  #  print("Drive Letter", driveletter, "  ", the_drive_name)
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
                        "value": the_drive_name
                    },
                    {
                        "id": f'KillerBOSS.TP.Plugins.Windows.drive.size_{driveletter}',
                        "value": get_size(partition_usage.total)
                    },
                    {
                        "id": f'KillerBOSS.TP.Plugins.Windows.drive.free_{driveletter}',
                        "value": str(freespace)
                    },
                    {
                        "id": f'KillerBOSS.TP.Plugins.Windows.drive.used_{driveletter}',
                        "value": str(usedspace)
                    },
                    {
                        "id": f'KillerBOSS.TP.Plugins.Windows.drive.percent_{driveletter}',
                        "value": str(str_percent)
                    },
                    ])
    except Exception as e:
        print("Disk usage", e)
     
     ### Total Read/Write since boot       
    # get IO statistics since boot
    try:
        disk_io = psutil.disk_io_counters()

        network = network_usage()

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
                    
                 
                elif monitor_number != 0:
                    img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")  ##   Instead of making a temp file we get it direct from raw to clipboard
                    copy_im_to_clipboard(img)
                  #  TPClient.stateUpdate("KillerBOSS.TP.Plugins.winsettings.winsettings.publicIP", getFrame_base64(img).decode())

            if clipboard == False:
                mss.tools.to_png(sct_img.rgb, sct_img.size, output=filename + ".png")
                print("Image saved -> "+ filename+ ".png" )
                
        except IndexError:
            print("This Monitor does not exist")
            

# Setup TouchPortal connection
TPClient = TouchPortalAPI.Client('Windows-Tools')

global Timer
stop_run_continuously = run_continuously() ### Schedule "Thread" ###
running = False
updateXY = True
counter = 0

def updateStates():
    global global_states, Timer, counter
    Timer = threading.Timer(0.4, updateStates)
    Timer.start()
    if running:
        counter = counter + 1
        if counter%2 == 0:
        #     print("STEP 2")
            """
            Updates every 5 loops 
            """
            #TPClient.stateUpdate("KillerBOSS.TP.Plugins.Windows.livetime", str(time_booted()))
            TPClient.stateUpdate("KillerBOSS.TP.Plugins.Application.currentFocusedAPP", pygetwindow.getActiveWindowTitle())
            activeWindow = getActiveExecutablePath()
            if activeWindow != None:
                activeWindow = os.path.basename(getActiveExecutablePath())

            """Checking for Default Input/Output Devices"""
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
                except UnicodeDecodeError as err:
                    print("Unicode Decode Error: ", err)

            """Advance Mouse"""
            if updateXY:
                TPClient.stateUpdate('KillerBOSS.TP.Plugins.AdvanceMouse.MousePos.X', str(pyautogui.position()[0]))
                TPClient.stateUpdate('KillerBOSS.TP.Plugins.AdvanceMouse.MousePos.Y', str(pyautogui.position()[1]))
                #  print("DONE")
            else:
                Timer.cancel()
                print('canceled the Timer')



"""============         ON START               ==========="""
running = False
#updateStates()
@TPClient.on('info')
def onStart(data):
    if settings := data.get('settings'):
        handleSettings(settings, False)

    """ Getting Powerplans and Updating Choice List + Current Power Plan Stte"""
    pplans = get_powerplans()
    pplans_list =[]
    for item in pplans:
        pplans_list.append(item)
    TPClient.stateUpdate("KillerBOSS.TP.Plugins.winsettings.powerplan_current", get_powerplans(currentcheck=True))
    TPClient.choiceUpdate("KillerBOSS.TP.Plugins.winsettings.powerplan_choice", pplans_list)  ### Updating Power Plan Choices


    """Updating Choices for Windows Settings options from util.py"""
    TPClient.choiceUpdate("KillerBOSS.TP.Plugins.winsettings.choice", activate_windows_setting())
    
    """Getting TTS Output Devices and Updating Choices"""
    voices = [voice.name for voice in getAllVoices()]
    TPClient.choiceUpdate("KillerBOSS.TP.Plugins.TextToSpeech.voices", voices)
    
    
    """Setting Default Audio Slider"""
    TPClient.send({
            "type":"connectorUpdate",
            "connectorId":"pc_Windows-Tools_KillerBOSS.TP.Plugins.VolumeMixer.connectors.APPcontrol|KillerBOSS.TP.Plugins.VolumeMixer.slidercontrol=Master Volume",
            "value": str(getMasterVolume())
            })
    
    global running
    running = True
    updateStates()
    

def run_callback():
    MagicManager.magic_session(AppRow)

"""============         SETTINGS          ==============="""
settings = {}
def handleSettings(settings, on_connect=False):
    newsettings = { list(settings[i])[0] : list(settings[i].values())[0] for i in range(len(settings)) }
    global stop_run_continuously
    
    ###Stopping Scheduled Tasks and Clearing List
    stop_run_continuously.set()
    schedule.clear()
    time.sleep(2)

    for scheduleFunc in [(disk_usage, "Update Interval: Hard Drive"),
                         (get_ip_loop, "Update Interval: Public IP"),
                         (check_number_of_monitors, "Update Interval: Active Monitors"),
                         (get_windows_update, "Update Interval: Active Windows"), 
                         (vd_check, "Update Interval: Active Virtual Desktops")]:
        interval = float(newsettings[scheduleFunc[1]])
        
        if int(newsettings[scheduleFunc[1]]) > 0:
            schedule.every(interval).seconds.do(scheduleFunc[0])
            print(f"{scheduleFunc[1]} is now {interval}")
        else:
            print(f"{scheduleFunc[1]} is TURNED OFF")
    
    """IF wanting Disk Usage we are Firing it First
        So Audio States Don't Mix Later"""
    if int(newsettings['Update Interval: Hard Drive']) > 0:
        disk_usage()

    """Mandatory Loops, TimeBooted needs moved to settings"""
    #schedule.every(5).minutes.do(get_ip_loop)
    
    th=threading.Thread(target=timebooted_loop)
    th.start()

    """Starting Schedule Again"""
    stop_run_continuously = run_continuously()
    pythoncom.CoInitialize()
    run_callback()
    return settings




"""          Settings Handler         """
@TPClient.on(TouchPortalAPI.TYPES.onSettingUpdate)
def onSettingUpdate(data):
    if (settings := data.get('values')):
        handleSettings(settings, False)


"""============         ACTIONS              =========="""
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
                #sleep(0.05)
        else:
            break



"""Getting Input and Output Devices and Updating Choices"""
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


import os
from time import sleep

import pygetwindowmp as pygetwindow
import pythoncom
import TouchPortalAPI
from pycaw.constants import AudioSessionState
from pycaw.magic import MagicManager  # ## These are used in main.py
from pycaw.magic import MagicSession
from TouchPortalAPI import TYPES

from utils import audioController, clipboard, util
import threading
from utils.magnifier import *
from utils.virtualDesktop import *
from utils.TextToSpeech import *
from utils.util import winextra, windowsSettings

TPClient = TouchPortalAPI.Client('Windows-Tools')

volumeprocess = ["Master Volume", "Current app"]
audio_exempt_list = []
running = False


def updateVolumeMixerChoicelist():
    TPClient.choiceUpdate('KillerBOSS.TP.Plugins.VolumeMixer.Increase/DecreaseVolume.process', volumeprocess)
    TPClient.choiceUpdate('KillerBOSS.TP.Plugins.VolumeMixer.Mute/Unmute.process', volumeprocess)
    TPClient.choiceUpdate("KillerBOSS.TP.Plugins.VolumeMixer.slidercontrol", volumeprocess)

def removeAudioState(app_name):
    global volumeprocess
    TPClient.removeStateMany([
            f"KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{app_name}.muteState",
            f"KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{app_name}",
            f"KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{app_name}.active"])
    volumeprocess.remove(app_name)
    updateVolumeMixerChoicelist() # Update with new changes

def audioStateManager(app_name):
    global volumeprocess
    print("AUDIO EXEMPT LIST", audio_exempt_list)

    if app_name not in volumeprocess:
        TPClient.createStateMany([
                {
                    "id": f'KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{app_name}.muteState',
                    "desc": f"{app_name} Mute State",
                    "parentGroup": "Audio process state",
                    "value": ""
                },
                {
                    "id": f'KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{app_name}',
                    "desc": f"{app_name} Volume",
                    "parentGroup": "Audio process state",
                    "value": ""
                },
                {
                    "id": f'KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{app_name}.active',
                    "desc": f"{app_name} Active",
                    "parentGroup": "Audio process state",
                    "value": ""
                },
                ])
        volumeprocess.append(app_name)

        """UPDATING CHOICES WITH GLOBALS"""
        updateVolumeMixerChoicelist()
        print(f"{app_name} State Added")

    """ Checking for Exempt Audio"""
    if app_name in audio_exempt_list:
        removeAudioState(app_name)

import ctypes

import psutil
import win32process
import pyautogui


def getActiveExecutablePath():
    hWnd = ctypes.windll.user32.GetForegroundWindow()
    if hWnd == 0:
        return None # Note that this function doesn't use GetLastError().
    else:
        _, pid = win32process.GetWindowThreadProcessId(hWnd)
        return psutil.Process(pid).exe()

class WinAudioCallBack(MagicSession):
    def __init__(self):
        super().__init__(volume_callback=self.update_volume,
                         mute_callback=self.update_mute,
                         state_callback=self.update_state)

        # ______________ DISPLAY NAME ______________
        self.app_name = self.magic_root_session.app_exec
        #print(f":: new session: {self.app_name}")
        
        if self.app_name not in audio_exempt_list:
            # set initial:
            self.update_mute(self.mute)
            self.update_state(self.state)
            self.update_volume(self.volume)
        

    def update_state(self, new_state):
        """
        when status changed
        (see callback -> AudioSessionEvents -> OnStateChanged)
        """
        if self.app_name not in audio_exempt_list:
            if new_state == AudioSessionState.Inactive:
                # AudioSessionStateInactive
                """Sesssion is Inactive"""
                #print(f"{self.app_name} not active")
                TPClient.stateUpdate(f'KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{self.app_name}.active',"False")
    
            elif new_state == AudioSessionState.Active:
                """Session Active"""
                #print(f"{self.app_name} is an Active Session")
                TPClient.stateUpdate(f'KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{self.app_name}.active',"True")
    
        elif new_state == AudioSessionState.Expired:
            """Removing Expired States"""
            removeAudioState(self.app_name)

    
    def update_volume(self, new_volume):
        """
        when volume is changed externally - Updating Sliders and Volume States
        (see callback -> AudioSessionEvents -> OnSimpleVolumeChanged )
        """
        if self.app_name not in audio_exempt_list:
            TPClient.stateUpdate(f'KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{self.app_name}',str(round(new_volume*100)))
            #print(f"{self.app_name} NEW VOLUME", str(round(new_volume*100)))
            TPClient.send({
                "type":"connectorUpdate",
                "connectorId":f"pc_Windows-Tools_KillerBOSS.TP.Plugins.VolumeMixer.connectors.APPcontrol|KillerBOSS.TP.Plugins.VolumeMixer.slidercontrol={self.app_name}",
                "value": round(new_volume*100)
            })
            
            """Checking for Current App If Its Active, Adjust it also"""
            if (activeWindow := getActiveExecutablePath()) != "":
                if os.path.basename(activeWindow) == self.app_name:
                    TPClient.send({
                        "type":"connectorUpdate",
                        "connectorId":f"pc_Windows-Tools_KillerBOSS.TP.Plugins.VolumeMixer.connectors.APPcontrol|KillerBOSS.TP.Plugins.VolumeMixer.slidercontrol=Current app",
                        "value": round(new_volume*100)
                    })
                else:
                    TPClient.send({
                        "type":"connectorUpdate",
                        "connectorId":f"pc_Windows-Tools_KillerBOSS.TP.Plugins.VolumeMixer.connectors.APPcontrol|KillerBOSS.TP.Plugins.VolumeMixer.slidercontrol=Current app",
                        "value": 0
                    })

    def update_mute(self, muted):
        """ when mute state is changed by user or through other app """
        
        if self.app_name not in audio_exempt_list:
            audioStateManager(self.app_name)
            
            if muted:
                #print(f"{self.app_name} is unmuted")
                TPClient.stateUpdate(f"KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{self.app_name}.muteState", "Muted")
            # TPClient.stateUpdate("KillerBOSS.TP.Plugins.state.test", "True")
            else:
                #print(f"{self.app_name} is muted")
                TPClient.stateUpdate(f"KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{self.app_name}.muteState", "Un-Muted")
            #  TPClient.stateUpdate("KillerBOSS.TP.Plugins.state.test", "False")

def stateUpdate():
    counter = 0
    while running:
        sleep(0.1)
        counter += 1
        if counter%2 == 0:
            """
            Update every 2 seconds
            """
            TPClient.stateUpdate("KillerBOSS.TP.Plugins.Application.currentFocusedAPP", pygetwindow.getActiveWindowTitle())

        if counter%5 == 0:
            TPClient.send({
                "type":"connectorUpdate",
                "connectorId":"pc_Windows-Tools_KillerBOSS.TP.Plugins.VolumeMixer.connectors.APPcontrol|KillerBOSS.TP.Plugins.VolumeMixer.slidercontrol=Master Volume",
                "value": str(audioController.getMasterVolume())
                })

        mp = pyautogui.position()  ## Making the call for mouse position once instead of twice.. cause performance :)
        TPClient.stateUpdateMany([
        {
            "id": f'KillerBOSS.TP.Plugins.AdvanceMouse.MousePos.X',
            "value": str(mp[0])
        },
        {
            "id": f'KillerBOSS.TP.Plugins.AdvanceMouse.MousePos.Y',
            "value": str(mp[1])
        }])
    
    if counter >= 40: counter = 0

@TPClient.on(TYPES.onConnect)
def startManager(data):
    global running
    """Checking if Plugin needs updated"""
    github_check = TouchPortalAPI.Tools.updateCheck("KillerBOSS2019", "WinTools")
    plugin_version = str(data['pluginVersion'])
    plugin_version = plugin_version[:1] + "." + plugin_version[1:]
    if github_check[1:4] != plugin_version[0:3]:
        TPClient.showNotification(
                notificationId="KillerBOSS.TP.Plugins.Update_Check",
                title=f"WinTools v{github_check[1:4]} is available",
                msg="A new Wintools Version is available and ready to Download. This may include Bug Fixes and or New Features",
                options= [
                    {
                        "id":"Download Update",
                        "title":"Click here to Update"
                    }
                ])

    running = True

    """Updating Choices for Windows Settings options from util.py"""
    TPClient.choiceUpdate("KillerBOSS.TP.Plugins.winsettings.choice", util.windowsSettings())
    pythoncom.CoInitialize() # This somehow solves OSError: [WinError -2147221008] CoInitialize has not been called
    try:
        MagicManager.magic_session(WinAudioCallBack)
    except NotImplementedError as err:
        print(f"--------- Magic already in session!! ---------\n------{err}------")

    thread1 = threading.Thread(target=stateUpdate())
    thread1.start()

@TPClient.on(TYPES.onSettingUpdate)
def settingHandler(data):
    global audio_exempt_list

    audio_exempt_list = data['values'][6]['Audio State Exemption List']
    if audio_exempt_list == "Enter '.exe' name seperated by a comma for more than 1": audio_exempt_list = []

@TPClient.on(TouchPortalAPI.TYPES.onNotificationOptionClicked)
def notificationEvent(data):
    print(data)
    if data['optionId'] == 'Download Update':
        print("Directing user to download url")
        github_check = TouchPortalAPI.Tools.updateCheck("KillerBOSS2019", "WinTools")
        util.runCommandLine(f"Start https://github.com/KillerBOSS2019/WinTools/releases/tag/{github_check}")

@TPClient.on(TYPES.onAction)
def actionManager(data):

    # Audio action manager
    if data['actionId'] == 'KillerBOSS.TP.Plugins.VolumeMixer.Mute/Unmute':
        if data['data'][0]['value'] is not '':
            audioController.muteAndUnMute(data['data'][0]['value'], data['data'][1]['value'])
    if data['actionId'] == 'KillerBOSS.TP.Plugins.VolumeMixer.Increase/DecreaseVolume':
        audioController.volumeChanger(data['data'][0]['value'], data['data'][1]['value'], data['data'][2]['value'])
    
    if data['actionId'] == 'KillerBOSS.TP.Plugins.VolumeMixer.SetmasterVolume':
        audioController.setMasterVolume(data['data'][0]['value'])

    # clipboard action manager
    if data['actionId'] == "KillerBOSS.TP.Plugins.window.current":
        if data['data'][0]['value'] == "Clipboard":
            current_window = pygetwindow.getActiveWindowTitle()
            clipboard.screenshot_window(capture_type=3, window_title=current_window, clipboard=True)
            
        elif data['data'][0]['value'] == "File":
            current_window = pygetwindow.getActiveWindowTitle()
            afile_name = (data['data'][1]['value']) +"/" +(data['data'][2]['value']) 
            clipboard.screenshot_window(capture_type=3, window_title=current_window, clipboard=False, save_location=afile_name)

    if data['actionId'] == "KillerBOSS.TP.Plugins.screencapture.window.file.wildcard":
        global windows_active
        windows_active = util.get_windows()
        for thing in windows_active:
            if data['data'][0]['value'].lower() in thing.lower():
                print("We found", thing)
                if data['data'][4]['value'] == "Clipboard":
                    clipboard.screenshot_window(capture_type=int(data['data'][1]['value']), window_title=thing, clipboard=True)
                    
                elif data['data'][4]['value'] == "File":
                    print("File stuf")
                    afile_name = (data['data'][2]['value']) +"/" +(data['data'][3]['value']) 
                    clipboard.screenshot_window(capture_type=int(data['data'][1]['value']), window_title=thing, clipboard=False, save_location=afile_name)
                break

    if data['actionId'] == "KillerBOSS.TP.Plugins.screencapture.full.file":   
        if data['data'][1]['value'] == "Clipboard":
            try:
                clipboard.screenshot_monitor(monitor_number=(data['data'][0]['value']), clipboard=True)
            except:
                pass
        elif data['data'][1]['value'] == "File":
            try:
                afile_name = (data['data'][2]['value']) +"/" +(data['data'][3]['value'])    
                clipboard.screenshot_monitor(monitor_number=(data['data'][0]['value']), filename=afile_name, clipboard=False)
            except:
                pass
    
    if data['actionId'] == "KillerBOSS.TP.Plugins.screencapture.window.file":
        if (data['data'][0]['value']):
            if data['data'][4]['value'] == "Clipboard":
                clipboard.screenshot_window(capture_type=int(data['data'][1]['value']), window_title=data['data'][0]['value'], clipboard=True)
            if data['data'][4]['value'] == "File":
                afile_name = (data['data'][2]['value']) +"/" +(data['data'][3]['value'])    
                clipboard.screenshot_window(capture_type=int(data['data'][1]['value']), window_title=data['data'][0]['value'], clipboard=False, save_location=afile_name)

    if data['actionId'] == "KillerBOSS.TP.Plugins.capture.clipboard":
        clipboard.send_to_clipboard("text", data['data'][0]['value'] )
    
    if data['actionId'] == "KillerBOSS.TP.Plugins.capture.clipboard.toValue":
        """Save Clipboard Data to a Custom TP Value"""
        TPClient.stateUpdate(stateId=data['data'][0]['value'], stateValue=str(clipboard.get_clipboard_data()))

    # Mouse action manager
    if data['actionId'] == 'KillerBOSS.TP.Plugins.AdvanceMouse.HoldDownToggle':
        if data['data'][0]['value'] == 'Down':
            pyautogui.mouseDown(button=(data['data'][1]['value']).lower())
        elif data['data'][0]['value'] == "Up":
            pyautogui.mouseUp(button=(data['data'][1]['value']).lower())
    if data['actionId'] == "KillerBOSS.TP.Plugins.AdvanceMouse.teleport":
        util.AdvancedMouseFunction(int(data['data'][0]['value']),int(data['data'][1]['value']), int(data['data'][3]['value']), data['data'][2]['value'])
    if data['actionId'] == "KillerBOSS.TP.Plugins.AdvanceMouse.MouseClick":
        pyautogui.click(clicks=int(data['data'][0]['value']), button=data['data'][2]['value'], interval=float(data['data'][1]['value']))
    if data['actionId'] == 'KillerBOSS.TP.Plugins.AdvanceMouse.Function':
        try:
            pyautogui.move(int(data['data'][0]['value']), int(data['data'][1]['value']))
        except Exception:
            pass

    if data['actionId'] == "KillerBOSS.TP.Plugins.TextToSpeech.speak":
        TTSThread = threading.Thread(target=TextToSpeech, daemon=True, args=(data['data'][1]['value'], data['data'][0]['value'], int(
            data['data'][2]['value']), int(data['data'][3]['value']), data['data'][4]['value']))
        TTSThread.start()

    if data['actionId'] == "KillerBOSS.TP.Plugins.winextra.emojipanel":
        winextra("Emoji")

    if data['actionId'] == "KillerBOSS.TP.Plugins.winextra.keyboard":
        winextra("Keyboard")

    if data['actionId'] == "KillerBOSS.TP.Plugins.winsettings.action":
        windowsSettings(data['data'][0]['value'])

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
                audioController.volumeChanger(data['data'][0]['value'], 'Decrease', data['data'][2]['value'])
                sleep(0.05)
            elif data['data'][1]['value'] == "Increase":
                audioController.volumeChanger(data['data'][0]['value'], 'Increase', data['data'][2]['value'])
                sleep(0.05)
        elif TPClient.isActionBeingHeld('KillerBOSS.TP.Plugins.magnifier.onHold.actions'):
            if data['data'][0]['value'] == "Lens X":
                magnifer_dimensions(x=True, y=None, onhold=int(data['data'][1]['value']))
                sleep(0.05)
            if data['data'][0]['value'] == "Lens Y":
                print("uh hi?")
                magnifer_dimensions(x=False, y=True, onhold=int(data['data'][1]['value']))
                sleep(0.05)
            if data['data'][0]['value'] == "Zoom":
                mag_level(int(data['data'][1]['value']), onhold=True)
                sleep(0.05)
        # elif TPClient.isActionBeingHeld('KillerBOSS.TP.Plugins.winsettings.active.mouseCapture'): # Do we really need this?
        #     if not cap_live:
        #         img=capture_around_mouse(int(data['data'][1]['value']),int(data['data'][2]['value']))
        #         TPClient.stateUpdate(stateId="KillerBOSS.TP.Plugins.winsettings.active.mouseCapture", stateValue=getFrame_base64(img).decode())
        #         sleep(0.07)
        #     else:
        #         print("Attempted PUSH + HOLD CAP, but Capture is already live")
        #         sleep(0.35)
        else:
            break

@TPClient.on(TYPES.onConnectorChange)
def connectors(data):
    if data['connectorId'] == "KillerBOSS.TP.Plugins.VolumeMixer.connectors.APPcontrol":
        if data['data'][0]['value'] == "Master Volume" :
            audioController.setMasterVolume(data['value'])
        elif data['data'][0]['value'] == "Current app":
            activeWindow = getActiveExecutablePath()
          #  print(activeWindow)
            if activeWindow != "":
                audioController.volumeChanger(os.path.basename(activeWindow), "Set", data['value'])
        else:
            try:
                audioController.volumeChanger(data['data'][0]['value'], "Set", data['value'])
            except:
                pass
            
    if data['connectorId'] == "KillerBOSS.TP.Plugins.Magnifier.connectors.ZoomControl":
        if data['data'][0]['value'] == "Zoom" :
             mag_level(data['value']*16)
             
        if data['data'][0]['value'] == "Lens X" :
            magnifer_dimensions(x=data['value'])
        
        if data['data'][0]['value'] == "Lens Y" :
            magnifer_dimensions(y=data['value'])

def updateDeviceOutput(options):
    from utils.util import AudioDeviceCmdlets
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
    global oldcursors
    if data['actionId'] == 'KillerBOSS.TP.Plugins.ChangeAudioOutput':
        try:
            updateDeviceOutput(data['value'])
        except KeyError:
            pass
    if data['actionId'] == 'KillerBOSS.TP.Plugins.virtualdesktop.actions.move_window':
        vd_check()

try:
    TPClient.connect()  # blocking
except KeyboardInterrupt:
    print("Caught keyboard interrupt, exiting.")
except Exception:
    from traceback import format_exc
    print(f"Exception in TP Client:\n{format_exc()}")
finally:
    running = False
    TPClient.disconnect() # make sure it's stopped, no-op if already stopped.
import TouchPortalAPI
from TouchPortalAPI import TYPES
from utils import util
import os
from pycaw.magic import MagicManager, MagicSession   ### These are used in main.py
from pycaw.constants import AudioSessionState
from time import sleep

TPClient = TouchPortalAPI.Client('Windows-Tools')

volumeprocess = ["Master Volume", "Current app"]
audio_exempt_list = []


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
                    "value": ""
                },
                {
                    "id": f'KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{app_name}',
                    "desc": f"{app_name} Volume",
                    "value": ""
                },
                {
                    "id": f'KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{app_name}.active',
                    "desc": f"{app_name} Active",
                    "value": ""
                },
                ])
        volumeprocess.append(app_name)

        """UPDATING CHOICES WITH GLOBALS"""
        updateVolumeMixerChoicelist()
        print("new State Added")

    """ Checking for Exempt Audio"""
    if app_name in audio_exempt_list:
        removeAudioState(app_name)

class WinAudioCallBack(MagicSession):
    def __init__(self):
        super().__init__(volume_callback=self.update_volume,
                         mute_callback=self.update_mute,
                         state_callback=self.update_state)

        # ______________ DISPLAY NAME ______________
        self.app_name = self.magic_root_session.app_exec
        print(f":: new session: {self.app_name}")
        
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
                print(f"{self.app_name} not active")
                TPClient.stateUpdate(f'KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{self.app_name}.active',"False")
    
            elif new_state == AudioSessionState.Active:
                """Session Active"""
                print(f"{self.app_name} is an Active Session")
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
            print("NEW VOLUME", str(round(new_volume*100)))
            TPClient.send({
                "type":"connectorUpdate",
                "connectorId":f"pc_Windows-Tools_KillerBOSS.TP.Plugins.VolumeMixer.connectors.APPcontrol|KillerBOSS.TP.Plugins.VolumeMixer.slidercontrol={self.app_name}",
                "value": round(new_volume*100)
            })
            
            """Checking for Current App If Its Active, Adjust it also"""
            #activeWindow = getActiveExecutablePath()
            activeWindow = ""
            if activeWindow != "":
                if os.path.basename(activeWindow) == self.app_name:
                    TPClient.send({
                        "type":"connectorUpdate",
                        "connectorId":f"pc_Windows-Tools_KillerBOSS.TP.Plugins.VolumeMixer.connectors.APPcontrol|KillerBOSS.TP.Plugins.VolumeMixer.slidercontrol=Current app",
                        "value": round(new_volume*100)
                    })

    def update_mute(self, muted):
        """ when mute state is changed by user or through other app """
        
        if self.app_name not in audio_exempt_list:
            audioStateManager(self.app_name)
            
            if muted:
                print(f"{self.app_name} is unmuted")
                TPClient.stateUpdate(f"KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{self.app_name}.muteState", "Muted")
            # TPClient.stateUpdate("KillerBOSS.TP.Plugins.state.test", "True")
            else:
                print(f"{self.app_name} is muted")
                TPClient.stateUpdate(f"KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{self.app_name}.muteState", "Un-Muted")
            #  TPClient.stateUpdate("KillerBOSS.TP.Plugins.state.test", "False")

@TPClient.on(TYPES.onConnect)
def startManager(data):
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

    """Updating Choices for Windows Settings options from util.py"""
    print(util.windowsSettings())
    TPClient.choiceUpdate("KillerBOSS.TP.Plugins.winsettings.choice", util.windowsSettings())
    MagicManager.magic_session(WinAudioCallBack)


@TPClient.on(TYPES.onSettingUpdate)
def settingHandler(data):
    pass
    

TPClient.connect()
import os
from time import sleep

from constants import TPClient
from pycaw.constants import AudioSessionState
from pycaw.magic import MagicSession

from .WinAudioCallbackHelper import check_states, getActiveExecutablePath


class WinAudioCallBack(MagicSession):
    globalApp = ["Master Volume", "Current app"]

    def __init__(self, audio_exempt_list):
        super().__init__(volume_callback=self.update_volume,
                         mute_callback=self.update_mute,
                         state_callback=self.update_state)

        # ______________ DISPLAY NAME ______________
        self.app_name = self.magic_root_session.app_exec
        self.audio_exempt_list = audio_exempt_list
        print(f":: new session: {self.app_name}")

        if self.app_name not in self.audio_exempt_list:
            # set initial:
            self.update_mute(self.mute)
            self.update_state(self.state)
            self.update_volume(self.volume)
        

    def update_state(self, new_state):
        """
        when status changed
        (see callback -> AudioSessionEvents -> OnStateChanged)
        """
        if self.app_name not in self.audio_exempt_list:
            if new_state == AudioSessionState.Inactive:
                # AudioSessionStateInactive
                """Sesssion is Inactive"""
                print(f"{self.app_name} not active")
                TPClient.stateUpdate(f'KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{self.app_name}.active',"False")
    
            elif new_state == AudioSessionState.Active:
                """Session Active"""
                print(f"{self.app_name} is an Active Session")
                TPClient.stateUpdate(f'KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{self.app_name}.active',"True")
        if new_state == AudioSessionState.Expired:
            """Removing Expired States"""
            self.globalApp = check_states(self.app_name, self.globalApp, remove=True)
    
    def update_volume(self, new_volume):
        """
        when volume is changed externally - Updating Sliders and Volume States
        (see callback -> AudioSessionEvents -> OnSimpleVolumeChanged )
        """
        if self.app_name not in self.audio_exempt_list:
            TPClient.stateUpdate(f'KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{self.app_name}',str(round(new_volume*100)))
            print("NEW VOLUME", str(round(new_volume*100)))
            TPClient.send({
                "type":"connectorUpdate",
                "connectorId":f"pc_Windows-Tools_KillerBOSS.TP.Plugins.VolumeMixer.connectors.APPcontrol|KillerBOSS.TP.Plugins.VolumeMixer.slidercontrol={self.app_name}",
                "value": round(new_volume*100)
            })
            
            """Checking for Current App If Its Active, Adjust it also"""
            activeWindow = getActiveExecutablePath()
            if activeWindow != "":
                if os.path.basename(activeWindow) == self.app_name:
                    TPClient.send({
                        "type":"connectorUpdate",
                        "connectorId":f"pc_Windows-Tools_KillerBOSS.TP.Plugins.VolumeMixer.connectors.APPcontrol|KillerBOSS.TP.Plugins.VolumeMixer.slidercontrol=Current app",
                        "value": round(new_volume*100)
                    })            
        
    def update_mute(self, muted):
        """ when mute state is changed by user or through other app """
        
        if self.app_name not in self.audio_exempt_list:
            self.globalApp = check_states(self.app_name, self.globalApp)
            sleep(0.1)
            
            if muted:
                print(f"{self.app_name} is unmuted")
                TPClient.stateUpdate(f"KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{self.app_name}.muteState", "Muted")
            # TPClient.stateUpdate("KillerBOSS.TP.Plugins.state.test", "True")
            else:
                print(f"{self.app_name} is muted")
                TPClient.stateUpdate(f"KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{self.app_name}.muteState", "Un-Muted")
            #  TPClient.stateUpdate("KillerBOSS.TP.Plugins.state.test", "False")

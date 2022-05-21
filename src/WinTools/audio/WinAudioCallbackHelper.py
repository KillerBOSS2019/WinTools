import ctypes

import psutil
import win32process
from constants import TPClient


def check_states(app_name, volumeStates, remove=False):
    newVolumeState = volumeStates
    if remove:
        TPClient.removeStateMany([
            f'KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{app_name}.muteState',
            f'KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{app_name}',
            f'KillerBOSS.TP.Plugins.VolumeMixer.CreateState.{app_name}.active'
        ])
        if app_name in newVolumeState:
            newVolumeState.remove(app_name)
    elif app_name not in newVolumeState:
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
        newVolumeState.append(app_name)



        """UPDATING CHOICES WITH GLOBALS"""
        TPClient.choiceUpdate('KillerBOSS.TP.Plugins.VolumeMixer.Increase/DecreaseVolume.process', newVolumeState)
        TPClient.choiceUpdate('KillerBOSS.TP.Plugins.VolumeMixer.Mute/Unmute.process', newVolumeState)
        TPClient.choiceUpdate("KillerBOSS.TP.Plugins.VolumeMixer.slidercontrol", newVolumeState)
    return newVolumeState

def getActiveExecutablePath():
    hWnd = ctypes.windll.user32.GetForegroundWindow()
    if hWnd == 0:
        return None # Note that this function doesn't use GetLastError().
    else:
        _, pid = win32process.GetWindowThreadProcessId(hWnd)
        return psutil.Process(pid).exe()

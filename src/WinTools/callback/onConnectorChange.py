import os

from audio.mainVolume import setMasterVolume, volumeChanger
from audio.WinAudioCallbackHelper import getActiveExecutablePath
from utils.magnifier import mag_level, magnifer_dimensions


def onConnectorChange(data):
    if data['connectorId'] == "KillerBOSS.TP.Plugins.VolumeMixer.connectors.APPcontrol":
        if data['data'][0]['value'] == "Master Volume" :
            setMasterVolume(data['value'])
        elif data['data'][0]['value'] == "Current app":
            activeWindow = getActiveExecutablePath()
          #  print(activeWindow)
            if activeWindow != "":
                volumeChanger(os.path.basename(activeWindow), "Set", data['value'])
        else:
            try:
                volumeChanger(data['data'][0]['value'], "Set", data['value'])
            except:
                pass
            
    if data['connectorId'] == "KillerBOSS.TP.Plugins.Magnifier.connectors.ZoomControl":
        if data['data'][0]['value'] == "Zoom" :
             mag_level(data['value']*16)
             
        if data['data'][0]['value'] == "Lens X" :
            magnifer_dimensions(x=data['value'])
        
        if data['data'][0]['value'] == "Lens Y" :
            magnifer_dimensions(y=data['value'])
            
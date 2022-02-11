
from ctypes import POINTER, cast, windll
import threading
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import time

def getMasterVolume2():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    print(int(round(volume.GetMasterVolumeLevelScalar() * 100)))
   # return int(round(volume.GetMasterVolumeLevelScalar() * 100))



def get_master_volume_current():
    while True:
        print("HELLO THERE ARE YOU LISTENING TO ME", "--"*20)
        master_volume = str(getMasterVolume2())
        #PClient.send(
        #           {
        #               "type":"connectorUpdate",
        #               "connectorId":"pc_Windows-Tools_KillerBOSS.TP.Plugins.VolumeMixer.connectors.APPcontrol|KillerBOSS.TP.Plugins.VolumeMixer.slidercontrol=Master Volume",
        #               "value": master_volume
        #           }
        #       )
        print(master_volume)
        time.sleep(1)
    
th = threading.Thread(target=get_master_volume_current)
#th.start()


getMasterVolume2()
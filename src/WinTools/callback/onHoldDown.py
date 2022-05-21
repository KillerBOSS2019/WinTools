from time import sleep
from constants import TPClient
import pyautogui
from audio.mainVolume import volumeChanger
from utils.magnifier import (magnifer_dimensions, mag_level)
from utils.mouseCapture import getFrame_base64, capture_around_mouse

def onHoldDown(data):
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
        elif TPClient.isActionBeingHeld('KillerBOSS.TP.Plugins.winsettings.active.mouseCapture'):
            pass # Need help figure out how to fix this `cap_live`
            #if not cap_live:
            #    img=capture_around_mouse(int(data['data'][1]['value']),int(data['data'][2]['value']))
            #    TPClient.stateUpdate(stateId="KillerBOSS.TP.Plugins.winsettings.active.mouseCapture", stateValue=getFrame_base64(img).decode())
            #    sleep(0.07)
            #else:
            #    print("Attempted PUSH + HOLD CAP, but Capture is already live")
            #    sleep(0.35)
        else:
            break
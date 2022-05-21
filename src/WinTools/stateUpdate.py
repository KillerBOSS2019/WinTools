from threading import Thread
from time import sleep

import pyautogui
import pygetwindow

from constants import Tools, TPClient


def stateUpdate():
    """
    Update anything
    """
    sleep(10) # wait for TP to connect
    while TPClient.isConnected():
        try:
            TPClient.stateUpdate("KillerBOSS.TP.Plugins.Application.currentFocusedAPP", pygetwindow.getActiveWindowTitle())
        except:
            pass

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
        sleep(0.1)

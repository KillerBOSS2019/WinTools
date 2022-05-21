import os
import threading

import pyautogui
import pygetwindow
import sounddevice as sd
from audio.mainVolume import (AudioDeviceCmdlets, muteAndUnMute,
                              setMasterVolume, volumeChanger)
from constants import TPClient
from PIL import Image
from util import get_windows, move_win_button, resize_image, winextra
from utils.checkProcess import check_process
from utils.clipboard import get_clipboard_data, send_to_clipboard
from utils.magnifier import mag_level, magnifier
from utils.monitorPrimary import change_primary
from utils.mouseCapture import turn_on_cap
from utils.powerplan import change_pplan, get_powerplans
from utils.rotateDisplay import rotate_display
from utils.screenshotFunc import screenshot_monitor, screenshot_window
from utils.TextToSpeech import TextToSpeech
from utils.virtualDesktop import (create_vd, remove_vd, rename_vd,
                                  virtual_desktop)
from utils.winNotification import win_toast
from utils.winShutdown import win_shutdown

from callback.onStartprepare import activate_windows_setting


def AdvancedMouseFunction(x, y, delay, look):
    if look == 0:
        look = None
    if look == "None":
        try:
            pyautogui.moveTo(x, y, delay)
        except:
            pass
    elif look == "Start slow, end fast":
        try:
            pyautogui.moveTo(x, y, delay, pyautogui.easeInQuad)
        except:
            pass
    elif look == "Start fast, end slow":
        try:
            pyautogui.moveTo(x, y, delay, pyautogui.easeOutQuad)
        except:
            pass
    elif look == "Start and end fast, slow in middle":
        try:
            pyautogui.moveTo(x, y, delay, pyautogui.easeInOutQuad)
        except:
            pass
    elif look == "bounce at the end":
        try:
            pyautogui.moveTo(x, y, delay, pyautogui.easeInBounce)
        except:
            pass
    elif look == "rubber band at the end":
        try:
            pyautogui.moveTo(x, y, delay, pyautogui.easeInElastic)
        except:
            pass


def onAction(data):
    print(data)
    if data['actionId'] == 'KillerBOSS.TP.Plugins.VolumeMixer.Mute/Unmute':
        if data['data'][0]['value'] is not '':
            muteAndUnMute(data['data'][0]['value'], data['data'][1]['value'])
    if data['actionId'] == 'KillerBOSS.TP.Plugins.VolumeMixer.Increase/DecreaseVolume':
        volumeChanger(data['data'][0]['value'], data['data']
                      [1]['value'], data['data'][2]['value'])

    if data['actionId'] == 'KillerBOSS.TP.Plugins.VolumeMixer.SetmasterVolume':
        setMasterVolume(data['data'][0]['value'])

    if data['actionId'] == 'KillerBOSS.TP.Plugins.AdvanceMouse.HoldDownToggle':
        if data['data'][0]['value'] == 'Down':
            pyautogui.mouseDown(button=(data['data'][1]['value']).lower())
        elif data['data'][0]['value'] == "Up":
            pyautogui.mouseUp(button=(data['data'][1]['value']).lower())
    if data['actionId'] == "KillerBOSS.TP.Plugins.AdvanceMouse.teleport":
        AdvancedMouseFunction(int(data['data'][0]['value']), int(
            data['data'][1]['value']), int(data['data'][3]['value']), data['data'][2]['value'])
    if data['actionId'] == "KillerBOSS.TP.Plugins.AdvanceMouse.MouseClick":
        pyautogui.click(clicks=int(data['data'][0]['value']), button=(
            data['data'][2]['value']).lower(), interval=float(data['data'][1]['value']))
    if data['actionId'] == 'KillerBOSS.TP.Plugins.AdvanceMouse.Function':
        try:
            pyautogui.move(int(data['data'][0]['value']),
                           int(data['data'][1]['value']))
        except Exception:
            pass
    if data['actionId'] == 'KillerBOSS.TP.Plugins.ChangeAudioOutput' and data['data'][0]['value'] != "Pick One":
        print('powershell is running')
        AudioDeviceCmdlets(
            f"(Get-AudioDevice -list | Where-Object Name -like (\'{data['data'][1]['value']}') | Set-AudioDevice).Name", output=False)

# Screencap Window Current
    if data['actionId'] == "KillerBOSS.TP.Plugins.window.current":
        if data['data'][0]['value'] == "Clipboard":
            current_window = pygetwindow.getActiveWindowTitle()
            screenshot_window(
                capture_type=3, window_title=current_window, clipboard=True)

        elif data['data'][0]['value'] == "File":
            current_window = pygetwindow.getActiveWindowTitle()
            afile_name = (data['data'][1]['value']) + \
                "/" + (data['data'][2]['value'])
            screenshot_window(capture_type=3, window_title=current_window,
                              clipboard=False, save_location=afile_name)

# Screen Cap Window Wildcard to FILE
    if data['actionId'] == "KillerBOSS.TP.Plugins.screencapture.window.file.wildcard":
        global windows_active
        windows_active = get_windows()
        for thing in windows_active:
            if data['data'][0]['value'].lower() in thing.lower():
                print("We found", thing)
                if data['data'][4]['value'] == "Clipboard":
                    screenshot_window(capture_type=int(
                        data['data'][1]['value']), window_title=thing, clipboard=True)

                elif data['data'][4]['value'] == "File":
                    print("File stuf")
                    afile_name = (data['data'][2]['value']) + \
                        "/" + (data['data'][3]['value'])
                    screenshot_window(capture_type=int(
                        data['data'][1]['value']), window_title=thing, clipboard=False, save_location=afile_name)
                break

    if data['actionId'] == "KillerBOSS.TP.Plugins.screencapture.full.file":
        if data['data'][1]['value'] == "Clipboard":
            try:
                screenshot_monitor(monitor_number=(
                    data['data'][0]['value']), clipboard=True)
            except:
                pass
        elif data['data'][1]['value'] == "File":
            try:
                afile_name = (data['data'][2]['value']) + \
                    "/" + (data['data'][3]['value'])
                screenshot_monitor(monitor_number=(
                    data['data'][0]['value']), filename=afile_name, clipboard=False)
            except:
                pass

    if data['actionId'] == "KillerBOSS.TP.Plugins.screencapture.window.file":
        if (data['data'][0]['value']):
            if data['data'][4]['value'] == "Clipboard":
                screenshot_window(capture_type=int(
                    data['data'][1]['value']), window_title=data['data'][0]['value'], clipboard=True)
            if data['data'][4]['value'] == "File":
                afile_name = (data['data'][2]['value']) + \
                    "/" + (data['data'][3]['value'])
                screenshot_window(capture_type=int(
                    data['data'][1]['value']), window_title=data['data'][0]['value'], clipboard=False, save_location=afile_name)

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

        check_process(app_to_check, shortcut_to_open,
                      focus=focus_check, focus_type=afocus_type)

    if data['actionId'] == "KillerBOSS.TP.Plugins.capture.clipboard":
        send_to_clipboard("text", data['data'][0]['value'])

    if data['actionId'] == "KillerBOSS.TP.Plugins.capture.clipboard.toValue":
        """Save Clipboard Data to a Custom TP Value"""
        TPClient.stateUpdate(
            stateId=data['data'][0]['value'], stateValue=str(get_clipboard_data()))

    if data['actionId'] == "KillerBOSS.TP.Plugins.virtualdesktop.actions":
        choice = data['data'][0]['value']
        virtual_desktop(target_desktop=choice)

    if data['actionId'] == "KillerBOSS.TP.Plugins.winsettings.rotate_display":
        if data['data'][0]['value'] != "Pick a Monitor":
            rotate_display(int(data['data'][0]['value']),
                           data['data'][1]['value'])

    if data['actionId'] == "KillerBOSS.TP.Plugins.winsettings.shutdown":
        win_shutdown(data['data'][0]['value'])

    if data['actionId'] == "KillerBOSS.TP.Plugins.winsettings.primary_monitor":
        change_primary(data['data'][0]['value'])

    if data['actionId'] == "KillerBOSS.TP.Plugins.winsettings.powerplan":
        change_pplan(data['data'][0]['value'])
        TPClient.stateUpdate(
            "KillerBOSS.TP.Plugins.winsettings.powerplan_current", get_powerplans()[1])

    if data['actionId'] == "KillerBOSS.TP.Plugins.winsettings.move_window":
        move_win_button(data['data'][0]['value'])

    if data['actionId'] == "KillerBOSS.TP.Plugins.virtualdesktop.actions.move_window":
        choice = data['data'][0]['value']
        if data['data'][1]['value'] == "False":
            virtual_desktop(move=True, target_desktop=choice)
        if data['data'][1]['value'] == "True":
            virtual_desktop(move=True, target_desktop=choice, pinned=True)

    if data['actionId'] == 'KillerBOSS.TP.Plugins.virtualdesktop.create':
        create_vd(data['data'][0]['value'])
    if data['actionId'] == 'KillerBOSS.TP.Plugins.virtualdesktop.remove':
        remove_vd(data['data'][0]['value'])
    if data['actionId'] == 'KillerBOSS.TP.Plugins.virtualdesktop.rename':
        rename_vd(name=data['data'][0]['value'],
                  number=data['data'][1]['value'])

    if data['actionId'] == "KillerBOSS.TP.Plugins.magnifier.actions":
        magnifier(data['data'][0]['value'])

    if data['actionId'] == "KillerBOSS.TP.Plugins.magnifier.onHold.actions":
        if data['data'][0]['value'] == "Zoom":
            mag_level(int(data['data'][1]['value']))

    if data['actionId'] == "KillerBOSS.TP.Plugins.toast.create":
        win_toast(atitle=data['data'][0]['value'], amsg=data['data'][1]['value'], aduration=data['data'][2]['value'], icon=data['data']
                  [5]['value'], buttonText=data['data'][3]['value'], buttonlink=data['data'][4]['value'], sound=data['data'][6]['value'])

    if data['actionId'] == "KillerBOSS.TP.Plugins.winextra.emojipanel":
        winextra("Emoji")

    if data['actionId'] == "KillerBOSS.TP.Plugins.winextra.keyboard":
        winextra("Keyboard")

    if data['actionId'] == "KillerBOSS.TP.Plugins.winsettings.action":
        activate_windows_setting(data['data'][0]['value'])

    if data['actionId'] == "KillerBOSS.TP.Plugins.TextToSpeech.speak":
        TTSThread = threading.Thread(target=TextToSpeech, daemon=True, args=(data['data'][1]['value'], data['data'][0]['value'], int(
            data['data'][2]['value']), int(data['data'][3]['value']), data['data'][4]['value']))
        TTSThread.start()

    if data['actionId'] == "KillerBOSS.TP.Plugins.TextToSpeech.stop":
        if TTSThread.is_alive():
            sd.stop()
            pass

    if data['actionId'] == "KillerBOSS.TP.Plugins.winsettings.active.mouseCapture":
        if data['data'][0]['value'] == "ON":
            global cap_live
            if cap_live:
                print("DENIED, ONLY ONE ALLOWED")
            elif not cap_live:
                if data['data'][3]['value'] == "None":
                    print("NO OVERLAY WANTED")
                    th1 = threading.Thread(target=turn_on_cap, daemon=True, args=(
                        int(data['data'][1]['value']), int(data['data'][2]['value']), None))
                    cap_live = True
                    th1.start()

                if data['data'][3]['value']:
                    if ".png" in data['data'][3]['value']:
                        print("*"*20, "OVERLAY REQUESTED", "*"*20)
                        path = os.getcwd()
                        path = path + f"\mouse_overlays\\" + \
                            data['data'][3]['value']
                        print(path)
                        img = Image.open(path)
                        if img.size[0] == int(data['data'][1]['value']) and img.size[1] == int(data['data'][2]['value']):
                            print("sized properly")
                            th1 = threading.Thread(target=turn_on_cap, daemon=True, args=(
                                int(data['data'][1]['value']), int(data['data'][2]['value']), img))
                            cap_live = True
                            th1.start()

                        elif img.size[0] != int(data['data'][1]['value']):
                            print("NOT EQUAL WE AHVE TO RESIZE IT")
                            print(int(data['data'][1]['value']),
                                  int(data['data'][2]['value']))
                            resized = resize_image(
                                img, int(data['data'][1]['value']), int(data['data'][2]['value']))
                            th1 = threading.Thread(target=turn_on_cap, daemon=True,  args=(
                                int(data['data'][1]['value']), int(data['data'][2]['value']), resized))
                            cap_live = True
                            th1.start()

        elif data['data'][0]['value'] == "OFF":
            cap_live = False
            global if_running
            if_running = False
            pass

import json
import os
import re
import sys
import threading
import time
from subprocess import PIPE, Popen

import keyboard
import mouse
from TouchPortalAPI.logger import Logger
from TPPEntry import PLUGIN_ID

g_log = Logger(name=PLUGIN_ID)


class States:
    macro_recordState = False
    macroRecordThread = None

    macro_playState = False
    macroPlayThread = None


def record(name, macro_action, file='record.json'):
  #  print(name, macro_action)
    mouse_events = []
    keyboard_events = []

    mouse.hook(mouse_events.append)
    keyboard.start_recording()
    starttime = time.time()

    keyboard.wait('esc')
    keyboard_events = keyboard.stop_recording()
    mouse.unhook(mouse_events.append)

    with open(file, "r") as currentFile:
        try:
            currentFile = json.loads(currentFile.read())
        except json.JSONDecodeError:
            currentFile = {}

    with open(file, "w") as f:
        currentFile[name] = {
            "startTime": starttime,
            "KeyboardEvent": [keyboard_events[kevent].to_json() for kevent in range(0, len(keyboard_events))],
            "MouseEvent": str(mouse_events),
        }
        json.dump(currentFile, f, indent=2)


def play(name, file="record.json", speed=1):
    with open(file, "r") as f:
        replayFile = json.load(f)

    if name in replayFile:
        replayFile = replayFile[name]
    else:
        return -1

    keyboard_Event = []
    for index in range(len(replayFile["KeyboardEvent"])):
        kevent = json.loads(replayFile["KeyboardEvent"][index])
        keyboard_Event.append(keyboard.KeyboardEvent(**kevent))

    mouse_event = eval(replayFile["MouseEvent"])

    starttime = float(replayFile['startTime'])

    keyboard_time_interval = keyboard_Event[0].time - starttime

    mouse_time_interval = mouse_event[0].time - starttime

    # Keyboard threadings:
    k_thread = threading.Thread(target=lambda: time.sleep(
        keyboard_time_interval) == keyboard.play(keyboard_Event, speed_factor=speed))
    # Mouse threadings:
    m_thread = threading.Thread(target=lambda: time.sleep(
        mouse_time_interval) == mouse.play(mouse_event, speed_factor=speed))

    # start threads
    k_thread.start()
    m_thread.start()
    # waiting for both threadings to be completed
    k_thread.join()
    m_thread.join()


def getMacroProfile(file="./record.json"):
    replayFile = {}
    try:
        with open(file, "r") as f:
            replayFile = json.load(f)
    except:
        with open(file, "x") as f:
            f.write("{}")
    return replayFile


# record("hm")
# play("hm", speed=1)


def get_active_window_title():
    """ 
    For Linux 
    """

    root = Popen(['xprop', '-root', '_NET_ACTIVE_WINDOW'], stdout=PIPE)
    stdout, stderr = root.communicate()

    m = re.search(b'^_NET_ACTIVE_WINDOW.* ([\w]+)$', stdout)

    if m is not None:
        window_id = m.group(1)
        window = Popen(['xprop', '-id', window_id, 'WM_NAME'], stdout=PIPE)
        stdout, stderr = window.communicate()

        match = re.match(b'WM_NAME\(\w+\) = (?P<name>.+)$', stdout)
        if match is not None:
            return match.group('name').decode('UTF-8').strip('"')

    return 'Active window not found'


if __name__ == '__main__':
    pass
   # print( get_active_window_title() )

# screenshot_current_linux()


# https://www.wikitechy.com/tutorials/linux/how-to-take-a-screenshot-via-a-python-script-linux

# https://askubuntu.com/questions/1011507/screenshot-of-an-active-application-using-python

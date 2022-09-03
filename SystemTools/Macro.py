import threading
import mouse
import keyboard
import time

from mouse import ButtonEvent
from mouse import MoveEvent
from mouse import WheelEvent

from keyboard import KeyboardEvent

import json

class States:
    macro_recordState = False
    macroRecordThread = None

    macro_playState = False
    macroPlayThread = None


def record(name, file='record.json'):
    mouse_events = []
    keyboard_events = []

    keyboard.start_recording()
    starttime = time.time()

    mouse.hook(mouse_events.append)
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
    keyboard_time_interval /= speed

    mouse_time_interval = mouse_event[0].time - starttime
    mouse_time_interval /= speed

    #Keyboard threadings:
    k_thread = threading.Thread(target = lambda : time.sleep(keyboard_time_interval) == keyboard.play(keyboard_Event, speed_factor=speed) )
    #Mouse threadings:
    m_thread = threading.Thread(target = lambda : time.sleep(mouse_time_interval) == mouse.play(mouse_event, speed_factor=speed))
    
    #start threads
    k_thread.start()
    m_thread.start()
    #waiting for both threadings to be completed
    print(m_thread.is_alive())
    k_thread.join() 
    m_thread.join()

def getMacroProfile(file="record.json"):
    with open(file, "r") as f:
        replayFile = json.load(f)
    return replayFile


# record("hm")
# play("hm", speed=1)























